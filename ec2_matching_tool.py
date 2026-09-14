#!/usr/bin/env python3
"""
Tool to match RVTools export data with EC2 instance types
and generate output in Amazon EC2 Bulk Upload template format.
"""

import pandas as pd
import numpy as np
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font
import sys
import argparse
import os

def read_rvtools_data(filepath):
    """Read RVTools export Excel file - only Vinfo tab with specific columns."""
    print(f"Reading RVTools data from {filepath}...")
    df = pd.read_excel(filepath, sheet_name='vInfo')
    
    # Define required columns and their alternatives
    column_requirements = {
        'VM': ['VM', 'Name'],
        'CPUs': ['CPUs', 'CPU', 'vCPU'],
        'Active Memory': ['Active Memory'],  # Don't use In Use MiB as it may contain disk usage
        'Memory': ['Memory', 'Total Memory'],
        'Total disk capacity MiB': ['In Use MiB', 'Provisioned MiB', 'Total disk capacity MiB', 'Disk Capacity', 'Disk'],
        'OS according to the configuration file': ['OS according to the configuration file', 'OS', 'Operating System']
    }
    
    # Map available columns to required names
    final_columns = {}
    for required, alternatives in column_requirements.items():
        found = False
        for alt in alternatives:
            if alt in df.columns:
                final_columns[required] = alt
                found = True
                break
        if not found:
            print(f"Warning: Column '{required}' not found (tried: {alternatives})")
            # Use a default value for missing columns
            if required == 'Active Memory':
                pass  # Will fall back to Memory in process_matches
            elif required == 'OS according to the configuration file':
                # Add a default OS column
                df['OS according to the configuration file'] = 'Linux'
                final_columns[required] = 'OS according to the configuration file'
            else:
                raise KeyError(f"Required column '{required}' not found and no default available")
    
    # Select and rename columns
    columns_to_select = [final_columns[col] for col in column_requirements.keys() if col in final_columns]
    selected_df = df[columns_to_select].copy()
    selected_df = selected_df.rename(columns={v: k for k, v in final_columns.items()})
    
    # Reset index to avoid comparison issues
    selected_df = selected_df.reset_index(drop=True)
    
    print(f"Found {len(selected_df)} VMs in RVTools vInfo tab")
    return selected_df

def read_ec2_instances(filepath):
    """Read EC2 instance list from CloudPrice_Reference.xlsx (Linux tab as default)."""
    print(f"Reading EC2 instance list from {filepath} (Linux tab)...")
    df = pd.read_excel(filepath, sheet_name='Linux')
    
    # Clean column names
    df.columns = df.columns.str.strip()
    
    # Rename columns to match expected names
    df = df.rename(columns={
        'ProcessorVCPUCount': 'CPU',
        'MemorySizeInGB': 'Instance Memory',
        'InstanceType': 'API Name'
    })
    
    # Reset index to avoid comparison issues
    df = df.reset_index(drop=True)
    
    print(f"Found {len(df)} EC2 instance types")
    return df

def read_price_reference(filepath, os_type='Linux'):
    """Read CloudPrice reference Excel file for specific OS type."""
    sheet_name = 'windows' if os_type == 'Windows Server' else 'Linux'
    print(f"Reading CloudPrice reference from {filepath} (sheet: {sheet_name})...")
    df = pd.read_excel(filepath, sheet_name=sheet_name)
    # Clean column names
    df.columns = df.columns.str.strip()
    print(f"Found {len(df)} priced instance types for {os_type}")
    return df

def read_all_price_sheets(filepath):
    """Read both Linux and Windows price sheets from CloudPrice reference."""
    print(f"Reading all price sheets from {filepath}...")
    linux_prices = pd.read_excel(filepath, sheet_name='Linux')
    windows_prices = pd.read_excel(filepath, sheet_name='windows')
    
    # Clean column names
    linux_prices.columns = linux_prices.columns.str.strip()
    windows_prices.columns = windows_prices.columns.str.strip()
    
    # Reset index to avoid comparison issues
    linux_prices = linux_prices.reset_index(drop=True)
    windows_prices = windows_prices.reset_index(drop=True)
    
    print(f"Found {len(linux_prices)} Linux and {len(windows_prices)} Windows instance types")
    return {'Linux': linux_prices, 'Windows Server': windows_prices}

def match_vm_to_ec2(vm_cpus, vm_memory_mb, ec2_df, price_df):
    """
    Match a VM to the best EC2 instance type based on CPU, Memory, and Price.
    Returns the API Name of the best matching instance type.
    
    Strategy: Find EC2 instances that meet or exceed the VM's CPU and Memory
    requirements, then select the cheapest one.
    """
    # Ensure inputs are scalar values
    vm_cpus = float(vm_cpus)
    vm_memory_mb = float(vm_memory_mb)
    vm_memory_gb = vm_memory_mb / 1024  # Convert MB to GB
    
    # Filter EC2 instances that meet or exceed VM requirements
    matching = ec2_df[
        (ec2_df['CPU'].values >= vm_cpus) & 
        (ec2_df['Instance Memory'].values >= vm_memory_gb)
    ].copy()
    
    if len(matching) == 0:
        print(f"  Warning: No EC2 instance found for CPU={vm_cpus}, Memory={vm_memory_gb:.2f}GB")
        return None
    
    # Merge with price data to get pricing information
    # Match on API Name from ec2_df to InstanceType in price_df
    matching = matching.merge(
        price_df[['InstanceType', 'PricePerHour']],
        left_on='API Name',
        right_on='InstanceType',
        how='left',
        suffixes=('', '_price')
    )
    
    # Use the price column from the merge (avoiding duplicate column names)
    if 'PricePerHour_price' in matching.columns:
        matching['PricePerHour'] = matching['PricePerHour_price']
        matching = matching.drop(columns=['PricePerHour_price'])
    
    # Filter out instances without pricing data
    matching_with_price = matching[matching['PricePerHour'].notna()].copy()
    
    if len(matching_with_price) == 0:
        print(f"  Warning: No pricing data available for matching instances, falling back to resource-based selection")
        # Fallback to original logic
        matching['total_resources'] = matching['CPU'] + matching['Instance Memory']
        matching = matching.sort_values('total_resources', ascending=True)
        best_match = matching.iloc[0]
        return best_match['API Name']
    
    # Sort by price (ascending) to find cheapest
    matching_with_price = matching_with_price.sort_values('PricePerHour', ascending=True)
    
    best_match = matching_with_price.iloc[0]
    return best_match['API Name']

def process_matches(rvtools_df, ec2_df, price_sheets):
    """Process all VMs and match to EC2 instance types with price consideration."""
    print("\nMatching VMs to EC2 instance types with price optimization...")
    
    results = []
    unmatched = []
    
    for idx in range(len(rvtools_df)):
        row = rvtools_df.iloc[idx]
        vm_name = row['VM']
        vm_cpus = int(row['CPUs'])
        vm_total_memory_mb = float(row['Memory'])
        vm_disk_mib = float(row['Total disk capacity MiB']) if 'Total disk capacity MiB' in rvtools_df.columns else 0
        vm_os = row['OS according to the configuration file'] if 'OS according to the configuration file' in rvtools_df.columns else 'Linux'
        
        # Use Active Memory if > 0 and less than total Memory (VM running), otherwise use total Memory (VM not running)
        active_memory = float(row['Active Memory']) if 'Active Memory' in rvtools_df.columns and pd.notna(row['Active Memory']) else None
        if active_memory is not None and active_memory > 0 and active_memory < vm_total_memory_mb:
            vm_memory_mb = active_memory
            memory_source = 'Active Memory'
        else:
            vm_memory_mb = vm_total_memory_mb
            memory_source = 'Total Memory'
        
        # Map VM OS to EC2 OS type for pricing
        ec2_os = map_os_to_ec2(vm_os)
        
        # Get appropriate price sheet based on OS
        price_df = price_sheets.get(ec2_os, price_sheets['Linux'])
        
        ec2_type = match_vm_to_ec2(vm_cpus, vm_memory_mb, ec2_df, price_df)
        
        # Get price for the selected instance
        price_info = price_df[price_df['InstanceType'] == ec2_type]['PricePerHour']
        price_per_hour = price_info.values[0] if len(price_info) > 0 else None
        
        if ec2_type:
            results.append({
                'VM': vm_name,
                'CPUs': vm_cpus,
                'Memory_MB': vm_memory_mb,
                'Memory_GB': vm_memory_mb / 1024,
                'Disk_MiB': vm_disk_mib,
                'Disk_GB': vm_disk_mib / 1024,
                'OS': vm_os,
                'EC2_OS': ec2_os,
                'Memory_Source': memory_source,
                'EC2_Instance_Type': ec2_type,
                'Price_Per_Hour': price_per_hour
            })
        else:
            unmatched.append({
                'VM': vm_name,
                'CPUs': vm_cpus,
                'Memory_MB': vm_memory_mb,
                'Memory_GB': vm_memory_mb / 1024,
                'Disk_MiB': vm_disk_mib,
                'Disk_GB': vm_disk_mib / 1024,
                'OS': vm_os,
                'EC2_OS': ec2_os,
                'Memory_Source': memory_source
            })
    
    print(f"Successfully matched {len(results)} VMs")
    if unmatched:
        print(f"Could not match {len(unmatched)} VMs")
        for vm in unmatched:
            print(f"  - {vm['VM']}: CPU={vm['CPUs']}, Memory={vm['Memory_GB']:.2f}GB")
    
    return pd.DataFrame(results), pd.DataFrame(unmatched)

def map_os_to_ec2(vm_os):
    """Map RVTools OS to EC2 operating system type."""
    if pd.isna(vm_os):
        return 'Linux'
    
    vm_os_lower = str(vm_os).lower()
    
    # Windows OS mapping
    if any(win in vm_os_lower for win in ['windows', 'win', 'microsoft']):
        return 'Windows Server'
    
    # Linux/Unix mapping
    if any(linux in vm_os_lower for linux in ['linux', 'ubuntu', 'debian', 'rhel', 'centos', 'red hat', 'suse', 'oracle linux']):
        return 'Linux'
    
    # Other Unix
    if any(unix in vm_os_lower for unix in ['unix', 'solaris', 'aix', 'hp-ux']):
        return 'Other Unix'
    
    # Default
    return 'Linux'

def create_template_output(results_df, template_path, output_path):
    """Create output file in Amazon EC2 Bulk Upload template format."""
    print(f"\nCreating output file {output_path}...")
    
    # Load the template workbook
    wb = load_workbook(template_path)
    
    # Get the Inputs sheet
    ws = wb['Inputs']
    
    # Clear existing data rows (keep header rows 1-3)
    # Delete rows from 4 onwards to remove all pre-filled data
    ws.delete_rows(4, ws.max_row - 3)
    
    # Column mapping based on template structure (Row 3)
    # Col 2: Group, Col 3: Description, Col 4: AWS Region, 
    # Col 5: Operating System, Col 6: Instance Type, Col 7: Tenancy,
    # Col 8: Number of Instances, Col 9: Assumed Usage, Col 10: Usage Type,
    # Col 11: Purchasing Option, Col 12: Storage Type, Col 13: Storage amount,
    # Col 14: IOPS, Col 15: Throughput, Col 16: Snapshot Frequency, Col 17: Snapshot amount
    
    # Write data starting from row 4
    for idx, row in results_df.iterrows():
        row_num = idx + 4  # Start from row 4
        
        # Group (optional) - use VM name as group identifier
        ws.cell(row=row_num, column=2, value='Production')
        
        # Description (optional) - use VM name
        ws.cell(row=row_num, column=3, value=str(row['VM'])[:50])  # Limit to 50 chars
        
        # AWS Region (required) - default to eu-central-1
        ws.cell(row=row_num, column=4, value='eu-central-1')
        
        # Operating System (required) - map from RVTools OS
        ec2_os = map_os_to_ec2(row['OS'])
        ws.cell(row=row_num, column=5, value=ec2_os)
        
        # Instance Type (required) - matched EC2 type
        ws.cell(row=row_num, column=6, value=row['EC2_Instance_Type'])
        
        # Tenancy (required) - default to Shared
        ws.cell(row=row_num, column=7, value='Shared Instances')
        
        # Number of Instances (required) - default to 1
        ws.cell(row=row_num, column=8, value=1)
        
        # Assumed Usage (required for Hours/Week) - empty
        ws.cell(row=row_num, column=9, value=None)
        
        # Usage Type (required) - Always On
        ws.cell(row=row_num, column=10, value='Always On')
        
        # Purchasing Option (required) - 1 Yr No Upfront Compute Savings Plan
        ws.cell(row=row_num, column=11, value='1 Yr No Upfront Compute Savings Plan')
        
        # Storage Type (optional) - gp3
        ws.cell(row=row_num, column=12, value='General Purpose SSD (gp3)')
        
        # Storage amount (optional) - use VM disk capacity
        storage_gb = max(8, int(row['Disk_GB']))  # Minimum 8GB
        ws.cell(row=row_num, column=13, value=storage_gb)
        
        # IOPS (optional for gp3) - empty
        ws.cell(row=row_num, column=14, value=None)
        
        # Throughput (optional for gp3) - empty
        ws.cell(row=row_num, column=15, value=None)
        
        # Snapshot Frequency (optional) - empty
        ws.cell(row=row_num, column=16, value=None)
        
        # Snapshot amount (optional) - empty
        ws.cell(row=row_num, column=17, value=None)
    
    # Save the workbook
    wb.save(output_path)
    print(f"Output file created successfully: {output_path}")
    print(f"Total instances written: {len(results_df)}")

def parse_arguments():
    """Parse command-line arguments."""
    parser = argparse.ArgumentParser(
        description='Match RVTools export data with EC2 instance types and generate AWS bulk upload template.',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python ec2_matching_tool.py
  python ec2_matching_tool.py --rvtools "My VMs.xlsx" --output "My_Output.xlsx"
  python ec2_matching_tool.py --rvtools "DAP YAPI.xlsx" --price "CloudPrice_Reference.xlsx"
        """
    )
    
    parser.add_argument(
        '--rvtools',
        type=str,
        default='DAP YAPI.xlsx',
        help='Path to RVTools export Excel file (default: DAP YAPI.xlsx)'
    )
    
    parser.add_argument(
        '--price-reference',
        type=str,
        default='CloudPrice_Reference.xlsx',
        help='Path to CloudPrice reference Excel file (contains EC2 instances and pricing) (default: CloudPrice_Reference.xlsx)'
    )
    
    parser.add_argument(
        '--template',
        type=str,
        default='Amazon_EC2_Instances_BulkUpload_Template_Commercial.xlsx',
        help='Path to AWS bulk upload template Excel file (default: Amazon_EC2_Instances_BulkUpload_Template_Commercial.xlsx)'
    )
    
    parser.add_argument(
        '--output',
        type=str,
        default='Amazon_EC2_output.xlsx',
        help='Path to output Excel file (default: Amazon_EC2_output.xlsx)'
    )
    
    return parser.parse_args()

def main():
    # Parse command-line arguments
    args = parse_arguments()
    
    # Validate input files exist
    for file_arg, file_path in [
        ('RVTools file', args.rvtools),
        ('Price reference file', args.price_reference),
        ('Template file', args.template)
    ]:
        if not os.path.exists(file_path):
            print(f"Error: {file_arg} not found: {file_path}")
            sys.exit(1)
    
    try:
        # Read input files
        rvtools_df = read_rvtools_data(args.rvtools)
        
        # Read EC2 instances from CloudPrice Reference (Linux tab)
        ec2_df = read_ec2_instances(args.price_reference)
        
        # Read all price sheets (Linux and Windows)
        price_sheets = read_all_price_sheets(args.price_reference)
        
        # Process matches with price consideration
        matched_df, unmatched_df = process_matches(rvtools_df, ec2_df, price_sheets)
        
        # Create output
        if len(matched_df) > 0:
            create_template_output(matched_df, args.template, args.output)
            
            # Print summary
            print("\n" + "="*60)
            print("MATCHING SUMMARY")
            print("="*60)
            print(f"Total VMs processed: {len(rvtools_df)}")
            print(f"Successfully matched: {len(matched_df)}")
            print(f"Unmatched: {len(unmatched_df)}")
            print(f"\nOutput file: {args.output}")
            print("\nSample matched VMs:")
            print(matched_df[['VM', 'CPUs', 'Memory_GB', 'Memory_Source', 'Disk_GB', 'OS', 'EC2_Instance_Type', 'Price_Per_Hour']].head(10).to_string(index=False))
        else:
            print("No VMs were matched. Please check the data.")
            
    except FileNotFoundError as e:
        print(f"Error: File not found - {e}")
        sys.exit(1)
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main()
