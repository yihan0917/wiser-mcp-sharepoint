#!/usr/bin/env python3
"""
Test script for downloading Excel files from SharePoint and verifying content
"""
import sys
import os
import pandas as pd
import datetime
from pathlib import Path

sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from mcp_sharepoint.resources import download_document, list_documents

def test_download_excel():
    """Test downloading Excel file and reading its content"""
    print("🚀 Testing Excel File Download from SharePoint\n")
    
    # Configuration - modify these values for your specific file
    folder_name = "Data"  # Change to your folder
    excel_filename = "2023 Recruiting Dataset  .xlsx"  # Change to your Excel file
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S") # Save with timestamp to avoid overwriting
    local_download_path = f"/Users/yihan/Desktop/{excel_filename}_{timestamp}.xlsx"  # Local path to save
    
    print(f"📁 Folder: {folder_name}")
    print(f"📊 Excel File: {excel_filename}")
    print(f"💾 Local Path: {local_download_path}")
    print()
    
    # Step 1: List available Excel files in the folder
    print("🔍 Step 1: Checking available Excel files...")
    try:
        documents = list_documents(folder_name)
        excel_files = [d for d in documents if d['name'].lower().endswith(('.xlsx', '.xls'))]
        
        if excel_files:
            print(f"✅ Found {len(excel_files)} Excel files:")
            for i, excel_file in enumerate(excel_files, 1):
                size = excel_file.get('size', 0)
                print(f"  {i}. {excel_file['name']} ({size:,} bytes)")
            
            # Check if our target file exists
            target_file = next((f for f in excel_files if f['name'] == excel_filename), None)
            if not target_file:
                print(f"\n❌ Target file '{excel_filename}' not found.")
                print("Available Excel files:")
                for excel_file in excel_files:
                    print(f"  - {excel_file['name']}")
                
                # Use the first available Excel file for testing
                if excel_files:
                    target_file = excel_files[0]
                    excel_filename = target_file['name']
                    print(f"\n🔄 Using first available file: '{excel_filename}'")
                else:
                    print("No Excel files found to test with.")
                    return
        else:
            print(f"❌ No Excel files found in folder '{folder_name}'")
            return
            
    except Exception as e:
        print(f"❌ Error listing documents: {e}")
        return
    
    # Step 2: Download the Excel file
    print(f"\n📥 Step 2: Downloading '{excel_filename}'...")
    try:
        result = download_document(folder_name, excel_filename, local_download_path)
        
        if result.get('success'):
            print(f"✅ Download successful!")
            print(f"  - Message: {result['message']}")
            print(f"  - Local Path: {result['local_path']}")
            print(f"  - File Size: {result['size']:,} bytes")
        else:
            print(f"❌ Download failed: {result.get('message')}")
            return
            
    except Exception as e:
        print(f"❌ Error during download: {e}")
        return
    
    # Step 3: Verify the file exists locally
    print(f"\n🔍 Step 3: Verifying downloaded file...")
    if os.path.exists(local_download_path):
        file_size = os.path.getsize(local_download_path)
        print(f"✅ File exists locally: {file_size:,} bytes")
    else:
        print(f"❌ Downloaded file not found at {local_download_path}")
        return
    
    # Step 4: Read and display Excel content
    print(f"\n📊 Step 4: Reading Excel content...")
    try:
        # Read Excel file with pandas
        df = pd.read_excel(local_download_path)
        
        print(f"✅ Excel file successfully read!")
        print(f"  - Shape: {df.shape[0]} rows × {df.shape[1]} columns")
        print(f"  - Columns: {list(df.columns)}")
        
        # Display first few rows
        print(f"\n📋 First 5 rows of data:")
        print(df.head().to_string(index=False))
        
        # Display data types
        print(f"\n🔢 Data types:")
        for col, dtype in df.dtypes.items():
            print(f"  - {col}: {dtype}")
        
        # Display basic statistics for numeric columns
        numeric_cols = df.select_dtypes(include=['number']).columns
        if len(numeric_cols) > 0:
            print(f"\n📈 Basic statistics for numeric columns:")
            print(df[numeric_cols].describe().to_string())
        
        # Display unique values for categorical columns (first few)
        categorical_cols = df.select_dtypes(include=['object']).columns
        if len(categorical_cols) > 0:
            print(f"\n📝 Sample values for text columns:")
            for col in categorical_cols[:3]:  # Show first 3 categorical columns
                unique_vals = df[col].unique()[:5]  # Show first 5 unique values
                print(f"  - {col}: {list(unique_vals)}")
                if len(df[col].unique()) > 5:
                    print(f"    ... and {len(df[col].unique()) - 5} more unique values")
        
    except Exception as e:
        print(f"❌ Error reading Excel file: {e}")
        print("This might indicate the file is corrupted or not a valid Excel file.")
        
        # Try to read as binary and show first few bytes
        try:
            with open(local_download_path, 'rb') as f:
                first_bytes = f.read(50)
                print(f"First 50 bytes (hex): {first_bytes.hex()}")
                print(f"First 50 bytes (repr): {repr(first_bytes)}")
        except Exception as read_error:
            print(f"Could not even read as binary: {read_error}")
        return
    
    # # Step 5: Cleanup (optional)
    # print(f"\n🧹 Step 5: Cleanup...")
    # try:
    #     os.remove(local_download_path)
    #     print(f"✅ Temporary file removed: {local_download_path}")
    # except Exception as e:
    #     print(f"⚠️  Could not remove temporary file: {e}")
    
    print(f"\n✅ Excel download and verification completed successfully!")

if __name__ == "__main__":
    test_download_excel()
