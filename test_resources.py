#!/usr/bin/env python3
"""
Test script for SharePoint Graph API resources
"""
import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from mcp_sharepoint.resources import list_folders, list_documents, get_document_content
from mcp_sharepoint.common import ACCESS_TOKEN, SITE_ID, DRIVE_ID, logger

def test_basic_connection():
    """Test if we have valid connection info"""
    print("🔍 Testing Basic Connection...")
    print(f"ACCESS_TOKEN: {'✅ Set' if ACCESS_TOKEN else '❌ Missing'}")
    print(f"SITE_ID: {SITE_ID[:50] + '...' if SITE_ID else '❌ Missing'}")
    print(f"DRIVE_ID: {DRIVE_ID[:50] + '...' if DRIVE_ID else '❌ Missing'}")
    print()

def test_list_folders():
    """Test listing folders in root directory"""
    print("📁 Testing list_folders()...")
    try:
        folders = list_folders()
        print(f"✅ Found {len(folders)} folders:")
        for folder in folders[:5]:  # Show first 5
            print(f"  - {folder.get('name', 'Unknown')}")
        if len(folders) > 5:
            print(f"  ... and {len(folders) - 5} more")
        return folders
    except Exception as e:
        print(f"❌ Error: {e}")
        return []

def test_list_documents():
    """Test listing documents in a folder"""
    print("\n📄 Testing list_documents()...")
    
    # First get folders to test with
    folders = list_folders()
    if not folders:
        print("❌ No folders found to test with")
        return []
    
    # Test with first folder
    test_folder = folders[0]['name']
    print(f"Testing with folder: '{test_folder}'")
    
    try:
        documents = list_documents(test_folder)
        print(f"✅ Found {len(documents)} documents:")
        for doc in documents[:3]:  # Show first 3
            size = doc.get('size', 0)
            size_str = f"{size:,} bytes" if size else "Unknown size"
            print(f"  - {doc.get('name', 'Unknown')} ({size_str})")
        if len(documents) > 3:
            print(f"  ... and {len(documents) - 3} more")
        return documents, test_folder
    except Exception as e:
        print(f"❌ Error: {e}")
        return [], test_folder

def test_get_document_content():
    """Test getting document content"""
    print("\n📖 Testing get_document_content()...")
    
    # Get a document to test with
    documents, folder_name = test_list_documents()
    if not documents:
        print("❌ No documents found to test with")
        return
    
    # Find a small text file to test with
    test_doc = None
    for doc in documents:
        name = doc.get('name', '').lower()
        size = doc.get('size', 0)
        if any(name.endswith(ext) for ext in ['.txt', '.md', '.json', '.csv']) and size < 10000:
            test_doc = doc
            break
    
    if not test_doc:
        # Just use the first document
        test_doc = documents[0]
    
    print(f"Testing with document: '{test_doc['name']}'")
    
    try:
        content = get_document_content(folder_name, test_doc['name'])
        
        if 'error' in content:
            print(f"❌ Error: {content['error']}")
        else:
            content_type = content.get('content_type', 'unknown')
            size = content.get('size', 0)
            print(f"✅ Successfully retrieved content:")
            print(f"  - Type: {content_type}")
            print(f"  - Size: {size:,} bytes")
            
            if content_type == 'text':
                preview = content.get('content', '')[:200]
                print(f"  - Preview: {preview}{'...' if len(content.get('content', '')) > 200 else ''}")
            else:
                print(f"  - Binary content (base64 encoded)")
                
    except Exception as e:
        print(f"❌ Error: {e}")

def test_get_excel_content(folder_name, file_name):
    """Test getting document content from an Excel file"""
    print("\n📖 Testing get_document_content() from an Excel file...")
    
    print(f"Testing with document: '{file_name}'")
    
    try:
        content = get_document_content(folder_name, file_name)
        
        if 'error' in content:
            print(f"❌ Error: {content['error']}")
        else:
            content_type = content.get('content_type', 'unknown')
            size = content.get('size', 0)
            print(f"✅ Successfully retrieved content:")
            print(f"  - Type: {content_type}")
            print(f"  - Size: {size:,} bytes")
            
            if content_type == 'text':
                preview = content.get('content', '')[:200]
                print(f"  - Preview: {preview}{'...' if len(content.get('content', '')) > 200 else ''}")
            else:
                print(f"  - Binary content (base64 encoded)")
                
    except Exception as e:
        print(f"❌ Error: {e}")

def main():
    """Run all tests"""
    print("🚀 Testing SharePoint Graph API Resources\n")
    
    # Test basic connection
    test_basic_connection()
    
    # Test folder listing
    test_list_folders()
    
    # Test document listing and content retrieval
    test_get_document_content()

    test_get_excel_content('Data', '2023 Recruiting Dataset  .xlsx')
    
    print("\n✅ Testing completed!")

if __name__ == "__main__":
    main()
