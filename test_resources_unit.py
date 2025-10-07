#!/usr/bin/env python3
"""
Unit tests for SharePoint Graph API resources
"""
import pytest
import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from mcp_sharepoint.resources import list_folders, list_documents, get_document_content
from mcp_sharepoint.common import ACCESS_TOKEN, SITE_ID, DRIVE_ID

class TestSharePointResources:
    
    def test_connection_variables(self):
        """Test that connection variables are set"""
        assert ACCESS_TOKEN is not None, "ACCESS_TOKEN should be set"
        assert SITE_ID is not None, "SITE_ID should be set"
        assert DRIVE_ID is not None, "DRIVE_ID should be set"
    
    def test_list_folders(self):
        """Test listing folders"""
        folders = list_folders()
        assert isinstance(folders, list), "Should return a list"
        
        if folders:
            folder = folders[0]
            assert 'name' in folder, "Folder should have name"
            assert 'type' in folder, "Folder should have type"
            assert folder['type'] == 'folder', "Type should be 'folder'"
    
    def test_list_documents(self):
        """Test listing documents"""
        # First get a folder to test with
        folders = list_folders()
        if not folders:
            pytest.skip("No folders available for testing")
        
        folder_name = folders[0]['name']
        documents = list_documents(folder_name)
        assert isinstance(documents, list), "Should return a list"
        
        if documents:
            doc = documents[0]
            assert 'name' in doc, "Document should have name"
            assert 'type' in doc, "Document should have type"
            assert doc['type'] == 'file', "Type should be 'file'"
    
    def test_get_document_content_nonexistent(self):
        """Test getting content of non-existent document"""
        content = get_document_content("nonexistent_folder", "nonexistent_file.txt")
        assert 'error' in content or content.get('name') == 'nonexistent_file.txt'

if __name__ == "__main__":
    pytest.main([__file__, "-v"])
