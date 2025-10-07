"""
SharePoint MCP tools using Microsoft Graph API
"""
import base64, os
from functools import wraps
from typing import Optional, Dict, Any
from .common import logger, mcp, ACCESS_TOKEN, SITE_ID, DRIVE_ID, make_graph_request
from .resources import list_folders, list_documents, get_document_content, download_document

# Helper functions
def _handle_sp_operation(func):
    """Decorator for SharePoint operations with error handling"""
    @wraps(func)
    async def wrapper(*args, **kwargs):
        try:
            return await func(*args, **kwargs)
        except Exception as e:
            logger.error(f"Error in {func.__name__}: {str(e)}")
            return {"success": False, "message": f"Operation failed: {str(e)}"}
    return wrapper

def _file_success_response(file_info: Dict[str, Any], message: str) -> Dict[str, Any]:
    """Standard success response for file operations"""
    return {
        "success": True,
        "message": message,
        "file": {
            "name": file_info.get('name'),
            "id": file_info.get('id'),
            "url": file_info.get('webUrl'),
            "size": file_info.get('size', 0)
        }
    }

def _folder_success_response(folder_info: Dict[str, Any], message: str) -> Dict[str, Any]:
    """Standard success response for folder operations"""
    return {
        "success": True,
        "message": message,
        "folder": {
            "name": folder_info.get('name'),
            "id": folder_info.get('id'),
            "url": folder_info.get('webUrl')
        }
    }

# Basic tool implementations
@mcp.tool(name="List_SharePoint_Folders", description="List folders in the specified SharePoint directory or root if not specified")
async def list_folders_tool(parent_folder: Optional[str] = None):
    """List folders in the specified SharePoint directory or root if not specified"""
    return list_folders(parent_folder)

@mcp.tool(name="List_SharePoint_Documents", description="List all documents in a specified SharePoint folder")
async def list_documents_tool(folder_name: str):
    """List all documents in a specified SharePoint folder"""
    return list_documents(folder_name)

@mcp.tool(name="Get_Document_Content", description="Get content of a document in SharePoint")
async def get_document_content_tool(folder_name: str, file_name: str):
    """Get content of a document in SharePoint"""
    return get_document_content(folder_name, file_name)

@mcp.tool(name="Download_Document", description="Download a document from SharePoint to local filesystem")
@_handle_sp_operation
async def download_document_tool(folder_name: str, file_name: str, local_path: str):
    """Download a document from SharePoint to local filesystem with fallback support"""
    return download_document(folder_name, file_name, local_path)

# Graph API specific tools
@mcp.tool(name="Create_Folder", description="Create a new folder in the specified directory or root if not specified")
@_handle_sp_operation
async def create_folder(folder_name: str, parent_folder: Optional[str] = None):
    """Create a new folder using Graph API"""
    logger.info(f"Creating folder '{folder_name}' in {parent_folder or 'root directory'}")
    
    # Check for existing folder
    existing_folders = list_folders(parent_folder)
    if any(f["name"] == folder_name for f in existing_folders):
        return {"success": False, "message": f"Folder {folder_name} already exists"}
    
    try:
        if not parent_folder or parent_folder == "":
            endpoint = f"sites/{SITE_ID}/drives/{DRIVE_ID}/root/children"
        else:
            endpoint = f"sites/{SITE_ID}/drives/{DRIVE_ID}/root:/{parent_folder}:/children"
        
        folder_data = {
            "name": folder_name,
            "folder": {},
            "@microsoft.graph.conflictBehavior": "rename"
        }
        
        response = make_graph_request("POST", endpoint, folder_data)
        
        if response and response.status_code == 201:
            folder_info = response.json()
            return _folder_success_response(folder_info, f"Folder {folder_name} created successfully")
        else:
            return {"success": False, "message": f"Failed to create folder: {response.status_code if response else 'No response'}"}
            
    except Exception as e:
        return {"success": False, "message": f"Error creating folder: {str(e)}"}

@mcp.tool(name="Upload_Document", description="Upload a new file to a SharePoint directory")
@_handle_sp_operation
async def upload_document(folder_name: str, file_name: str, content: str, is_base64: bool = False):
    """Upload a new file using Graph API"""
    logger.info(f"Uploading document {file_name} to folder {folder_name}")
    
    try:
        # Convert content
        file_content = base64.b64decode(content) if is_base64 else content.encode('utf-8')
        
        # Build endpoint
        if not folder_name or folder_name == "":
            endpoint = f"sites/{SITE_ID}/drives/{DRIVE_ID}/root:/{file_name}:/content"
        else:
            endpoint = f"sites/{SITE_ID}/drives/{DRIVE_ID}/root:/{folder_name}/{file_name}:/content"
        
        # Determine content type
        content_type = "text/plain" if file_name.lower().endswith('.txt') else "application/octet-stream"
        
        response = make_graph_request("PUT", endpoint, file_content, content_type)
        
        if response and response.status_code in [200, 201]:
            file_info = response.json()
            return _file_success_response(file_info, f"File {file_name} uploaded successfully")
        else:
            return {"success": False, "message": f"Failed to upload file: {response.status_code if response else 'No response'}"}
            
    except Exception as e:
        return {"success": False, "message": f"Error uploading file: {str(e)}"}

@mcp.tool(name="Delete_Document", description="Delete a document from a SharePoint directory")
@_handle_sp_operation
async def delete_document(folder_name: str, file_name: str):
    """Delete a document using Graph API"""
    logger.info(f"Deleting document {file_name} from folder {folder_name}")
    
    try:
        # Build endpoint
        if not folder_name or folder_name == "":
            endpoint = f"sites/{SITE_ID}/drives/{DRIVE_ID}/root:/{file_name}"
        else:
            endpoint = f"sites/{SITE_ID}/drives/{DRIVE_ID}/root:/{folder_name}/{file_name}"
        
        # Check if file exists
        check_response = make_graph_request("GET", endpoint)
        if not check_response or check_response.status_code != 200:
            return {"success": False, "message": f"File {file_name} does not exist"}
        
        # Delete the file
        response = make_graph_request("DELETE", endpoint)
        
        if response and response.status_code == 204:
            return {"success": True, "message": f"File {file_name} deleted successfully"}
        else:
            return {"success": False, "message": f"Failed to delete file: {response.status_code if response else 'No response'}"}
            
    except Exception as e:
        return {"success": False, "message": f"Error deleting file: {str(e)}"}
