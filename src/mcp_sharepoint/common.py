"""
MCP SharePoint server using Microsoft Graph API
"""
import os, logging
import requests
from dotenv import load_dotenv
from mcp.server.fastmcp import FastMCP
from msal import ConfidentialClientApplication

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger('mcp_sharepoint')

# Load environment variables
load_dotenv()

# Configuration
SHP_ID_APP = os.getenv('SHP_ID_APP')
SHP_ID_APP_SECRET = os.getenv('SHP_ID_APP_SECRET')
SHP_SITE_URL = os.getenv('SHP_SITE_URL')
SHP_DOC_LIBRARY = os.getenv('SHP_DOC_LIBRARY', 'Shared Documents')
SHP_TENANT_ID = os.getenv('SHP_TENANT_ID')

# Validate required environment variables
if not SHP_SITE_URL:
    logger.error("SHP_SITE_URL environment variable not set.")
    raise ValueError("SHP_SITE_URL environment variable not set.")
if not SHP_ID_APP:
    logger.error("SHP_ID_APP environment variable not set.")
    raise ValueError("SHP_ID_APP environment variable not set.")
if not SHP_ID_APP_SECRET:
    logger.error("SHP_ID_APP_SECRET environment variable not set.")
    raise ValueError("SHP_ID_APP_SECRET environment variable not set.")
if not SHP_TENANT_ID:
    logger.error("SHP_TENANT_ID environment variable not set.")
    raise ValueError("SHP_TENANT_ID environment variable not set.")

# Microsoft Graph API configuration
AUTHORITY = f"https://login.microsoftonline.com/{SHP_TENANT_ID}"
SCOPES = ["https://graph.microsoft.com/.default"]

# Global variables for Graph API
ACCESS_TOKEN = None
SITE_ID = None
DRIVE_ID = None

# Initialize MCP server
mcp = FastMCP(
    name="mcp_sharepoint",
    instructions=f"This server provides tools to interact with SharePoint documents and folders in {SHP_DOC_LIBRARY} using Microsoft Graph API"
)

# MSAL application for authentication
msal_app = ConfidentialClientApplication(
    SHP_ID_APP,
    authority=AUTHORITY,
    client_credential=SHP_ID_APP_SECRET,
)

def get_access_token():
    """Get access token using MSAL (Microsoft Graph API)"""
    global ACCESS_TOKEN # ← Declares intent to modify the global ACCESS_TOKEN variable
    try:
        # Step 1: Try to get cached token (fast)
        result = msal_app.acquire_token_silent(SCOPES, account=None) # ← account=None means No specific user account - Look for any cached token that matches the scopes. No specific user account because your app (daemon app) authenticates as itself (not as any specific user)
        
        # Step 2: If no cached token, acquire new token (slower)
        if not result:
            result = msal_app.acquire_token_for_client(scopes=SCOPES)
        
        if "access_token" in result:
            ACCESS_TOKEN = result["access_token"] # ← Modifies the global ACCESS_TOKEN
            logger.info("Successfully acquired access token")
            return ACCESS_TOKEN
        else:
            logger.error(f"Token acquisition failed: {result.get('error')}")
            logger.error(f"Error description: {result.get('error_description')}")
            return None
            
    except Exception as e:
        logger.error(f"Error getting access token: {e}")
        return None

def get_site_info():
    """Get SharePoint site information and store site ID"""
    global SITE_ID
    try:
        # Parse site URL to get tenant and site name
        site_parts = SHP_SITE_URL.rstrip('/').split('/')
        site_name = site_parts[-1] if 'sites' in site_parts else None
        tenant_name = site_parts[2].split('.')[0] if len(site_parts) > 2 else None
        
        if not site_name or not tenant_name:
            logger.error("Could not parse site name from SHP_SITE_URL")
            return None
        
        # Graph API endpoint for SharePoint site
        graph_url = f"https://graph.microsoft.com/v1.0/sites/{tenant_name}.sharepoint.com:/sites/{site_name}"
        
        headers = {
            "Authorization": f"Bearer {ACCESS_TOKEN}",
            "Accept": "application/json"
        }
        
        response = requests.get(graph_url, headers=headers)
        
        if response.status_code == 200:
            site_info = response.json()
            SITE_ID = site_info.get('id')
            logger.info(f"Successfully connected to SharePoint site: {site_info.get('displayName')}")
            logger.info(f"Site ID: {SITE_ID}")
            return site_info
        else:
            logger.error(f"Failed to get site info: {response.status_code}")
            logger.error(f"Response: {response.text}")
            return None
            
    except Exception as e:
        logger.error(f"Error getting site info: {e}")
        return None

def get_drives():
    """Get all document libraries (drives) in the site"""
    global DRIVE_ID
    try:
        graph_url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/drives"
        
        headers = {
            "Authorization": f"Bearer {ACCESS_TOKEN}",
            "Accept": "application/json"
        }
        
        response = requests.get(graph_url, headers=headers)
        
        if response.status_code == 200:
            drives = response.json()
            logger.info(f"Found {len(drives.get('value', []))} document libraries")
            
            for drive in drives.get('value', []):
                drive_id = drive.get('id')
                drive_name = drive.get('name')
                drive_type = drive.get('driveType', 'unknown')
                
                logger.info(f"Drive: {drive_name} ({drive_type}) - ID: {drive_id}")
                
                # Use the first drive as default (usually "Documents")
                if not DRIVE_ID:
                    DRIVE_ID = drive_id
                    logger.info(f"Using drive '{drive_name}' as default")
                
            return drives.get('value', [])
        else:
            logger.error(f"Failed to get drives: {response.status_code}")
            logger.error(f"Response: {response.text}")
            return []
            
    except Exception as e:
        logger.error(f"Error getting drives: {e}")
        return []

def initialize_sharepoint():
    """Initialize SharePoint connection using Graph API"""
    logger.info("Initializing SharePoint connection with Microsoft Graph API...")
    
    # Step 1: Get access token
    if not get_access_token():
        logger.error("Failed to get access token")
        return False
    
    # Step 2: Get site information
    if not get_site_info():
        logger.error("Failed to get site information")
        return False
    
    # Step 3: Get document libraries
    if not get_drives():
        logger.error("Failed to get document libraries")
        return False
    
    logger.info("SharePoint initialization completed successfully")
    return True

def get_graph_headers():
    """Get standard headers for Graph API requests"""
    return {
        "Authorization": f"Bearer {ACCESS_TOKEN}",
        "Accept": "application/json"
    }

def make_graph_request(method, endpoint, data=None, content_type=None):
    """Make a request to Microsoft Graph API"""
    base_url = "https://graph.microsoft.com/v1.0"
    full_url = f"{base_url}/{endpoint}"
    
    headers = get_graph_headers()
    if content_type:
        headers["Content-Type"] = content_type
    elif data:
        headers["Content-Type"] = "application/json"
    
    try:
        if method.upper() == "GET":
            response = requests.get(full_url, headers=headers)
        elif method.upper() == "POST":
            response = requests.post(full_url, headers=headers, json=data if isinstance(data, dict) else None, data=data if not isinstance(data, dict) else None)
        elif method.upper() == "PATCH":
            response = requests.patch(full_url, headers=headers, json=data)
        elif method.upper() == "PUT":
            response = requests.put(full_url, headers=headers, data=data)
        elif method.upper() == "DELETE":
            response = requests.delete(full_url, headers=headers)
        else:
            raise ValueError(f"Unsupported HTTP method: {method}")
        
        return response
        
    except Exception as e:
        logger.error(f"Error making Graph API request: {e}")
        return None

# Initialize SharePoint connection on module import
# What happens when you import: from mcp_sharepoint.common import mcp
# 1. Python loads common.py
# 2. Executes all top-level code
# 3. Reaches: if not initialize_sharepoint():
# 4. Calls initialize_sharepoint()
# 5. Calls get_access_token() → sets global ACCESS_TOKEN
# 6. Calls get_site_info() → sets global SITE_ID  
# 7. Calls get_drives() → sets global DRIVE_ID
# 8. Module is ready for use
if not initialize_sharepoint():
    logger.warning("SharePoint initialization failed - some operations may not work")