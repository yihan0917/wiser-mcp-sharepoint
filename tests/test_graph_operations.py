import os
import requests
import json
from msal import ConfidentialClientApplication
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Configuration
CLIENT_ID = os.getenv('SHP_ID_APP')
CLIENT_SECRET = os.getenv('SHP_ID_APP_SECRET')
TENANT_ID = os.getenv('SHP_TENANT_ID')
SITE_URL = os.getenv('SHP_SITE_URL')

# Microsoft Graph API configuration
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["https://graph.microsoft.com/.default"]

# Global variables for site info
SITE_ID = None
DRIVE_ID = None
ACCESS_TOKEN = None

def get_access_token():
    """Get access token using MSAL"""
    global ACCESS_TOKEN
    try:
        app = ConfidentialClientApplication(
            CLIENT_ID,
            authority=AUTHORITY,
            client_credential=CLIENT_SECRET,
        )
        
        result = app.acquire_token_silent(SCOPES, account=None)
        if not result:
            result = app.acquire_token_for_client(scopes=SCOPES)
        
        if "access_token" in result:
            ACCESS_TOKEN = result["access_token"]
            return ACCESS_TOKEN
        else:
            print(f"Token acquisition failed: {result.get('error')}")
            return None
            
    except Exception as e:
        print(f"Error getting access token: {e}")
        return None

def get_site_info():
    """Get SharePoint site information and store site ID"""
    global SITE_ID
    try:
        site_parts = SITE_URL.rstrip('/').split('/')
        site_name = site_parts[-1]
        tenant_name = site_parts[2].split('.')[0]
        
        graph_url = f"https://graph.microsoft.com/v1.0/sites/{tenant_name}.sharepoint.com:/sites/{site_name}"
        
        headers = {
            "Authorization": f"Bearer {ACCESS_TOKEN}",
            "Accept": "application/json"
        }
        
        response = requests.get(graph_url, headers=headers)
        
        if response.status_code == 200:
            site_info = response.json()
            SITE_ID = site_info.get('id')
            print(f"✅ Site ID: {SITE_ID}")
            return site_info
        else:
            print(f"❌ Failed to get site info: {response.status_code}")
            return None
            
    except Exception as e:
        print(f"Error getting site info: {e}")
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
            print(f"\n📁 Document Libraries ({len(drives.get('value', []))}):")
            
            for drive in drives.get('value', []):
                drive_id = drive.get('id')
                drive_name = drive.get('name')
                drive_type = drive.get('driveType', 'unknown')
                
                print(f"  - {drive_name} ({drive_type})")
                print(f"    ID: {drive_id}")
                
                # Use the first drive as default for testing
                if not DRIVE_ID:
                    DRIVE_ID = drive_id
                    print(f"    ⭐ Using this drive for testing")
                
            return drives.get('value', [])
        else:
            print(f"❌ Failed to get drives: {response.status_code}")
            print(f"Response: {response.text}")
            return []
            
    except Exception as e:
        print(f"Error getting drives: {e}")
        return []

def list_root_items():
    """List items in the root of the document library"""
    try:
        graph_url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/drives/{DRIVE_ID}/root/children"
        
        headers = {
            "Authorization": f"Bearer {ACCESS_TOKEN}",
            "Accept": "application/json"
        }
        
        response = requests.get(graph_url, headers=headers)
        
        if response.status_code == 200:
            items = response.json()
            print(f"\n📂 Root Directory Items ({len(items.get('value', []))}):")
            
            for item in items.get('value', []):
                item_name = item.get('name')
                item_type = 'folder' if 'folder' in item else 'file'
                item_size = item.get('size', 0)
                created = item.get('createdDateTime', 'Unknown')
                modified = item.get('lastModifiedDateTime', 'Unknown')
                
                icon = "📁" if item_type == 'folder' else "📄"
                size_str = f" ({item_size} bytes)" if item_type == 'file' else ""
                
                print(f"  {icon} {item_name}{size_str}")
                print(f"    Created: {created[:19] if created != 'Unknown' else 'Unknown'}")
                print(f"    Modified: {modified[:19] if modified != 'Unknown' else 'Unknown'}")
                
                if item_type == 'file':
                    download_url = item.get('@microsoft.graph.downloadUrl')
                    if download_url:
                        print(f"    Download URL available: Yes")
                
            return items.get('value', [])
        else:
            print(f"❌ Failed to list root items: {response.status_code}")
            print(f"Response: {response.text}")
            return []
            
    except Exception as e:
        print(f"Error listing root items: {e}")
        return []

def get_file_content(file_name):
    """Get file content using direct path - more reliable"""
    try:
        # Direct path approach
        file_url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/drives/{DRIVE_ID}/root:/{file_name}"
        headers = {"Authorization": f"Bearer {ACCESS_TOKEN}", "Accept": "application/json"}
        
        response = requests.get(file_url, headers=headers)
        
        if response.status_code == 200:
            file_info = response.json()
            download_url = file_info.get('@microsoft.graph.downloadUrl')
            
            if download_url:
                content_response = requests.get(download_url)
                if content_response.status_code == 200:
                    print(f"✅ Successfully got '{file_name}'")
                    print(f"   Size: {len(content_response.content)} bytes")
                    
                    try:
                        text_content = content_response.content.decode('utf-8')
                        print(f"   Content: {text_content}")
                    except:
                        print(f"   Binary content")
                    
                    return content_response.content
        
        print(f"❌ Could not get file '{file_name}': {response.status_code}")
        return None
        
    except Exception as e:
        print(f"Error getting file: {e}")
        return None

def create_test_folder(folder_name="test_folder_graph"):
    """Create a test folder"""
    try:
        graph_url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/drives/{DRIVE_ID}/root/children"
        
        headers = {
            "Authorization": f"Bearer {ACCESS_TOKEN}",
            "Content-Type": "application/json"
        }
        
        folder_data = {
            "name": folder_name,
            "folder": {},
            "@microsoft.graph.conflictBehavior": "rename"
        }
        
        response = requests.post(graph_url, headers=headers, json=folder_data)
        
        if response.status_code == 201:
            folder_info = response.json()
            print(f"✅ Created folder: {folder_info.get('name')}")
            print(f"   ID: {folder_info.get('id')}")
            return folder_info
        else:
            print(f"❌ Failed to create folder: {response.status_code}")
            print(f"Response: {response.text}")
            return None
            
    except Exception as e:
        print(f"Error creating folder: {e}")
        return None

def upload_test_file(file_name="test_file.txt", content="Hello from Graph API!"):
    """Upload a test file"""
    try:
        graph_url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/drives/{DRIVE_ID}/root:/{file_name}:/content"
        
        headers = {
            "Authorization": f"Bearer {ACCESS_TOKEN}",
            "Content-Type": "text/plain"
        }
        
        response = requests.put(graph_url, headers=headers, data=content.encode('utf-8'))
        
        if response.status_code in [200, 201]:
            file_info = response.json()
            print(f"✅ Uploaded file: {file_info.get('name')}")
            print(f"   Size: {file_info.get('size')} bytes")
            print(f"   ID: {file_info.get('id')}")
            return file_info
        else:
            print(f"❌ Failed to upload file: {response.status_code}")
            print(f"Response: {response.text}")
            return None
            
    except Exception as e:
        print(f"Error uploading file: {e}")
        return None

def main():
    """Main test function"""
    print("🔐 Testing Microsoft Graph API SharePoint Operations...")
    print("=" * 60)
    
    # Step 1: Get access token
    print("🔑 Getting access token...")
    if not get_access_token():
        print("❌ Failed to get access token")
        return
    print("✅ Access token acquired")
    
    # Step 2: Get site info
    print("\n🌐 Getting site information...")
    if not get_site_info():
        print("❌ Failed to get site info")
        return
    
    # Step 3: Get drives
    print("\n📚 Getting document libraries...")
    drives = get_drives()
    if not drives:
        print("❌ No drives found")
        return
    
    # Step 4: List root items
    print("\n📂 Listing root directory items...")
    root_items = list_root_items()
    
    # Step 6: Create test folder
    print("\n📁 Creating test folder...")
    test_folder = create_test_folder()
    
    # Step 7: Upload test file
    print("\n📄 Uploading test file...")
    test_file = upload_test_file()
    
    # Step 8: Get file content (if we uploaded successfully)
    if test_file:
        print("\n📖 Reading test file content...")
        get_file_content("test_file.txt")

    print("\n" + "=" * 60)
    print("✅ SharePoint operations test completed!")

if __name__ == "__main__":
    main()
