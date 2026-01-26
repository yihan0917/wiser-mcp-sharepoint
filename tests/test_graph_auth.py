import os
import requests
from msal import ConfidentialClientApplication
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Configuration
CLIENT_ID = os.getenv('SHP_ID_APP')  # Your Azure AD App ID
CLIENT_SECRET = os.getenv('SHP_ID_APP_SECRET')  # Your Azure AD App Secret
TENANT_ID = os.getenv('SHP_TENANT_ID')  # Your Azure AD Tenant ID
SITE_URL = os.getenv('SHP_SITE_URL')  # Your SharePoint site URL

# Microsoft Graph API configuration
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["https://graph.microsoft.com/.default"]

def get_access_token():
    """Get access token using MSAL"""
    try:
        app = ConfidentialClientApplication(
            CLIENT_ID,
            authority=AUTHORITY,
            client_credential=CLIENT_SECRET,
        )
        
        # Acquire token for client credentials flow
        # Step 1: Try to get cached token (fast)
        result = app.acquire_token_silent(SCOPES, account=None)
        
        # Step 2: If no cached token, acquire new token (slower)
        if not result:
            result = app.acquire_token_for_client(scopes=SCOPES)
        
        if "access_token" in result:
            return result["access_token"]
        else:
            print(f"Token acquisition failed: {result.get('error')}")
            print(f"Error description: {result.get('error_description')}")
            return None
            
    except Exception as e:
        print(f"Error getting access token: {e}")
        return None

def test_sharepoint_access(access_token):
    """Test SharePoint access using Graph API"""
    try:
        # Extract site name from SharePoint URL
        # Example: https://tenant.sharepoint.com/sites/sitename -> sitename
        site_parts = SITE_URL.rstrip('/').split('/')
        site_name = site_parts[-1] if 'sites' in site_parts else None
        tenant_name = site_parts[2].split('.')[0] if len(site_parts) > 2 else None
        
        if not site_name or not tenant_name:
            print("Could not parse site name from URL")
            return False
        
        # Graph API endpoint for SharePoint site
        graph_url = f"https://graph.microsoft.com/v1.0/sites/{tenant_name}.sharepoint.com:/sites/{site_name}"
        
        # The Accept header tells the server what format you want the response in. Without Accept header: Server uses default format (usually JSON for Graph API)
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Accept": "application/json"
        }
        
        # Making one HTTP GET request to the Graph API endpoint for SharePoint site
        response = requests.get(graph_url, headers=headers)
        
        if response.status_code == 200:
            site_info = response.json()
            print(f"✅ Successfully connected to SharePoint site!")
            print(f"Site Name: {site_info.get('displayName', 'N/A')}")
            print(f"Site ID: {site_info.get('id', 'N/A')}")
            print(f"Web URL: {site_info.get('webUrl', 'N/A')}")
            return True
        else:
            print(f"❌ Failed to access SharePoint site")
            print(f"Status Code: {response.status_code}")
            print(f"Response: {response.text}")
            return False
            
    except Exception as e:
        print(f"Error testing SharePoint access: {e}")
        return False

def test_document_libraries(access_token):
    """Test accessing document libraries"""
    try:
        site_parts = SITE_URL.rstrip('/').split('/')
        site_name = site_parts[-1] if 'sites' in site_parts else None
        tenant_name = site_parts[2].split('.')[0] if len(site_parts) > 2 else None
        
        # Get document libraries (drives)
        graph_url = f"https://graph.microsoft.com/v1.0/sites/{tenant_name}.sharepoint.com:/sites/{site_name}:/drives"
        
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Accept": "application/json"
        }
        
        response = requests.get(graph_url, headers=headers)
        
        if response.status_code == 200:
            drives = response.json()
            print(f"\n📁 Document Libraries found: {len(drives.get('value', []))}")
            for drive in drives.get('value', [])[:5]:  # Show first 5
                print(f"  - {drive.get('name', 'N/A')} (ID: {drive.get('id', 'N/A')})")
            return True
        else:
            print(f"❌ Failed to get document libraries")
            print(f"Status Code: {response.status_code}")
            print(f"Response: {response.text}")
            return False
            
    except Exception as e:
        print(f"Error getting document libraries: {e}")
        return False

def main():
    """Main test function"""
    print("🔐 Testing Microsoft Graph API Authentication...")
    print(f"Tenant ID: {TENANT_ID}")
    print(f"Client ID: {CLIENT_ID}")
    print(f"Site URL: {SITE_URL}")
    print("-" * 50)
    
    # Check required environment variables
    if not all([CLIENT_ID, CLIENT_SECRET, TENANT_ID, SITE_URL]):
        print("❌ Missing required environment variables:")
        if not CLIENT_ID: print("  - SHP_ID_APP")
        if not CLIENT_SECRET: print("  - SHP_ID_APP_SECRET") 
        if not TENANT_ID: print("  - SHP_TENANT_ID")
        if not SITE_URL: print("  - SHP_SITE_URL")
        return
    
    # Get access token
    print("🔑 Acquiring access token...")
    access_token = get_access_token()
    
    if not access_token:
        print("❌ Failed to get access token")
        return
    
    print("✅ Access token acquired successfully")
    
    # Test SharePoint site access
    print("\n🌐 Testing SharePoint site access...")
    if test_sharepoint_access(access_token):
        # Test document libraries
        print("\n📚 Testing document libraries access...")
        test_document_libraries(access_token)
    
    print("\n" + "=" * 50)
    print("Test completed!")

if __name__ == "__main__":
    main()
