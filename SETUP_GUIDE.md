# SharePoint MCP Server Setup Guide

This guide documents the complete setup process for the SharePoint MCP server.

## Notes

- **Graph API is recommended** for new implementations due to better reliability
- The `.egg-info` folder is created during `pip install -e .` and is normal
- Always keep your `.env` file secure and never commit it to git
- Server requires valid SharePoint credentials to start properly
- Graph API provides better error messages and debugging information than Office365 REST API (Legacy)

## Prerequisites

- Python 3.10 or higher
- Access to Azure Portal (for SharePoint app registration)
- SharePoint site access

## 1. Environment Setup

### Check Python Version
```bash
python3 --version
python3.10 --version
python3.11 --version
python3.12 --version
```

### Create Virtual Environment
```bash
# Use Python 3.10+ (project requires >=3.10)
python3.10 -m venv venv
source venv/bin/activate
```

### Install Dependencies
```bash
pip install -e .
```

**Note:** After running `pip install -e .`, you'll see a `mcp_sharepoint.egg-info` folder created under `src/`. This is normal behavior:
- **What it is**: Package metadata created by setuptools during editable installation
- **Purpose**: Stores dependency information, package metadata, and source file lists
- **Safe to ignore**: It's automatically git ignored and will be recreated if deleted
- **Editable mode**: The `-e` flag allows code changes to be immediately reflected without reinstalling

## 2. Configuration

### Create Environment File
```bash
cp .env.example .env
```

### Configure .env File
Edit `.env` with your actual SharePoint credentials:
```bash
# Azure AD Application Settings
SHP_ID_APP=your-actual-azure-app-client-id
SHP_ID_APP_SECRET=your-actual-azure-app-client-secret
SHP_TENANT_ID=your-actual-microsoft-tenant-id

# SharePoint Settings
SHP_SITE_URL=https://your-tenant.sharepoint.com/sites/your-site
SHP_DOC_LIBRARY=Shared Documents/mcp_server

# Tree operation limits (keep defaults)
SHP_MAX_DEPTH=15
SHP_MAX_FOLDERS_PER_LEVEL=100
SHP_LEVEL_DELAY=0.5
```

## 3. Azure App Registration Setup

### Required API Permissions
In Azure Portal → App Registrations → Your App → API Permissions:

**Add these Application permissions (not Delegated):**
- **SharePoint** → `Sites.FullControl.All`
- **Microsoft Graph** → `Sites.Read.All` or `Sites.ReadWrite.All`

### Grant Admin Consent
Click **"Grant admin consent for [Your Organization]"** after adding permissions.

### Verify App Settings
- **Authentication** → "Allow public client flows" should be **disabled**
- **Certificates & secrets** → Ensure client secret is valid and not expired

## 4. Starting the Server

### Method 1: Using Script Entry Point (Recommended)
```bash
mcp.sharepoint
```

### Method 2: Direct Module Execution
```bash
python -m mcp_sharepoint.server
```

### Method 3: Direct Script Execution
```bash
python src/mcp_sharepoint/server.py
```

## 5. Testing the Server

### Using MCP Inspector
```bash
# Start inspector with server
npx @modelcontextprotocol/inspector -- python -m mcp_sharepoint.server
```

Then open browser to `http://localhost:6274` (or the URL shown in terminal).

### Test Authentication Directly (Test Office365 REST API (Legacy))
Create `test_auth.py`:
```python
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.client_credential import ClientCredential
import os
from dotenv import load_dotenv

load_dotenv()
credentials = ClientCredential(os.getenv('SHP_ID_APP'), os.getenv('SHP_ID_APP_SECRET'))
ctx = ClientContext(os.getenv('SHP_SITE_URL')).with_credentials(credentials)

try:
    web = ctx.web.get().execute_query()
    print(f'Connected successfully to: {web.title}')
except Exception as e:
    print(f'Authentication failed: {e}')
```

Run test:
```bash
python test_auth.py
```

## 6. Troubleshooting

The test_auth.py threw an error: "Authentication failed: (None, None, '401 Client Error: Unauthorized for url: https://wisersolutionsinc.sharepoint.com/sites/M365CLI/_api/Web')". 

### Common Issues

**401 Unauthorized Error:**
- Check Azure app permissions are granted
- Verify admin consent was given
- Ensure client secret hasn't expired
- Confirm SharePoint site URL is correct

**No Tools in Inspector:**
- Verify `.env` file has correct credentials
- Check server logs for authentication errors
- Restart inspector after fixing credentials

**Module Import Errors:**
- Ensure virtual environment is activated
- Verify all dependencies are installed with `pip install -e .`

### File Structure
```
wiser-mcp-sharepoint/
├── .env                    # Your credentials (git ignored)
├── .env.example           # Template
├── pyproject.toml         # Package configuration
├── src/
│   └── mcp_sharepoint/
│       ├── __init__.py
│       ├── server.py      # Main server entry point
│       ├── common.py      # Configuration and setup
│       ├── tools.py       # MCP tools
│       ├── resources.py   # MCP resources
│       └── mcp_sharepoint.egg-info/  # Package metadata (auto-generated)
├── venv/                  # Virtual environment
└── test_auth.py          # Authentication test script
```

## 7. Testing Graph API Authentication

This MCP server supports two authentication approaches:

### 1. Office365 REST API (Legacy)
The original implementation using `Office365-REST-Python-Client` library.

### 2. Microsoft Graph API (Recommended) 🆕
A modern, more reliable approach using Microsoft Graph API with MSAL authentication.

**Why Graph API is better:**
- ✅ Better compatibility with Azure AD app registrations
- ✅ More reliable authentication (no 401 errors)
- ✅ Modern Microsoft authentication library (MSAL)
- ✅ Consistent API endpoints and error handling
- ✅ Better documentation and support

### How Microsoft Graph API Works

#### Components Required
1. **MSAL (Microsoft Authentication Library)** - Handles OAuth2 authentication
2. **Azure AD App Registration** - Your app's identity in Microsoft's system
3. **Access Tokens** - Temporary credentials for API calls
4. **Graph API Endpoints** - RESTful URLs for SharePoint operations

### Authentication Flow

**Step-by-step:**
1. **App Registration**: Your app is registered in Azure AD with specific permissions
2. **Client Credentials Flow**: App uses Client ID + Secret to prove its identity
3. **Token Acquisition**: MSAL exchanges credentials for an access token
4. **API Calls**: Token is used in HTTP headers to authenticate SharePoint requests

### Request Flow Example
```python
# 1. Get access token
app = ConfidentialClientApplication(client_id, authority, client_credential=secret)
result = app.acquire_token_for_client(scopes=["[https://graph.microsoft.com/.default](https://graph.microsoft.com/.default)"])
access_token = result["access_token"]

# 2. Make API call
headers = {"Authorization": f"Bearer {access_token}"}
response = requests.get("[https://graph.microsoft.com/v1.0/sites/{site-id}/drives](https://graph.microsoft.com/v1.0/sites/{site-id}/drives)", headers=headers)

# 3. Process response
data = response.json()
```

### Test Graph API Authentication
Create `test_graph_auth.py`.

### Test Comprehensive SharePoint Operations
Create `test_graph_operations.py`.

### Create requirements_graph.txt for Graph API
Create `requirements_graph.txt`.

## 8. Migration to Graph API

If you're currently using Office365 REST API and want to migrate to Graph API:

1. **Test Graph API first**: Run [test_graph_auth.py](cci:7://file:///Users/yihan/Documents/sharepoint%20mcp/wiser-mcp-sharepoint/test_graph_auth.py:0:0-0:0) to ensure it works
2. **Update dependencies**: Ensure `msal` and `requests` are installed  
3. **Modify server code**: Update [common.py](cci:7://file:///Users/yihan/Documents/sharepoint%20mcp/wiser-mcp-sharepoint/src/mcp_sharepoint/common.py:0:0-0:0), [resources.py](cci:7://file:///Users/yihan/Documents/sharepoint%20mcp/wiser-mcp-sharepoint/src/mcp_sharepoint/resources.py:0:0-0:0), and [tools.py](cci:7://file:///Users/yihan/Documents/sharepoint%20mcp/wiser-mcp-sharepoint/src/mcp_sharepoint/tools.py:0:0-0:0) to use Graph API
4. **Test thoroughly**: Use [test_graph_operations.py](cci:7://file:///Users/yihan/Documents/sharepoint%20mcp/wiser-mcp-sharepoint/test_graph_operations.py:0:0-0:0) to verify all operations work
