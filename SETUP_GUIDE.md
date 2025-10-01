# SharePoint MCP Server Setup Guide

This guide documents the complete setup process for the SharePoint MCP server.

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

### Test Authentication Directly
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

## 7. Next Steps

Once authentication is working:
1. Test SharePoint tools in MCP Inspector
2. Integrate with Claude Desktop or other MCP clients
3. Explore available SharePoint operations

## Notes

- The `.egg-info` folder is created during `pip install -e .` and is normal
- Always keep your `.env` file secure and never commit it to git
- Server requires valid SharePoint credentials to start properly
