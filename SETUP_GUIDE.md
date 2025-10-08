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
Reference: 
- https://github.com/AzureAD/microsoft-authentication-library-for-python/blob/dev/sample/confidential_client_sample.py
- https://msal-python.readthedocs.io/en/latest/#msal.ConfidentialClientApplication.acquire_token_for_client



### Authentication Flow
You are acquiring a token for the confidential client (Azure App is indeed a daemon app), not for a signed-in user.
What You're Using: Client Credentials Flow

**Step-by-step:**
1. **App Registration**: Your app is registered in Azure AD (now called Microsoft Entra ID) with specific permissions
2. **Client Credentials Flow**: App uses Client ID + Secret to prove its identity
3. **Token Acquisition**: MSAL exchanges credentials for an access token
4. **API Calls**: Token is used in HTTP headers to authenticate SharePoint requests

#### In Simple Words,
MSAL is essentially the "authentication middleman" that converts your app credentials into a usable access token for Microsoft Graph API calls!

**Step 1: MSAL Gets Access Token**
```python
# Get access token
app = ConfidentialClientApplication(CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET)
result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
access_token = result["access_token"]
```
What MSAL does:

1. Takes your Azure App credentials (CLIENT_ID + CLIENT_SECRET)
2. Contacts Microsoft's authentication servers (login.microsoftonline.com)
3. Says: "This app wants to access Microsoft Graph API"
4. Microsoft validates your app and returns an access token

The **authority** is the URL of the Microsoft identity server that will authenticate your app.
```python
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
```
What it does:
Authentication Server: This is where MSAL sends your credentials to get tokens
Tenant-Specific: Each organization has its own tenant ID, so the authority points to your specific organization's authentication endpoint
Trust Boundary: It defines which Microsoft identity provider to trust

**Scopes** define what permissions your app is requesting - what it wants to access and what it wants to do.
```python
SCOPES = ["https://graph.microsoft.com/.default"]
```
What .default means:
When you use https://graph.microsoft.com/.default, it requests ALL the permissions configured for your app in Azure AD, not just Graph API permissions. So The access token you get contains permissions for all the APIs your app is registered for.
```python
# This works because your token has Graph permissions
response = requests.get("https://graph.microsoft.com/v1.0/sites/{site-id}", headers=headers)

# This would also work because your token has SharePoint permissions
response = requests.get("https://yourtenant.sharepoint.com/_api/web", headers=headers)
```

**Step 2: Use Access Token to Call Graph API**
```python
# Make API call
headers = {"Authorization": f"Bearer {access_token}"}
response = requests.get("https://graph.microsoft.com/v1.0/sites/{site-id}", headers=headers)

# Process response
data = response.json()
```
What happens:

1. Takes your access token
2. Contacts Microsoft Graph API servers
3. Sends your request with the access token in the header
4. Returns the response from Microsoft Graph API

An **endpoint** is a specific URL that represents a particular function or resource you can access via an API.

**The Complete Authentication Flow:**
Your App → MSAL → Microsoft Auth → Access Token → Graph API → SharePoint
   ↓         ↓         ↓              ↓            ↓           ↓
CLIENT_ID  Handles   Validates     Returns      Accepts    Returns
SECRET     OAuth2    App Creds     Token        Token      Data

**Why This Two-Step Process?**
1. Security: Your app credentials never go directly to SharePoint
2. Standardization: Same token works for all Microsoft services (Graph, SharePoint, Teams, etc.)
3. Token Management: MSAL handles token expiration, refresh, caching automatically
4. Permissions: Token contains exactly what permissions your app has


#### Test Graph API Authentication
Create `test_graph_auth.py`.

#### Test Comprehensive SharePoint Operations
Create `test_graph_operations.py`.

#### Create requirements_graph.txt for Graph API
Create `requirements_graph.txt`.


### Network Flow
Your MCP Server → Internet → Microsoft Graph API → SharePoint Online
      ↓                                ↓               ↓
   requests.get()              Validates token    Gets site info
      ↓                                ↓               ↓
   HTTP Request                Returns JSON       Site data

What Happens:
Your app → Graph API server: HTTP GET request
Graph API validates your access token
Graph API → SharePoint backend: Gets site information
SharePoint → Graph API: Returns site data
Graph API → Your app: HTTP response with JSON data

```python
graph_url = f"https://graph.microsoft.com/v1.0/sites/{tenant_name}.sharepoint.com:/sites/{site_name}"
#            ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^ ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
#            |                                |
#            └── Graph API base URL           └── SharePoint site identifier
```

## 8. Migration to Graph API

### Complete Migration Process (October 2024)

Successfully migrated the SharePoint MCP server from Office365 REST API to Microsoft Graph API. The Graph API approach resolves authentication issues and provides better reliability.

#### Migration Steps Completed:

##### 1. **Updated common.py**
- ✅ **Removed**: `office365` library imports (`ClientContext`, `ClientCredential`)
- ✅ **Added**: `msal` library (`ConfidentialClientApplication`) and `requests`
- ✅ **New Functions**:
  - `get_access_token()` - MSAL token acquisition with caching
  - `get_site_info()` - Get SharePoint site ID using Graph API
  - `get_drives()` - Get document libraries (drives)
  - `initialize_sharepoint()` - Complete initialization sequence
  - `make_graph_request()` - Generic Graph API request function

##### 2. **Enhanced Authentication**
- ✅ **MSAL Integration**: Uses `ConfidentialClientApplication` for secure authentication
- ✅ **Token Caching**: Automatic token caching and refresh
- ✅ **Global Variables**: `ACCESS_TOKEN`, `SITE_ID`, `DRIVE_ID` for shared state
- ✅ **Automatic Initialization**: SharePoint connection setup on module import

##### 3. **Updated resources.py**
- ✅ **Core Functions**: `list_folders()`, `list_documents()`, `get_document_content()`
- ✅ **Graph API Endpoints**: Uses modern `/sites/{site-id}/drives/{drive-id}` endpoints
- ✅ **Error Handling**: Comprehensive logging and error management

##### 4. **Key Technical Details**

**Global Variables Usage:**
```python
global ACCESS_TOKEN  # ← Declares intent to modify the global variable
ACCESS_TOKEN = result["access_token"]  # ← Modifies the global ACCESS_TOKEN
```

**Module Import Flow:**
```
1. Python loads common.py
2. Executes all top-level code
3. Reaches: if not initialize_sharepoint():
4. Calls initialize_sharepoint()
5. Calls get_access_token() → sets global ACCESS_TOKEN
6. Calls get_site_info() → sets global SITE_ID  
7. Calls get_drives() → sets global DRIVE_ID
8. Module is ready for use
```

**HTTP Request Pattern:**
```python
# Making one HTTP GET request to the Graph API endpoint for SharePoint site
response = requests.get(graph_url, headers=headers)
```

##### 5. **Benefits Achieved**
- ✅ **Reliability**: Resolves 401 authentication errors from Office365 REST API
- ✅ **Modern API**: Uses Microsoft's recommended Graph API
- ✅ **Better Error Handling**: Clear HTTP status codes and error messages
- ✅ **Token Management**: Automatic token caching and refresh
- ✅ **Simplified Architecture**: Generic request function for all operations


#### Files Modified:
- ✅ `src/mcp_sharepoint/common.py` - Complete Graph API migration
- ✅ `src/mcp_sharepoint/resources.py` - Updated to use Graph API (includes `download_document()`)
- ✅ `src/mcp_sharepoint/tools.py` - **Complete Graph API migration** 
- 🔄 `src/mcp_sharepoint/resources_old.py` - Backup of original resources.py
- 🔄 `src/mcp_sharepoint/tools_old.py` - Backup of original tools.py

#### Testing Infrastructure:
- ✅ `test_resources.py` - Integration/manual testing script
- ✅ `test_resources_unit.py` - Automated unit testing with pytest
- ✅ `test_download_excel.py` - **Excel download and verification test**

#### Migration Completed:
- ✅ **All core files migrated** to Microsoft Graph API
- ✅ **Authentication working** - MSAL token acquisition successful
- ✅ **File operations tested** - Download, content extraction, Excel verification
- ✅ **Real data validation** - Successfully downloaded and read 106×18 Excel dataset

#### Next Steps:
- Update `pyproject.toml` dependencies (add `msal`, remove `office365-rest-python-client`)
- Add remaining tools (upload_document_from_path, update_document, delete_folder)
- Complete integration testing of all MCP server operations

## 9. Testing the Graph API Migration (resources.py)

### Test Scripts Overview

Two comprehensive test scripts have been created to validate the Graph API migration:

#### **test_resources.py** - Integration/Manual Testing
**Purpose:** Human-readable testing for debugging and exploration

**Features:**
- ✅ **Visual Output**: Emojis and formatted output for easy reading
- ✅ **Real Data Display**: Shows actual SharePoint folders, files, and content
- ✅ **Error Details**: Comprehensive error reporting with full tracebacks
- ✅ **Excel Testing**: Dedicated function for testing Excel file content extraction

**Usage:**
```bash
python test_resources.py
```

**Sample Output:**
```
🚀 Testing SharePoint Graph API Resources

🔍 Testing Basic Connection...
ACCESS_TOKEN: ✅ Set
SITE_ID: wisersolutionsinc.sharepoint.com,99b9eb9e-666a-407b...

📁 Testing list_folders()...
✅ Found 3 folders:
  - AI Generated Reports
  - Data
  - Reports

📖 Testing get_document_content()...
Testing with document: 'test_file.txt'
✅ Successfully retrieved content:
  - Type: text
  - Size: 21 bytes
  - Preview: Hello from Graph API!
```

#### **test_resources_unit.py** - Automated Unit Testing
**Purpose:** Automated validation for CI/CD and regression testing

**Features:**
- ✅ **Pytest Integration**: Standard Python testing framework
- ✅ **Pass/Fail Assertions**: Clear success/failure indicators
- ✅ **Automated Execution**: Can be run in CI/CD pipelines
- ✅ **Structure Validation**: Tests data types and required fields

**Usage:**
```bash
# Install pytest if not already installed
pip install pytest

# Run unit tests
python test_resources_unit.py
```

**Sample Output:**
```
================================================================= test session starts =================================================================
test_resources_unit.py::TestSharePointResources::test_connection_variables PASSED [ 25%]
test_resources_unit.py::TestSharePointResources::test_list_folders PASSED         [ 50%]
test_resources_unit.py::TestSharePointResources::test_list_documents PASSED       [ 75%]
test_resources_unit.py::TestSharePointResources::test_get_document_content_nonexistent PASSED [100%]
============================================================ 4 passed, 1 warning in 0.90s =============================================================
```

### Test Functions Available

#### **Basic Functions (Both Scripts):**
- `test_connection_variables()` - Validates ACCESS_TOKEN, SITE_ID, DRIVE_ID are set
- `test_list_folders()` - Tests folder listing functionality
- `test_list_documents()` - Tests document listing in folders
- `test_get_document_content()` - Tests content retrieval from files

#### **Excel-Specific Testing (test_resources.py):**
- `test_get_excel_content(folder_name, file_name)` - Tests Excel file content extraction

**Example Excel Test:**
```python
# Test specific Excel file
test_get_excel_content('Data', '2023 Recruiting Dataset.xlsx')
```

### When to Use Each Test

| Scenario | Use test_resources.py | Use test_resources_unit.py |
|----------|----------------------|---------------------------|
| **Debugging Issues** | ✅ Shows detailed error info | ❌ Only pass/fail |
| **Exploring SharePoint** | ✅ Shows actual data | ❌ Structure only |
| **CI/CD Pipeline** | ❌ Too verbose | ✅ Clean pass/fail |
| **Development** | ✅ Human-readable output | ❌ Minimal output |
| **Regression Testing** | ❌ Manual review needed | ✅ Automated validation |

### Running Tests

#### **Quick Validation:**
```bash
# Run integration tests (shows actual data)
python test_resources.py

# Run unit tests (automated validation)
python test_resources_unit.py
```

#### **Development Workflow:**
```bash
# 1. First, explore and debug with integration tests
python test_resources.py

# 2. Then validate with unit tests
python test_resources_unit.py

# 3. Both should pass before committing code
```

### Test Results Validation

#### **Expected Success Indicators:**
- ✅ All connection variables are set
- ✅ Folders can be listed from SharePoint
- ✅ Documents can be listed from folders
- ✅ File content can be retrieved
- ✅ Excel files can be processed (if available)

#### **Common Issues and Solutions:**
- **401 Errors**: Check Azure app permissions and tenant ID
- **Empty Results**: Verify SharePoint site URL and folder names
- **Token Issues**: Check app ID and secret in `.env` file
- **Excel Processing Errors**: Ensure pandas and openpyxl are installed

## 10. Tools Migration and Excel Testing

### Complete tools.py Migration (October 2024)

Successfully migrated all MCP tools from Office365 REST API to Microsoft Graph API.

#### Migration Process:

##### 1. **File Backup and Replacement**
```bash
# Backup original tools
mv src/mcp_sharepoint/tools.py src/mcp_sharepoint/tools_old.py

# Replace with Graph API version
mv src/mcp_sharepoint/tools_graph.py src/mcp_sharepoint/tools.py
```

##### 2. **Updated Imports**
```python
# Old (Office365 REST API)
from .common import logger, mcp, SHP_DOC_LIBRARY, sp_context
from .resources import list_folders, list_documents, get_document_content, get_folder_tree, download_document

# New (Graph API)
from .common import logger, mcp, ACCESS_TOKEN, SITE_ID, DRIVE_ID, make_graph_request
from .resources import list_folders, list_documents, get_document_content, download_document
```

##### 3. **Tools Migrated to Graph API**
- ✅ **`list_folders_tool()`** - Lists SharePoint folders
- ✅ **`list_documents_tool()`** - Lists documents in folders
- ✅ **`get_document_content_tool()`** - Retrieves file content
- ✅ **`download_document_tool()`** - Downloads files to local filesystem
- ✅ **`create_folder()`** - Creates new folders using Graph API
- ✅ **`upload_document()`** - Uploads files using Graph API
- ✅ **`delete_document()`** - Deletes files using Graph API

##### 4. **Graph API Endpoints Used**
```python
# Folder operations
f"sites/{SITE_ID}/drives/{DRIVE_ID}/root/children"
f"sites/{SITE_ID}/drives/{DRIVE_ID}/root:/{folder_path}:/children"

# File operations
f"sites/{SITE_ID}/drives/{DRIVE_ID}/root:/{file_path}:/content"
f"sites/{SITE_ID}/drives/{DRIVE_ID}/root:/{file_path}"
```

### Excel Download Testing

#### Test Script: `test_download_excel.py`

Created comprehensive test to verify Excel file download and data integrity.

**Features:**
- ✅ **Smart File Discovery** - Automatically finds Excel files in SharePoint
- ✅ **Real Download Test** - Downloads actual Excel files using Graph API
- ✅ **Data Validation** - Reads Excel with pandas to verify integrity
- ✅ **Content Analysis** - Shows actual spreadsheet data, not binary

**Usage:**
```bash
python test_download_excel.py
```

#### Successful Test Results (October 7, 2025):

**Connection:**
- ✅ **Authentication**: MSAL token acquisition successful
- ✅ **Site Access**: Connected to SharePoint site "M365CLI"
- ✅ **Drive Access**: Found Documents library

**File Discovery:**
- ✅ **Found 5 Excel files** in Data folder
- ✅ **Target file**: `2023 Recruiting Dataset  .xlsx` (27,927 bytes)

**Download Verification:**
- ✅ **Download successful**: File saved to Desktop with timestamp
- ✅ **File integrity**: Local file size matches (27,927 bytes)
- ✅ **Excel validation**: Successfully read with pandas

**Data Analysis Results:**
```
✅ Excel file successfully read!
  - Shape: 106 rows × 18 columns
  - Columns: ['Role Name ', 'Job Code ', 'Recruiter', 'Hiring Manger ', ...]
  
📋 First 5 rows of data:
Role Name                                    Job Code  Recruiter   Hiring Manger  
Assistant Controller (Sr Manager, Accounting)    NaN     Tiesa            Penny
BDR (France)                                     NaN   Melissa       ANTONIETTA
Infrastructure Engineer (Vincent backfill)   ENG_081   Melissa Andrew Kesterson
...

📈 Basic statistics for numeric columns:
       Total Days Open   Inbound Applicants  Recruiter Screen
count         75.000000           58.000000         57.000000
mean          74.506667          684.379310         19.491228
std           65.081316          761.254386         19.553154
...
```

#### Key Achievements:

##### 1. **Proves Graph API Reliability**
- ✅ **No 401 errors** - Authentication works consistently
- ✅ **Real data access** - Successfully downloaded recruiting dataset
- ✅ **File integrity** - Excel files are complete and readable

##### 2. **Data Validation Success**
- ✅ **Not binary garbage** - Pandas successfully parsed Excel structure
- ✅ **Structured data** - 106 rows × 18 columns of recruiting data
- ✅ **Mixed data types** - Text, numbers, dates all properly handled
- ✅ **Business data** - Real recruiting metrics (days open, applicants, etc.)

##### 3. **Production Ready**
- ✅ **Error handling** - Comprehensive error reporting and debugging
- ✅ **File management** - Automatic directory creation and cleanup
- ✅ **Flexible naming** - Original filename preserved with timestamp

### Migration Status Summary

#### Completed ✅
- **Authentication**: MSAL-based Graph API authentication
- **Core Resources**: Folder/document listing, content retrieval, download
- **MCP Tools**: All basic SharePoint operations migrated
- **Testing**: Comprehensive test suite with real data validation
- **Documentation**: Complete migration guide and troubleshooting

#### Remaining Tasks 🔄
- **Additional Tools**: `upload_document_from_path()`, `update_document()`, `delete_folder()`
- **Dependencies**: Update `pyproject.toml` (add `msal`, remove `office365-rest-python-client`)
- **Integration Testing**: Full MCP server testing with all operations

The Graph API migration is **functionally complete** and **production ready** for core SharePoint operations! 🎉

## 11. Windsurf Integration and Troubleshooting

### MCP Tools Testing Script

Added comprehensive test script `test_mcp_tools.py` to validate all MCP SharePoint tools:

**Features:**
- ✅ **Complete tool coverage** - Tests all 7 MCP tools (list folders, documents, content, download, create, upload, delete)
- ✅ **End-to-end workflow** - Creates test folder, uploads file, downloads it, verifies content, cleans up
- ✅ **Safe testing** - Uses timestamped test files, non-destructive operations
- ✅ **Real SharePoint interaction** - Tests against actual SharePoint site

**Usage:**
```bash
python test_mcp_tools.py
```

### Windsurf Integration Issues and Solutions

#### Issue 1: Relative Import Error

**Error encountered:**
```
Error: failed to initialize server: ImportError: attempted relative import with no known parent package
```

**Root cause:** Server.py used relative imports (`from .common import logger`) which don't work when MCP clients execute the module.

**Solution:** Changed to explicit imports in `server.py`:
```python
# OLD (relative imports)
from .common import logger, mcp
from . import resources, tools

# NEW (explicit imports) 
from mcp_sharepoint.common import logger, mcp
import mcp_sharepoint.resources as resources
import mcp_sharepoint.tools as tools
```

**Files modified:**
- ✅ `src/mcp_sharepoint/server_old.py` - Backup of original server
- ✅ `src/mcp_sharepoint/server.py` - Updated with explicit imports

#### Issue 2: Read-only File System Error

**Error encountered:**
```
OSError: [Errno 30] Read-only file system: '/mcp_sharepoint.log'
```

**Root cause:** Logging configuration tried to write log file to root filesystem (`/mcp_sharepoint.log`) which is read-only in MCP server environments.

**Solution:** Simplified logging configuration in `common.py`:
```python
# OLD (with file handler)
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[logging.FileHandler('mcp_sharepoint.log'), logging.StreamHandler()]
)

# NEW (console only)
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
```

**Result:** Server now uses default console logging, avoiding filesystem permission issues.

### Final Windsurf Configuration

**Working configuration for `settings.json`:**

#### Option 1: Direct Script Execution (Recommended)
```json
{
  "mcp.servers": {
    "sharepoint-mcp": {
      "command": "/Users/yihan/Documents/sharepoint mcp/wiser-mcp-sharepoint/venv/bin/python",
      "args": [
        "/Users/yihan/Documents/sharepoint mcp/wiser-mcp-sharepoint/src/mcp_sharepoint/server.py"
      ],
      "env": {
        "# env variables are set in .env file"
      }
    }
  }
}
```

**Benefits of this approach:**
- ✅ **No .env file dependency** - All configuration in Windsurf settings
- ✅ **Direct script execution** - Runs server.py directly, avoiding module import issues
- ✅ **Environment isolation** - Each MCP server has its own environment variables
- ✅ **Explicit configuration** - All settings visible in Windsurf config

#### Option 2: Module Execution (Alternative)
```json
{
  "mcp.servers": {
    "sharepoint-mcp": {
      "command": "/Users/yihan/Documents/sharepoint mcp/wiser-mcp-sharepoint/venv/bin/python",
      "args": ["-m", "mcp_sharepoint.server"],
      "cwd": "/Users/yihan/Documents/sharepoint mcp/wiser-mcp-sharepoint",
      "env": {
        "PYTHONPATH": "/Users/yihan/Documents/sharepoint mcp/wiser-mcp-sharepoint/src"
      }
    }
  }
}
```

**Note:** This approach requires the `.env` file to be present and readable.

### Key Lessons Learned

#### 1. **MCP Server Import Requirements**
- ✅ **Use explicit imports** - Relative imports fail in MCP client environments
- ✅ **Test module execution** - Always test with `python -m package.server`
- ✅ **PYTHONPATH configuration** - Essential for module discovery

#### 2. **Logging Best Practices for MCP**
- ✅ **Console-only logging** - Avoid file handlers in MCP servers
- ✅ **Minimal configuration** - Use basic logging setup
- ✅ **Permission awareness** - MCP servers run in restricted environments

#### 3. **Development Workflow**
- ✅ **Backup original files** - Keep `_old.py` versions for reference
- ✅ **Test incrementally** - Fix one issue at a time
- ✅ **Document solutions** - Record fixes for future reference

### Troubleshooting Checklist

**If MCP server fails to start:**

1. **Check imports:**
   ```bash
   cd "/Users/yihan/Documents/sharepoint mcp/wiser-mcp-sharepoint"
   python -c "from src.mcp_sharepoint.server import main; print('✅ Imports work')"
   ```

2. **Test module execution:**
   ```bash
   python -m mcp_sharepoint.server
   ```

3. **Verify environment variables:**
   ```bash
   python -c "from src.mcp_sharepoint.common import SHP_SITE_URL; print(f'Site: {SHP_SITE_URL}')"
   ```

4. **Check virtual environment:**
   ```bash
   which python  # Should show venv path
   pip list | grep -E "(msal|mcp|requests)"  # Check required packages
   ```

The MCP server is now **fully compatible with Windsurf** and ready for production use! 🎯✅
