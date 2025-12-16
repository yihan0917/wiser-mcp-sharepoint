# Wiser SharePoint MCP Server

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/licenses/MIT)

A powerful MCP Server for Microsoft SharePoint integration with advanced HR analytics and reporting capabilities. This server enables AI agents like Claude and Windsurf to interact with SharePoint documents, analyze HR data, and generate professional reports.

## 🎯 What This Server Does

This MCP server connects AI agents to your SharePoint workspace, allowing them to:
- 📁 Browse and manage SharePoint folders and documents
- 📊 Analyze HR recruiting data from Excel files
- 📈 Generate professional PowerPoint presentations with insights
- 🔍 Search through organizational context and documentation
- 📝 Extract and process content from various file types (PDF, Word, Excel, PowerPoint)

## ✨ Key Features

### � **SharePoint Document Management**
- List, create, and delete folders
- Upload, download, and manage documents
- Intelligent text extraction from PDFs, Word docs, Excel, and PowerPoint files

### 📊 **HR Analytics & Reporting**
- Validate data quality in HR Excel files
- Calculate comprehensive recruiting metrics (time-to-hire, fill rates, etc.)
- Analyze hiring trends, bottlenecks, and performance
- Generate chart data for visualizations

### 📈 **Professional Report Generation**
- Create PowerPoint presentations with automated charts and AI-generated insights
- Create Excel files with embedded charts
- Customizable templates and professional styling

### 🔍 **Context Management**
- Search organizational documentation and policies
- Access role descriptions and hiring guidelines
- Query column definitions for data analysis

## 🚀 Getting Started

Follow these steps to set up and use the MCP server with your AI agent (Claude, Windsurf, etc.).

### Prerequisites

- **Python 3.10 or higher** installed on your computer
- **Microsoft Azure account** with access to create app registrations
- **SharePoint site** where you want the MCP server to access documents

### Step 1: Get SharePoint Credentials

You need to register an application in Azure to get credentials for SharePoint access.

1. **Go to Azure Portal**: Visit [portal.azure.com](https://portal.azure.com)
2. **Navigate to App Registrations**: Search for "App registrations" in the top search bar
3. **Create New Registration**:
   - Click "New registration"
   - Name: "SharePoint MCP Server" (or any name you prefer)
   - Supported account types: "Accounts in this organizational directory only"
   - Click "Register"

4. **Get Your Credentials**:
   - **Tenant ID**: Copy from the Overview page
   - **Client ID (Application ID)**: Copy from the Overview page
   - **Client Secret**: 
     - Go to "Certificates & secrets" → "New client secret"
     - Add description, set expiration
     - **Copy the secret value immediately** (you won't see it again!)

5. **Set SharePoint Permissions**:
   - Go to "API permissions" → "Add a permission"
   - Select "SharePoint" → "Application permissions"
   - Add these permissions:
     - `Sites.ReadWrite.All`
     - `Files.ReadWrite.All`
   - Click "Grant admin consent" (requires admin rights)

6. **Get Your SharePoint Site URL**:
   - Go to your SharePoint site in a browser
   - Copy the URL (e.g., `https://yourcompany.sharepoint.com/sites/YourSite`)

### Step 2: Clone and Install

1. **Clone the repository**:
   ```bash
   git clone https://github.com/yihan0917/wiser-mcp-sharepoint.git
   cd wiser-mcp-sharepoint
   ```

2. **Create a virtual environment** (recommended):
   ```bash
   python -m venv .venv
   
   # On macOS/Linux:
   source .venv/bin/activate
   
   # On Windows:
   .venv\Scripts\activate
   ```

3. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

### Step 3: Configure Environment Variables

Create a `.env` file in the project root directory with your credentials:

```bash
# Required SharePoint Credentials
SHP_ID_APP=your-client-id-here
SHP_ID_APP_SECRET=your-client-secret-here
SHP_TENANT_ID=your-tenant-id-here
SHP_SITE_URL=https://yourcompany.sharepoint.com/sites/YourSite
SHP_DOC_LIBRARY=Shared Documents
```

**Important**: Replace the placeholder values with your actual credentials from Step 1.

### Step 4: Test the Server

Test that the server runs correctly:

```bash
python -m mcp_sharepoint
```

If successful, you should see log messages indicating the server has started.

### Step 5: Connect to Your AI Agent

#### For Claude Desktop

1. **Open Settings** → **Developer** → **Edit Config**. This will point you to the MCP config file 'claude_desktop_config.json'. Open it with a text editor.

2. **Edit the config file** and add this MCP server configuration:

```json
{
  "mcpServers": {
    "wiser-sharepoint": {
      "args": [
        "/Users/username/full/path/to/wiser-mcp-sharepoint/src/mcp_sharepoint/server.py"
      ],
      "command": "/Users/username/full/path/to/wiser-mcp-sharepoint/venv/bin/python",
      "env": {
        "SHP_ID_APP": "your-client-id",
        "SHP_ID_APP_SECRET": "your-client-secret",
        "SHP_TENANT_ID": "your-tenant-id",
        "SHP_SITE_URL": "https://yourcompany.sharepoint.com/sites/YourSite",
        "SHP_DOC_LIBRARY": "Shared Documents"
      }
    }
  }
}
```

**Replace** `/full/path/to/wiser-mcp-sharepoint` with the actual path where you cloned the repository

3. **Save and Restart Claude Desktop**

#### For Windsurf

1. **Open Windsurf** → **MCP Marketplace**
2. **Click on the settings icon (gear icon)**. This will open the MCP config file.
3. **Edit the config file** and add this MCP server configuration:

```json
{
  "mcpServers": {
    "wiser-sharepoint": {
      "args": [
        "/Users/username/full/path/to/wiser-mcp-sharepoint/src/mcp_sharepoint/server.py"
      ],
      "command": "/Users/username/full/path/to/wiser-mcp-sharepoint/venv/bin/python",
      "env": {
        "SHP_ID_APP": "your-client-id",
        "SHP_ID_APP_SECRET": "your-client-secret",
        "SHP_TENANT_ID": "your-tenant-id",
        "SHP_SITE_URL": "https://yourcompany.sharepoint.com/sites/YourSite",
        "SHP_DOC_LIBRARY": "Shared Documents"
      }
    }
  }
}
```
**Replace** `/full/path/to/wiser-mcp-sharepoint` with the actual path where you cloned the repository

4. **Save and Restart Windsurf**

### Step 6: Verify Connection

In your AI agent (Claude or Windsurf), try asking:

> "Can you list the folders in my SharePoint?"

If the MCP server is connected correctly, the AI will use the `List_SharePoint_Folders` tool to show your SharePoint folders.

## 🛠️ Available Tools

The server provides **20+ tools** organized into these categories:

### 📁 SharePoint Document Management

| Tool | Description |
|------|-------------|
| `List_SharePoint_Folders` | List all folders in a directory or root |
| `List_SharePoint_Documents` | List all documents in a specific folder |
| `Get_Document_Content` | Get content from a document (with text extraction for PDF, Word, Excel, PowerPoint) |
| `Create_Folder` | Create a new folder in SharePoint |
| `Upload_Document` | Upload a new document to SharePoint |
| `Delete_Document` | Delete a document from SharePoint |
| `Download_Document` | Download a document to your local filesystem |

### 📊 HR Analytics & Data Quality

| Tool | Description |
|------|-------------|
| `Validate_Excel_Data_Quality` | Check data quality in HR Excel files, identify missing data and anomalies |
| `Calculate_HR_Metrics` | Calculate recruiting metrics like time-to-hire, fill rates, and hiring trends |
| `Analyze_HR_File_Complete` | Comprehensive analysis of HR data with insights and recommendations |
| `Generate_Chart_Data` | Generate data formatted for charts and visualizations |

### 📈 Report Generation

| Tool | Description |
|------|-------------|
| `Create_PowerPoint_Report` | Generate professional PowerPoint presentations with automated charts and AI-generated insights |
| `Create_Excel_With_Charts` | Create Excel files with embedded charts from HR data |

### 🔍 Context & Knowledge Management

| Tool | Description |
|------|-------------|
| `Get_Column_Definition` | Get definition for a specific Excel column |
| `Search_Column_Definitions` | Search for column definitions by keyword |
| `Get_All_Column_Definitions` | Get all available column definitions |
| `Get_Matching_Columns` | Get definitions for multiple columns at once |
| `Get_Context_Summary` | Get summary of all loaded context files |
| `Search_All_Context` | Search across all organizational documentation |
| `Search_Specific_Context_File` | Search within a specific context file (e.g., hiring guide, role descriptions) |
| `List_Available_Context_Files` | List all available context files by category |
| `Find_Relevant_Context_Files` | Find context files based on natural language description |

## 💡 Usage Examples

### Example 1: Analyze Recruiting Data

> "Analyze the file 'Q3_Recruiting_Data.xlsx' in the 'Recruiting Data' folder and tell me the average time-to-hire"

The AI will use `Analyze_HR_File_Complete` to process the Excel file and provide insights.

### Example 2: Generate a Report

> "Create a PowerPoint presentation analyzing Q3 recruiting performance from 'Training_Time_In_Step_Q3_97_records.xlsx' in 'Recruiting Data' folder, with insights and recommendations. Save to 'AI Generated Reports' folder."

The AI will use `Create_PowerPoint_Report` to generate a professional presentation with charts and analysis.

### Example 3: Search Documentation

> "What are the requirements for a Senior Software Engineer role?"

The AI will use `Search_All_Context` or `Find_Relevant_Context_Files` to search through role descriptions and provide the information.

### Example 4: Upload a Document

> "Upload this meeting summary as 'Team_Meeting_Notes.docx' to the 'Meeting Notes' folder"

The AI will use `Upload_Document` to save the document to SharePoint.

## 🔧 Troubleshooting

### Server Won't Start

- **Check Python version**: Ensure you have Python 3.10 or higher (`python --version`)
- **Verify dependencies**: Run `pip install -r requirements.txt` again
- **Check .env file**: Ensure all required variables are set correctly

### Authentication Errors

- **Verify credentials**: Double-check your Client ID, Client Secret, and Tenant ID
- **Check permissions**: Ensure your Azure app has the required SharePoint permissions
- **Admin consent**: Make sure admin consent was granted for the API permissions

### Can't Find Files

- **Check SHP_DOC_LIBRARY**: Ensure it matches your SharePoint library path (usually "Shared Documents")
- **Verify folder names**: Folder names are case-sensitive
- **Check permissions**: Ensure your Azure app has access to the SharePoint site

### AI Agent Not Connecting

- **Restart the agent**: Close and reopen Claude Desktop or Windsurf
- **Check config path**: Verify the `cwd` path in your MCP configuration is correct
- **View logs**: Check the AI agent's logs for error messages

## 🏗️ Architecture

The server uses:
- **Microsoft Graph API** for reliable SharePoint access
- **MSAL authentication** for secure credential management
- **Async/await** throughout for non-blocking operations
- **Smart error handling** with decorators for cleaner code
- **Context management system** for intelligent organizational knowledge retrieval

## 🐛 Debugging

For advanced debugging, use the [MCP Inspector](https://github.com/modelcontextprotocol/inspector):

```bash
npx @modelcontextprotocol/inspector python -m mcp_sharepoint
```

This opens a web interface where you can test tools and see detailed request/response logs.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

Copyright (c) 2025 sofias tech