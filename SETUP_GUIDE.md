# SharePoint MCP Server Setup Guide

This guide documents the complete setup process for the SharePoint MCP server with HR Analytics and Interactive Visualizations.

## Installation

### Quick Install (Recommended)

Install the MCP SharePoint server with all dependencies:

```bash
pip install -e .
```

### Manual Dependency Installation

If you prefer to install dependencies manually:

```bash
# Core dependencies
pip install mcp>=1.2.1 msal>=1.24.0 requests>=2.31.0 python-dotenv>=1.0.0

# Document processing
pip install pymupdf>=1.23.0 python-docx>=1.1.0 python-pptx>=0.6.21 openpyxl>=3.1.0

# Analytics and visualization (NEW in v0.2.0)
pip install pandas>=2.0.0 numpy>=1.24.0 plotly>=5.17.0
```

### Installation Verification

After installation, verify all dependencies are available:

```python
import mcp_sharepoint
import plotly
import pandas
import numpy
print("✅ All dependencies installed successfully!")
```

## Version History

### v0.2.0 - HR Analytics & Visualization
- ✅ Interactive web dashboards with Plotly
- ✅ Excel export with embedded charts  
- ✅ Comprehensive HR metrics calculation
- ✅ Data quality validation
- ✅ Chart data generation for 5+ visualization types

### v0.1.6 - Enhanced Document Processing
- ✅ Microsoft Graph API integration
- ✅ Text extraction from PDF, Word, Excel, PowerPoint
- ✅ Column context system for HR data
- ✅ Professional Word document creation

### v0.3.0 - Comprehensive Context Management System
- ✅ **Advanced Context Manager** - Intelligent context loading and injection
- ✅ **22 Context Files** - 123,137 characters of business, HR, and role context
- ✅ **5 Context Categories** - columns, metrics, business, recruiting, roles
- ✅ **Tool-Context Mapping** - Automatic context injection based on tool usage
- ✅ **Role Descriptions** - Complete career paths for Software Engineering, Data Management, Engineering Leadership, ML/DS/DA
- ✅ **Search & Discovery** - Search across all context files, column definition matching
- ✅ **Context-Aware Analytics** - AI provides intelligent recommendations based on company context

### v0.4.0 - Dynamic Column Detection & Context-Driven Analysis
- ✅ **Zero Hardcoded Column Names** - Works with ANY HR dataset structure
- ✅ **Intelligent Column Detection** - Auto-categorizes columns into 9 types (date, time_step, role, location, person, department, numeric, categorical, id)
- ✅ **Pattern-Based Recognition** - Detects column purpose from names and content
- ✅ **Dynamic Metrics Calculation** - Generates appropriate metrics based on available data
- ✅ **Flexible Chart Generation** - Creates visualizations from detected column types
- ✅ **Context-Integrated Insights** - Full 123K chars of context injected into analysis process
- ✅ **Business-Aware Recommendations** - AI references role descriptions, hiring guides, and career paths
- ✅ **Removed Definition Slides** - Context used for intelligent analysis, not just display
- ✅ **5 New Chart Types** - time_by_step, role_distribution, location_distribution, time_distribution, department_distribution

### v0.5.0 - AI-Driven Insights Architecture 🆕
- ✅ **Hybrid Analysis Model** - Pre-calculated metrics as reference + raw data for AI creativity
- ✅ **Enhanced Tool Outputs** - Tools now return structured data + raw summaries + business context
- ✅ **AI-Empowered Analysis** - AI can discover patterns beyond pre-defined metrics
- ✅ **Data Preview Access** - Tools provide 10-15 row samples for AI pattern examination
- ✅ **Statistical Summaries** - Mean, median, std, min, max, quartiles for all numeric columns
- ✅ **Categorical Distributions** - Value counts for categorical columns (top 15-20 values)
- ✅ **Context-Aware Guidance** - Explicit instructions for AI to think creatively
- ✅ **Flexible Recommendations** - AI generates custom insights based on actual data patterns
- ✅ **Business Context Integration** - 2000-3000 chars of context per tool call
- ✅ **Column Definition Matching** - Automatic lookup of column meanings from context files

#### Key Improvements in v0.4.0

**1. Dynamic Column Detection (`analytics_helper.py`)**
```python
def detect_column_types(df) -> Dict[str, List[str]]:
    """Intelligently categorizes columns without hardcoded names"""
    # Returns: {
    #   'time_step_columns': [...],  # Numeric columns with 'time' in name
    #   'role_columns': [...],        # Columns matching role patterns
    #   'location_columns': [...],    # Columns matching location patterns
    #   ...
    # }
```

**2. Context Integration (`tools.py`)**
```python
# Get full context (123K chars) for the tool
tool_context = context_manager.get_context_for_tool('Create_PowerPoint_Report')

# Get column definitions for actual columns in dataset
column_definitions = {col: context_manager.get_column_definition(col) 
                     for col in df.columns}

# Pass context to analysis
analysis = {
    'metrics': metrics,
    'context': tool_context,           # ← Full business context
    'column_definitions': column_definitions,  # ← Column explanations
    'data_summary': {...}
}
```

**3. Context-Aware Insights**
- Insights now reference role descriptions, hiring guides, and career frameworks
- Recommendations aligned with company values and business strategy
- Technical role analysis considers ML/DS/DA position descriptions
- Location analysis references multi-region business strategy

**4. Flexibility Benefits**
- ✅ Works with time-in-step data (e.g., "Time in Application Status: In-Review")
- ✅ Works with traditional HR data (e.g., "Total Days Open")
- ✅ Works with custom column names from any ATS or HRIS system
- ✅ No code changes needed when column names change
- ✅ Automatically adapts to new data structures

#### Key Improvements in v0.5.0

**1. Hybrid Analysis Architecture**

The v0.5.0 release transforms MCP tools from **constrained calculators** to **AI-empowered analysis platforms**:

**Before (Constrained):**
```json
{
  "metrics": {
    "average_time": 75,
    "total_positions": 97
  }
}
```
❌ AI could only report: "Average time is 75 days"

**After (AI-Empowered):**
```json
{
  "pre_calculated_metrics": { /* your metrics as reference */ },
  "data_summary": {
    "data_preview": [/* actual rows */],
    "numeric_summaries": {/* stats for each column */},
    "categorical_distributions": {/* value counts */}
  },
  "context": {
    "column_definitions": {/* what columns mean */},
    "business_context": "/* hiring guide, company info */",
    "analysis_guidance": "/* instructions to think creatively */"
  }
}
```
✅ AI can discover: "Engineering roles average 120 days vs 45 for Sales. Technical Interview step shows high variability (3-20 days). Manager 'John Smith' consistently fills positions 30% faster."

**2. Enhanced Tool Outputs**

Both `Calculate_HR_Metrics` and `Analyze_HR_File_Complete` now return:

- **Pre-calculated metrics** - Your Python code's trusted baseline calculations (as REFERENCE)
- **Raw data summaries** - Statistical summaries, distributions, data preview for AI analysis
- **Business context** - 2000-3000 chars from context manager (hiring guides, role descriptions, etc.)
- **Column definitions** - Automatic lookup of what each column means
- **Analysis guidance** - Explicit instructions for AI to think beyond pre-defined metrics

**3. What This Enables**

**Your metrics provide the foundation. AI builds the insights.**

- ✅ **Consistency:** Pre-calculated metrics ensure accuracy
- ✅ **Creativity:** AI discovers patterns you didn't anticipate  
- ✅ **Context-awareness:** Recommendations aligned with your business practices
- ✅ **Adaptability:** Works with any data structure automatically
- ✅ **Extensibility:** Add new metrics anytime; AI uses them as reference

**Example:**
- Pre-calculated metric: "Average time-to-hire: 95 days"
- AI discovers: "Technical Interview varies 3-20 days (high variability)"
- AI correlates: "Senior Engineering roles take 40% longer at Final Interview"
- AI recommends (context-aware): "Based on your L5+ career path requirements, consider standardized panel interviews"
- AI identifies: "August approvals fill faster than September (vacation impact?)"

**4. Tool-Specific Enhancements**

**`Calculate_HR_Metrics`:**
```python
# Now returns:
{
  "pre_calculated_metrics": {...},      # Your baseline metrics
  "data_summary": {
    "data_preview": [...],              # 10 rows for pattern examination
    "numeric_summaries": {...},         # Stats for all numeric columns
    "categorical_distributions": {...}  # Top 20 values per category
  },
  "context": {
    "column_definitions": {...},        # Column meanings
    "business_context": "...",          # 2000 chars of context
    "analysis_guidance": "..."          # Instructions for AI
  }
}
```

**`Analyze_HR_File_Complete`:**
```python
# Now returns:
{
  "data_summary": {
    "data_preview": [...],              # 15 rows for deeper analysis
    "numeric_summaries": {...},
    "categorical_distributions": {...}
  },
  "pre_calculated_analysis": {
    "data_quality": {...},              # Your validation results
    "hr_metrics": {...},                # Your calculated metrics
    "suggested_charts": [...],          # Your chart suggestions
    "basic_recommendations": [...]      # Your rule-based recommendations
  },
  "context": {
    "column_definitions": {...},
    "business_context": "...",          # 3000 chars of context
    "analysis_guidance": "..."          # Creative thinking instructions
  }
}
```

**5. Documentation**

See `AI_INSIGHTS_ARCHITECTURE.md` for complete details on:
- Architecture philosophy and design principles
- Before/after comparisons with examples
- Real-world usage scenarios
- Best practices for tool development
- Future enhancement suggestions

## Context Management System

### Overview

The SharePoint MCP server includes a comprehensive context management system that provides AI tools with rich business, HR, and organizational context. This enables intelligent, context-aware recommendations and analysis.

### Context Categories

The system organizes context into 5 categories:

#### 1. **Columns** (6,728 chars)
- `column_definitions.md` - 73 Excel column definitions for HR data
- Covers job information, personnel data, application tracking, timing metrics, sourcing, and costs

#### 2. **Metrics** (2,661 chars)
- `metrics_definitions.md` - HR metrics and KPIs
- Time-to-hire, cost-per-hire, source effectiveness, diversity metrics

#### 3. **Business** (13,832 chars)
- `company_overview.md` - Company mission, values, history
- `engineering_overview.md` - Engineering culture, strategy, platforms

#### 4. **Recruiting** (30,093 chars)
- `hiring_guide.md` - Hiring processes and guidelines
- `engineering_career_path.md` - Engineering career progression framework (L1-L10)

#### 5. **Roles** (70,823 chars) - **22 files**
- **Software Engineering**: `software_engineer_role_description.md` (L1-L8)
- **Data Management**: `data_management_role_description.md` (L1-L7, 3 tracks)
- **Engineering Leadership**: `engineering_leadership_role_description.md` (6 levels)
- **ML/DS/DA Index**: `ml_ds_da_role_description.md`
- **ML/DS/DA Detailed** (12 files):
  - `Position-Description-MLE1-DS1.md` - ML Engineer I / Data Scientist I
  - `Position-Description-DA1.md` - Data Analyst I
  - `Position-Description-MLE2-DS2.md` - ML Engineer II / Data Scientist II
  - `Position-Description-DA2.md` - Data Analyst II
  - `Position-Description-SMLE-SDS.md` - Senior ML Engineer / Senior Data Scientist
  - `Position-Description-SDA.md` - Senior Data Analyst
  - `Position-Description-LSMLE-LSDS.md` - Lead Senior ML Engineer / Lead Senior Data Scientist
  - `Position-Description-LSDA.md` - Lead Senior Data Analyst
  - `Position-Description-PMLE-PDS.md` - Principal ML Engineer / Principal Data Scientist
  - `Position-Description-PDA.md` - Principal Data Analyst
  - `Position-Description-SEMLDS Manager.md` - Software Engineering ML/DS Manager
  - `Position-Description-Analytics Manager.md` - Analytics Manager

### Tool-Context Mapping

Tools automatically receive relevant context based on their function:

| Tool | Context Categories | Total Context Size |
|------|-------------------|-------------------|
| **Analyze_HR_File_Complete** | columns, business, metrics, roles | 93,917 chars |
| **Calculate_HR_Metrics** | columns, metrics, recruiting, roles | ~85,000 chars |
| **Validate_Excel_Data_Quality** | columns, recruiting, roles | ~107,000 chars |
| **Create_PowerPoint_Report** | columns, business, metrics, recruiting, roles | ~123,000 chars |
| **Create_Excel_With_Charts** | columns, metrics | ~9,400 chars |
| **Generate_Chart_Data** | columns, metrics | ~9,400 chars |

### Context Files Location

All context files are located in:
```
src/mcp_sharepoint/context/
├── column_definitions.md
├── metrics_definitions.md
├── company_overview.md
├── engineering_overview.md
├── hiring_guide.md
├── engineering_career_path.md                    # ← Renamed from career_path.md
├── software_engineer_role_description.md
├── data_management_role_description.md
├── engineering_leadership_role_description.md
├── ml_ds_da_role_description.md
├── in_store_operations_role_description.md       # ← NEW: Operations roles
├── Position-Description-MLE1-DS1.md
├── Position-Description-DA1.md
├── Position-Description-MLE2-DS2.md
├── Position-Description-DA2.md
├── Position-Description-SMLE-SDS.md
├── Position-Description-SDA.md
├── Position-Description-LSMLE-LSDS.md
├── Position-Description-LSDA.md
├── Position-Description-PMLE-PDS.md
├── Position-Description-PDA.md
├── Position-Description-SEMLDS Manager.md
└── Position-Description-Analytics Manager.md
```

### Testing Context System

Test the context manager integration:

```bash
python test_context_integration.py
```

**Expected output:**
```
✅ All tests passed! Context manager is working correctly.
✓ Total files loaded: 22
✓ Total characters: 123,137
✓ Categories: columns, metrics, business, recruiting, roles
✓ Column definitions: 73
```

### Adding New Context Files

To add new context files:

1. **Create markdown file** in `src/mcp_sharepoint/context/`
2. **Update context_manager.py** - Add file to appropriate category in `CONTEXT_CATEGORIES`
3. **Update tool mapping** (optional) - Add category to tools in `TOOL_CONTEXT_MAP`
4. **Test** - Run `test_context_integration.py` to verify

**Example:**
```python
# In context_manager.py
CONTEXT_CATEGORIES = {
    'roles': [
        'software_engineer_role_description.md',
        'your_new_role_file.md',  # ← Add here
    ]
}
```

### Recent Context Updates (November 2024)

#### 1. **Engineering Career Path Renamed & Reformatted**
- ✅ **Renamed**: `career_path.md` → `engineering_career_path.md`
- ✅ **Department Identifier**: Added clear "DEPARTMENT: ENGINEERING" header
- ✅ **Improved Structure**: Reorganized with proper Markdown hierarchy
  - Overview section (Why career paths, two tracks, leadership team)
  - Important Notes (bullet-pointed for clarity)
  - Career Path Levels (L1-L10 with clear descriptions)
- ✅ **Track Organization**: Titles now grouped by Maker/Architect/Manager tracks
- ✅ **Better Formatting**: Consistent headings, horizontal dividers, bold labels

**Why the rename?**  
The career path framework is specific to Engineering and shouldn't be confused with career paths in other departments (Operations, Sales, etc.).

#### 2. **In-Store Operations Roles Added**
- ✅ **New File**: `in_store_operations_role_description.md`
- ✅ **New Category**: Created 'operations' category separate from 'roles'
- ✅ **Department Clarity**: Clear header stating "DEPARTMENT: IN-STORE OPERATIONS"
- ✅ **Warning Label**: Explicit note to NOT apply these to similarly-titled roles in other departments

**Roles Covered:**
- **Retail Intelligence (RI)**: Data Validation Team (Sr. Data Quality Specialist, Team Leads)
- **User Support**: User Support Associate, Team Leads
- **REM/RI**: Technical Implementation Specialists, Team Leads, Operations Manager
- **ISPC**: Data Collection Specialists, Technical Implementation Specialists, Sr. Technical Operations Analyst, Operations Manager

**Why separate from Engineering roles?**  
To prevent AI confusion when analyzing HR data. A "Team Lead" in Operations has completely different responsibilities than a "Lead Engineer" in Engineering.

#### 3. **Context Manager Updates**
Updated `context_manager.py` to reflect these changes:

```python
CONTEXT_CATEGORIES = {
    'columns': ['column_definitions.md'],
    'metrics': ['metrics_definitions.md'],
    'business': ['company_overview.md', 'engineering_overview.md'],
    'recruiting': ['hiring_guide.md', 'engineering_career_path.md'],  # ← Updated
    'roles': [
        # Engineering roles (22 files)
        'software_engineer_role_description.md',
        'data_management_role_description.md',
        # ... other engineering roles
    ],
    'operations': [  # ← NEW category
        'in_store_operations_role_description.md'
    ]
}

TOOL_CONTEXT_MAP = {
    'Analyze_HR_File_Complete': ['columns', 'business', 'metrics', 'roles', 'operations'],  # ← Added operations
    'Calculate_HR_Metrics': ['columns', 'metrics', 'recruiting', 'roles', 'operations'],
    'Validate_Excel_Data_Quality': ['columns', 'recruiting', 'roles', 'operations'],
    'Create_PowerPoint_Report': ['columns', 'business', 'metrics', 'recruiting', 'roles', 'operations'],
    # ... other tools
}
```

#### 4. **Benefits of These Updates**

**Better Department Separation:**
- Engineering roles clearly labeled as Engineering-specific
- Operations roles clearly labeled as Operations-specific
- Prevents cross-department confusion in AI analysis

**Improved Formatting:**
- Consistent structure across all role description files
- Easy to scan with clear headings and sections
- Professional appearance with proper Markdown formatting

**Enhanced Context Awareness:**
- AI tools now understand department boundaries
- More accurate role matching and analysis
- Better recommendations aligned with specific department needs

### Context Manager Features

- ✅ **Automatic Loading** - All markdown files loaded on initialization
- ✅ **Category-Based Organization** - Context grouped by purpose
- ✅ **Tool-Specific Injection** - Only relevant context sent to each tool
- ✅ **Search Functionality** - Search across all context files
- ✅ **Column Matching** - Match Excel columns to definitions
- ✅ **Context Summary** - Get statistics on loaded context
- ✅ **Backward Compatible** - Works with existing tools without changes

### Benefits

1. **Intelligent Recommendations** - AI understands company values, hiring practices, role expectations
2. **Context-Aware Analysis** - Analysis considers organizational context
3. **Accurate Role Matching** - Validates job titles against known positions
4. **Career Path Guidance** - Provides progression insights for employees
5. **Consistent Terminology** - Uses company-specific definitions and metrics

## Architecture: Dynamic Analysis + Context Integration

### Two-Layer Intelligence System

The SharePoint MCP server uses a two-layer approach for intelligent HR analytics:

#### **Layer 1: Dynamic Data Analysis** (`analytics_helper.py`)
- **No hardcoded column names** - Works with any dataset structure
- **Pattern-based detection** - Identifies column types from names and content
- **Automatic type conversion** - Converts dates, numbers, and text appropriately
- **Flexible metrics** - Calculates metrics based on available columns
- **Adaptive charts** - Generates visualizations from detected data

**Key Methods:**
- `detect_column_types(df)` - Categorizes columns into 9 types
- `calculate_hiring_metrics(df)` - Dynamic metric calculation
- `generate_chart_data(df, chart_type)` - Flexible chart generation

#### **Layer 2: Context-Aware Insights** (`tools.py`)
- **Context injection** - Loads 123K chars of business context
- **Column definition lookup** - Matches columns to known definitions
- **Business-aligned recommendations** - References role descriptions and hiring guides
- **Organizational perspective** - Insights consider company values and strategy

**Integration Flow:**
```
1. Load Excel data → Parse into DataFrame
2. Get tool context → Load relevant context files (123K chars)
3. Get column definitions → Match actual columns to definitions
4. Detect column types → Categorize without hardcoding
5. Calculate metrics → Use detected columns dynamically
6. Generate charts → Create visualizations from available data
7. Generate insights → Combine data patterns + business context
8. Create PowerPoint → Context-aware presentation
```

### Why This Architecture?

**Separation of Concerns:**
- `analytics_helper.py` = Pure data analysis (no business logic)
- `tools.py` = Business logic (combines analysis + context)

**Benefits:**
- ✅ Analytics work with ANY dataset (no column name requirements)
- ✅ Context enriches insights (not required for basic analysis)
- ✅ Easy to test (analytics can be tested independently)
- ✅ Maintainable (context changes don't affect core analytics)
- ✅ Extensible (add new context without changing analytics)

### Example: How It Works

**Your Data:**
```
Columns: "Time in Application Status: In-Review/Recruiter Screen", "Job Title", "Job Country"
```

**Layer 1 (Dynamic Detection):**
```python
detect_column_types(df) returns:
{
  'time_step_columns': ['Time in Application Status: In-Review/Recruiter Screen'],
  'role_columns': ['Job Title'],
  'location_columns': ['Job Country']
}
```

**Layer 2 (Context Integration):**
```python
tool_context includes:
- Position-Description-DA1.md: "Data Analyst I responsibilities..."
- hiring_guide.md: "Recruiting process best practices..."
- career_path.md: "L1-L8 progression framework..."

Insights generated:
"💻 Technical roles: 3 types, 45 positions (46.4%)"
"📋 Hiring aligned with defined career paths and role frameworks"
"🎯 Recommendation: Focus on streamlining 'In-Review/Recruiter Screen' step"
```

## Setup Notes

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
├── SETUP_GUIDE.md         # This comprehensive setup guide
├── AI_INSIGHTS_ARCHITECTURE.md  # v0.5.0 AI-driven insights documentation 🆕
├── src/
│   └── mcp_sharepoint/
│       ├── __init__.py
│       ├── server.py      # Main server entry point
│       ├── common.py      # Configuration and Graph API setup
│       ├── tools.py       # MCP tools (Graph API)
│       ├── resources.py   # MCP resources (Graph API)
│       ├── context_manager.py  # Context management system 🆕
│       ├── analytics_helper.py # HR analytics and metrics
│       ├── context/       # Context files directory 🆕
│       │   ├── column_definitions.md
│       │   ├── metrics_definitions.md
│       │   ├── company_overview.md
│       │   ├── engineering_overview.md
│       │   ├── hiring_guide.md
│       │   ├── engineering_career_path.md  # ← Renamed from career_path.md
│       │   ├── software_engineer_role_description.md
│       │   ├── data_management_role_description.md
│       │   ├── engineering_leadership_role_description.md
│       │   ├── ml_ds_da_role_description.md
│       │   ├── in_store_operations_role_description.md  # ← NEW
│       │   ├── Position-Description-MLE1-DS1.md
│       │   ├── Position-Description-DA1.md
│       │   ├── Position-Description-MLE2-DS2.md
│       │   ├── Position-Description-DA2.md
│       │   ├── Position-Description-SMLE-SDS.md
│       │   ├── Position-Description-SDA.md
│       │   ├── Position-Description-LSMLE-LSDS.md
│       │   ├── Position-Description-LSDA.md
│       │   ├── Position-Description-PMLE-PDS.md
│       │   ├── Position-Description-PDA.md
│       │   ├── Position-Description-SEMLDS Manager.md
│       │   └── Position-Description-Analytics Manager.md
│       └── mcp_sharepoint.egg-info/  # Package metadata (auto-generated)
├── venv/                  # Virtual environment
├── test_auth.py          # Authentication test script
├── test_graph_auth.py    # Graph API authentication test
├── test_resources.py     # Integration testing
├── test_resources_unit.py # Unit testing
├── test_download_excel.py # Excel download test
├── test_mcp_tools.py     # MCP tools testing
└── test_context_integration.py # Context system test 🆕
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

## 12. Text Extraction Enhancement

### Overview

Enhanced the SharePoint MCP server with comprehensive text extraction capabilities, allowing users to retrieve human-readable content from various document types instead of binary data.

### Implementation Details

#### 1. **Enhanced `resources.py` with Text Extraction Functions**

Added specialized text extraction functions for different file types:

```python
def extract_text_from_pdf(pdf_content):
    """Extract text from PDF using PyMuPDF"""
    
def extract_text_from_excel(content_bytes, max_rows_per_sheet=None):
    """Extract text from Excel files with headers and configurable row limit"""
    
def extract_text_from_word(content_bytes):
    """Extract text from Word documents including tables"""
    
def extract_text_from_powerpoint(content_bytes):
    """Extract text from PowerPoint presentations (NEW!)"""
```

#### 2. **Updated File Type Support**

Extended `FILE_TYPES` configuration to include PowerPoint:

```python
FILE_TYPES = {
    'text': ['.txt', '.csv', '.json', '.xml', '.html', '.md', '.js', '.css', '.py'],
    'pdf': ['.pdf'],
    'excel': ['.xlsx', '.xls'],
    'word': ['.docx', '.doc'],
    'powerpoint': ['.pptx', '.ppt']  # NEW!
}
```

#### 3. **Enhanced `get_document_content()` Function**

The function now:
- ✅ **Automatically detects file types** based on extensions
- ✅ **Extracts human-readable text** for supported formats
- ✅ **Returns structured metadata** (page_count, sheet_count, slide_count)
- ✅ **Graceful fallback** to binary content if extraction fails
- ✅ **Maintains compatibility** with existing functionality

#### 4. **Enhanced Excel Text Extraction Features**

The improved Excel extraction function includes:
- 📊 **Column headers included** with "HEADERS:" prefix
- 🔢 **Configurable row limits** via `max_rows_per_sheet` parameter
- 📈 **Complete data extraction** when `max_rows_per_sheet=None` (default)
- 📋 **Visual formatting** with separator lines between headers and data
- 📝 **Row count summaries** showing total rows and truncation info
- 🗂️ **Empty sheet handling** with clear "(Empty sheet)" messages
- 📄 **Sheet separation** with blank lines for better readability

**Usage Examples:**
```python
# Extract all rows with headers (default behavior)
text, sheets = extract_text_from_excel(content_bytes)

# Extract only first 100 rows per sheet
text, sheets = extract_text_from_excel(content_bytes, max_rows_per_sheet=100)

# Extract only first 50 rows per sheet (old behavior equivalent)
text, sheets = extract_text_from_excel(content_bytes, max_rows_per_sheet=50)
```

#### 5. **PowerPoint Text Extraction Features**

The new PowerPoint extraction capability:
- 📊 **Slide-by-slide extraction** with clear separators
- 🔤 **Text from shapes and text boxes**
- 📋 **Table content extraction**
- 📈 **Slide count metadata**
- 🎯 **Formatted output** for easy reading

### Dependencies Added

Updated `pyproject.toml` with required packages:

```toml
dependencies = [
    # ... existing dependencies ...
    "python-pptx>=0.6.21",  # PowerPoint processing
    "msal>=1.24.0",         # Microsoft Graph authentication
    "requests>=2.31.0",     # HTTP requests
]
```

### Usage Examples

#### Before Enhancement (Binary Output)
```json
{
    "name": "report.pptx",
    "content_type": "binary",
    "content_base64": "UEsDBBQABgAIAAAAIQAVYrdQ...",
    "size": 590765
}
```

#### After Enhancement (Text Output)

**PowerPoint Example:**
```json
{
    "name": "report.pptx",
    "content_type": "text",
    "content": "=== Slide 1 ===\nHR Reporting - July 2024\nFor DA Team\n\n=== Slide 2 ===\nKey Metrics\n...",
    "original_type": "powerpoint",
    "slide_count": 5,
    "size": 590765
}
```

**Excel Example (Enhanced with Headers):**
```json
{
    "name": "data.xlsx",
    "content_type": "text",
    "content": "=== Sheet1 ===\nHEADERS: Name | Age | Department | Salary\n----------------------------------------\nJohn Doe | 30 | Engineering | 75000\nJane Smith | 28 | Marketing | 65000\n[Total rows: 150]\n\n=== Sheet2 ===\nHEADERS: Product | Quantity | Price\n----------------------------------\nLaptop | 10 | 1200\nMouse | 50 | 25\n[Total rows: 75]",
    "original_type": "excel",
    "sheet_count": 2,
    "size": 28456
}
```

### Benefits

- 🎯 **Human-readable content** instead of base64 binary data
- 📊 **Rich metadata** for better document understanding
- 🔄 **Backward compatibility** with existing tools
- 🛡️ **Error resilience** with graceful fallbacks
- 📈 **Enhanced user experience** for document analysis

### Tools Compatibility

No changes required to `tools.py` - all existing tools automatically benefit from the enhanced text extraction:

- ✅ `Get_Document_Content` - Now returns readable text
- ✅ `Download_Document` - Unchanged functionality
- ✅ `List_SharePoint_Documents` - Unchanged functionality
- ✅ All other tools - Fully compatible

### Testing

The enhanced functionality can be tested with:

```python
# Test PowerPoint extraction
result = get_document_content("Reports", "presentation.pptx")
print(result["content"])  # Human-readable text instead of binary

# Test Excel extraction
result = get_document_content("Data", "spreadsheet.xlsx")
print(result["sheet_count"])  # Number of sheets

# Test PDF extraction
result = get_document_content("Documents", "report.pdf")
print(result["page_count"])  # Number of pages
```

## 13. Word Document Upload Enhancement

### Overview

Enhanced the `Upload_Document` tool to properly handle Word document uploads by creating actual `.docx` files with proper formatting instead of unreadable text files.

### Problem Solved

Previously, when uploading Word documents through the MCP server, the files were saved as plain text, making them unreadable when opened in Microsoft Word. This enhancement ensures Word documents are properly formatted and human-readable.

### Implementation Details

#### 1. Added Required Dependencies

```python
import io
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
```

#### 2. Created Helper Function

Added `_create_word_document()` function in the Helper functions section of `tools.py`:

- **Markdown Parsing**: Supports `#`, `##`, `###` for headings
- **Text Formatting**: Handles `**bold text**` formatting
- **Document Structure**: Creates proper Word document with:
  - Centered main titles
  - Properly formatted section headings
  - Bold text formatting
  - Regular paragraphs
  - Proper spacing

#### 3. Enhanced Upload_Document Tool

Modified the `Upload_Document` function to:

- **Detect Word Documents**: Automatically identifies `.docx` files
- **Special Processing**: Routes Word documents through `_create_word_document()`
- **Proper MIME Type**: Sets correct content type: `application/vnd.openxmlformats-officedocument.wordprocessingml.document`
- **Backward Compatibility**: Maintains existing functionality for all other file types

### Features

- 📝 **Proper Formatting**: Creates actual Word documents with headings, bold text, and proper structure
- 🎯 **Human-Readable**: Documents open correctly in Microsoft Word
- 🔄 **Markdown Support**: Converts markdown-like syntax to Word formatting
- 🛡️ **Error Handling**: Includes fallback to create simple document if parsing fails
- ✅ **Backward Compatible**: No changes needed for other file types

### Usage Example

```python
# Upload a properly formatted Word document
content = """# NGPI Attribute Use Cases Analysis

## Document Overview
This document outlines use cases identified from analysis.

## Use Case 1: Data Quality Validation
**Added by:** Yi Han
**Description:** Implement comprehensive validation...

---

## Summary
These use cases demonstrate the critical role..."""

# This will now create a proper .docx file
result = await upload_document("", "analysis.docx", content)
```

### Testing

The enhanced Word document upload can be tested with:

```python
# Test Word document creation
result = upload_document("Documents", "test.docx", markdown_content)
print(result["success"])  # Should be True

# Verify the uploaded document is readable
content = get_document_content("Documents", "test.docx")
print(content["content_type"])  # Should be "text" (readable)
```

### Benefits

- 📄 **Professional Documents**: Creates properly formatted Word documents
- 🎯 **User-Friendly**: Documents are immediately readable in Word applications
- 🔧 **Seamless Integration**: Works transparently with existing MCP workflows
- 📊 **Rich Formatting**: Supports headings, bold text, and document structure
- 🛡️ **Robust**: Includes error handling and fallback mechanisms

## 14. Excel Column Context System

### Overview

Implemented a comprehensive column context system that automatically provides definitions and meanings for Excel columns when Windsurf analyzes spreadsheet files from SharePoint.

### Problem Solved

When analyzing Excel files, users often encounter columns with unclear meanings or domain-specific terminology. This system provides automatic context and definitions for any recognized columns, making data analysis more efficient and accurate.

### Implementation

#### 1. Created Column Definitions Dictionary (`column_definitions.md`)
- **Human-readable format**: Easy-to-edit Markdown file
- **Comprehensive coverage**: Includes HR/recruiting, general business, and data quality columns
- **Organized structure**: Grouped by categories (Job Information, Personnel, Dates, etc.)
- **Analytics guidance**: Includes suggested metrics for report generation

#### 2. Built Context Helper System (`context_helper.py`)
- **Markdown parser**: Automatically reads and parses column definitions
- **Smart matching**: Matches Excel headers against the dictionary
- **Auto-enhancement**: Adds column context to Excel content automatically
- **Search capabilities**: Find definitions by keyword or column name

#### 3. Added MCP Tools (`tools.py`)
Four new tools for querying column context:
- `Get_Column_Definition` - Look up specific column meanings
- `Search_Column_Definitions` - Find columns by keyword
- `Get_All_Column_Definitions` - View entire dictionary
- `Get_Matching_Columns` - Check which columns have definitions

#### 4. Enhanced Excel Processing (`resources.py`)
- **Automatic integration**: Excel content now includes column definitions
- **Seamless operation**: Works transparently with existing functionality
- **No user intervention**: Context is added automatically during file reading

### How It Works

**Automatic Enhancement**: When reading Excel files, the system:
1. Extracts column headers from each sheet
2. Matches headers against the definitions dictionary
3. Adds contextual information right after the headers

**Example Output**:
```
=== Sheet1 ===
HEADERS: Job Code | Recruiter | Date Opened | Source

--- COLUMN DEFINITIONS ---
• Job Code: Unique identifier of position by department
• Recruiter: TA Specialist managing the opening
• Date Opened: Date the position is posted
• Source: Where the candidate applied/sourced from (LinkedIn, website, internal referral, etc.)
--- END COLUMN DEFINITIONS ---

[Excel data continues...]
```

### Testing Results

Successfully tested with real SharePoint data:
- **File**: `2023 Recruiting Dataset .xlsx` in Data folder
- **Columns detected**: 18 columns identified
- **Matches found**: 7 columns automatically matched with definitions
- **Context provided**: Immediate understanding of column meanings

### Usage Examples

```python
# Manual column lookup
result = await get_column_definition_tool("Job Code")
# Returns: "Unique identifier of position by department"

# Search for related columns
results = await search_column_definitions_tool("recruiter")
# Finds all columns containing "recruiter" in name or definition

# Check which columns have definitions
matches = await get_matching_columns_tool(["Job Code", "Recruiter", "Unknown Column"])
# Returns definitions for recognized columns
```

### Benefits

- 🎯 **Contextual Analysis**: Excel data comes with built-in explanations
- 📚 **Knowledge Sharing**: Definitions are reusable across all users
- ✏️ **Easy Maintenance**: Simple Markdown format for updates
- 🔍 **Discoverable**: Search and browse available definitions
- 🤖 **AI-Friendly**: Windsurf gets immediate context for better analysis
- 📊 **Report Ready**: Includes metrics suggestions for analytics

### Maintenance

To add new column definitions:
1. Edit `column_definitions.md`
2. Add new entries in format: `- **Column Name**: Description`
3. Restart MCP server to load changes
4. New definitions automatically available to Windsurf

## 15. HR Analytics & Data Quality Tools (Phase 1)

### Overview

Implemented comprehensive analytics capabilities that transform the SharePoint MCP server from a simple file access tool into a powerful HR analytics platform. These tools provide data quality validation, statistical analysis, and chart-ready data generation for HR teams.

### Problem Solved

HR teams often struggle with:
- **Data Quality Issues**: Missing data, outliers, and format inconsistencies in Excel files
- **Manual Analysis**: Time-consuming manual calculation of hiring metrics
- **Visualization Prep**: Complex data preparation for charts and reports
- **Actionable Insights**: Difficulty identifying improvement opportunities from raw data

### Implementation

#### 1. Analytics Helper Module (`analytics_helper.py`)
- **Pandas Integration**: Robust DataFrame processing for Excel data
- **Data Cleaning**: Automatic date/numeric conversion and text standardization
- **Statistical Analysis**: Comprehensive HR metrics calculation
- **Chart Data Generation**: Ready-to-use data for 5 different visualization types

#### 2. New MCP Tools (4 tools added to `tools.py`)

**`Validate_Excel_Data_Quality`**
- Identifies missing data percentages by column
- Detects statistical outliers using IQR method
- Finds format issues (invalid date sequences, etc.)
- Provides comprehensive data quality summary

**`Calculate_HR_Metrics`**
- Hiring volume metrics (total, filled, open positions)
- Time-to-hire statistics (average, median, percentiles)
- Application metrics (total applications, averages)
- Source and location distribution analysis
- Cost analysis (agency fees, percentages)

**`Generate_Chart_Data`**
- **hiring_trends**: Monthly hiring patterns over time
- **source_effectiveness**: Candidate source distribution (pie chart)
- **time_to_hire_distribution**: Histogram of hiring timeframes
- **department_hiring**: Hiring volume by department (bar chart)
- **conversion_funnel**: Application-to-hire conversion rates

**`Analyze_HR_File_Complete`**
- Combines all analyses into one comprehensive report
- Generates multiple chart suggestions automatically
- Provides actionable recommendations based on data patterns
- Includes data summary with date ranges and column information

### Real-World Testing Results

Successfully tested with actual HR data (`2023 Recruiting Dataset .xlsx`):

#### Data Quality Insights
- **106 positions analyzed** across 18 columns
- **Missing data identified** in 17/18 columns (Agency Fee 91.5% missing)
- **Outliers detected**: 7 positions with 160+ day hiring cycles
- **Format issues found**: 2 positions with invalid date sequences

#### Key HR Metrics Calculated
- **Hiring Performance**: 71 filled positions, 35 open, 74.5 day average time-to-hire
- **Application Volume**: 39,694 total applications, 684 average per position
- **Source Analysis**: LinkedIn 63% of hires, Internal Promotion 15%
- **Cost Impact**: $88,968 in agency fees across 8.5% of positions
- **Geographic Distribution**: 5 countries (US, France, Canada, Australia, Mexico)

#### Automated Recommendations Generated
- **Data Quality**: "Review missing data in 17 columns for better analysis"
- **Source Diversification**: "Reduce over-reliance on LinkedIn (63% of hires)"
- **Process Optimization**: Identified positions taking 200+ days to fill

### Usage Examples

```python
# Complete analysis of HR file
result = await analyze_hr_file_complete_tool("Data", "2023 Recruiting Dataset.xlsx")
# Returns: data quality, metrics, chart suggestions, recommendations

# Data quality validation only
quality = await validate_excel_data_quality_tool("Data", "recruiting_data.xlsx")
# Returns: missing data %, outliers, format issues

# Generate specific chart data
chart = await generate_chart_data_tool("Data", "hiring_data.xlsx", "source_effectiveness")
# Returns: labels and values ready for pie chart

# Calculate HR metrics
metrics = await calculate_hr_metrics_tool("Data", "hr_file.xlsx")
# Returns: time-to-hire, costs, source distribution, etc.
```

### Chart Data Output Format

All chart tools return data in a standardized format:
```json
{
  "chart_type": "source_effectiveness",
  "data": {
    "labels": ["LinkedIn", "Internal Promotion", "Indeed"],
    "values": [43, 10, 6],
    "title": "Candidate Source Distribution"
  }
}
```

### Benefits for HR Teams

- 🔍 **Instant Data Quality Assessment**: Identify issues before analysis
- 📊 **Comprehensive Metrics**: 15+ key HR metrics calculated automatically
- 📈 **Visualization Ready**: Chart data formatted for immediate use
- 💡 **Actionable Insights**: Automated recommendations for process improvement
- ⏱️ **Time Savings**: Minutes instead of hours for analysis
- 🎯 **Consistent Analysis**: Standardized metrics across all files

### Supported Chart Types

1. **Hiring Trends**: Monthly hiring volume over time (line chart)
2. **Source Effectiveness**: Candidate source distribution (pie chart)
3. **Time-to-Hire Distribution**: Hiring duration buckets (histogram)
4. **Department Hiring**: Hiring volume by department (bar chart)
5. **Conversion Funnel**: Application-to-hire process flow (funnel chart)

### Data Quality Checks

- **Missing Data**: Percentage and count by column
- **Statistical Outliers**: IQR-based anomaly detection
- **Format Validation**: Date sequence consistency
- **Data Type Verification**: Automatic numeric/date conversion
- **Completeness Assessment**: Overall data quality scoring

### Recommendations Engine

The system automatically generates recommendations in 4 categories:
- **Data Quality**: Missing data and cleanup suggestions
- **Performance**: Time-to-hire and process efficiency
- **Cost Optimization**: Agency fee and sourcing cost analysis
- **Source Diversification**: Candidate pipeline risk assessment

### Technical Architecture

- **Pandas Backend**: Robust data processing and analysis
- **Statistical Methods**: IQR outlier detection, percentile calculations
- **Error Handling**: Graceful degradation for incomplete data
- **Memory Efficient**: Streaming processing for large datasets
- **Extensible Design**: Easy to add new metrics and chart types

## 16. Excel Visualization Implementation & Cleanup

### Overview

Successfully implemented and refined Excel chart generation capabilities, focusing on reliable Excel-based visualizations while removing problematic HTML dashboard components. This provides HR teams with professional Excel reports containing embedded charts.

### Phase 1 Implementation (Completed)

#### **Excel Chart Generation**
- ✅ **Embedded Charts**: Pie charts, bar charts, and line charts directly in Excel workbooks
- ✅ **Multiple Chart Types**: Source effectiveness (pie), hiring trends (line), time-to-hire distribution (bar)
- ✅ **Data Integration**: Raw data and charts in the same workbook for easy reference
- ✅ **SharePoint Upload**: Automatic upload of generated Excel files to SharePoint

#### **Chart Types Supported**
1. **Source Effectiveness**: Pie chart showing candidate source distribution
2. **Hiring Trends**: Line chart displaying monthly hiring patterns over time
3. **Time-to-Hire Distribution**: Bar chart showing hiring duration buckets
4. **Department Hiring**: Bar chart of hiring volume by department (when data available)

#### **Technical Implementation**
- **OpenPyXL Integration**: Professional Excel chart creation using openpyxl library
- **Chart Isolation**: Each chart instance is completely independent to prevent reuse errors
- **Data Sheets**: Separate sheets for raw data and chart analysis
- **Error Handling**: Graceful degradation when chart data is unavailable

#### **Chart Quality Features**
Based on user feedback, implemented professional chart formatting:
- **Clear Column Headers**: Bold, context-specific headers in data tables (e.g., "Source" | "Number of Hires")
- **Y-Axis Visibility**: Numerical scale displayed on all bar/line charts
- **Clean Background**: Removed gridlines for professional appearance
- **Time Series X-Axis**: Dates/months displayed on x-axis for trend charts
- **Clean Titles**: Professional chart titles without debug numbers
- **Self-Documenting**: Charts and data tables provide complete context

### Testing Results

Successfully tested with real HR data (`2023 Recruiting Dataset .xlsx`):
- **✅ Excel Generation**: Created `charts_2023 Recruiting Dataset_analysis.xlsx`
- **✅ Chart Embedding**: 3 charts successfully embedded (pie, line, bar)
- **✅ SharePoint Upload**: File uploaded and accessible via SharePoint
- **✅ Data Integrity**: All 106 positions and 18 columns preserved
- **✅ Professional Quality**: Charts ready for executive presentations

### Cleanup & Optimization (Phase 1 Refinement)

#### **Removed Components**
Based on user feedback and technical challenges, removed non-essential components:

1. **HTML Dashboard Generation**
   - Removed `Create_Interactive_Dashboard` MCP tool
   - Removed Plotly chart creation functionality
   - Removed HTML template system and CSS styling

2. **Complete Package Tool**
   - Removed `Generate_Visual_Report_Package` MCP tool
   - Eliminated dual-format generation complexity

3. **Dashboard Helper Methods**
   - Removed `create_dashboard()`, `create_plotly_chart()`, `_create_metrics_cards()`
   - Removed `_create_data_table()`, `_create_javascript_code()`, `save_dashboard_file()`

4. **Unused Dependencies**
   - Removed Plotly dependency from `pyproject.toml`
   - Cleaned up imports in `visualization_helper.py`

#### **Streamlined Architecture**
- **Focused Functionality**: Excel chart generation only
- **Reduced Complexity**: Single visualization format (Excel)
- **Improved Reliability**: Eliminated problematic HTML/JavaScript components
- **Cleaner Codebase**: Removed 200+ lines of unused code

### Current MCP Tools (Post-Cleanup)

#### **Analytics Tools**
1. **`Validate_Excel_Data_Quality`**: Data quality validation and issue identification
2. **`Calculate_HR_Metrics`**: Comprehensive HR metrics calculation
3. **`Generate_Chart_Data`**: Chart data generation for external tools
4. **`Analyze_HR_File_Complete`**: Complete analysis with recommendations

#### **Visualization Tools**
1. **`Create_Excel_With_Charts`**: Excel file generation with embedded charts

### Usage Examples

```python
# Generate Excel file with embedded charts
result = await create_excel_with_charts_tool("Data", "recruiting_data.xlsx")
# Returns: Excel file with pie, bar, and line charts + raw data

# Get comprehensive analysis
analysis = await analyze_hr_file_complete_tool("Data", "recruiting_data.xlsx") 
# Returns: metrics, data quality, recommendations

# Generate chart data for external tools
chart_data = await generate_chart_data_tool("Data", "recruiting_data.xlsx", "source_effectiveness")
# Returns: labels and values ready for any visualization tool
```

### File Output Structure

**Generated Excel File Structure:**
```
charts_[filename]_analysis.xlsx
├── Raw Data (Sheet 1)
│   ├── All original data preserved
│   └── 18 columns × 106 rows (example)
└── Charts & Analysis (Sheet 2)
    ├── Source Effectiveness (Pie Chart)
    ├── Hiring Trends (Line Chart)
    ├── Time-to-Hire Distribution (Bar Chart)
    └── Underlying chart data tables
```

### Benefits Achieved

- 🎯 **Focused Solution**: Excel-only approach eliminates complexity
- 📊 **Professional Output**: Charts ready for executive presentations
- 🔧 **Reliable Generation**: No HTML/JavaScript compatibility issues
- 📈 **Immediate Usability**: Works with existing Excel workflows
- 🚀 **Fast Performance**: Streamlined code executes quickly
- 🛠️ **Easy Maintenance**: Single visualization pathway to maintain

### Dependencies (Final)

**Core Requirements:**
- `pandas>=2.0.0` - Data processing and analysis
- `numpy>=1.24.0` - Statistical calculations
- `openpyxl>=3.1.0` - Excel file creation and chart embedding

**Removed Dependencies:**
- ~~`plotly>=5.17.0`~~ - No longer needed after HTML dashboard removal

## 17. Phase 2: PowerPoint Report Generation with AI-Powered Insights

### Overview

Successfully implemented professional PowerPoint presentation generation with intelligent, data-driven insights. This enhancement provides executive-ready presentations that automatically discover patterns and provide actionable recommendations.

### Implementation Details

#### **PowerPoint Helper Module** (`powerpoint_helper.py`)

Created comprehensive PowerPoint generation system with professional formatting:

**Key Features:**
- **Title Slide**: Centered title with professional blue color scheme, auto-generated date
- **Content Slides**: Colored header bars with white text, consistent formatting
- **Chart Integration**: Supports pie, line, and bar charts with proper legends
- **Layout Management**: Charts on left (5.5" wide), insights on right - no overlap
- **Table Support**: Professional tables with colored headers for data definitions
- **Color Scheme**: Professional blues (primary, secondary) and orange accents

**Technical Implementation:**
```python
class PowerPointHelper:
    - create_title_slide() - Professional title with centered layout
    - create_content_slide() - Slides with colored header bars
    - create_chart_slide() - Charts + insights with proper spacing
    - create_data_definitions_slide() - Table format with definitions
    - add_chart_to_slide() - Chart embedding with legend support
    - add_bullet_points() - Formatted bullet lists for insights
    - add_table() - Professional tables with header formatting
```

#### **MCP Tool: Create_PowerPoint_Report**

**Parameters:**
- `file_name` (required): Excel file to analyze
- `folder_name` (optional): Source folder, defaults to "Data"
- `output_folder` (optional): Destination folder, defaults to "AI Generated Reports"
- `presentation_title` (optional): Custom title, defaults to "HR Recruiting Analytics Report"

**Workflow:**
1. Reads Excel file from source folder
2. Performs data quality validation and HR metrics calculation
3. Generates 3 chart slides with AI-powered insights
4. Adds data definitions slide with column explanations
5. Uploads PowerPoint to SharePoint output folder
6. Returns success message with slide count

#### **AI-Powered Insight Generation**

Implemented `_generate_chart_insights()` function that performs deep data analysis:

**Source Effectiveness Insights:**
- Concentration risk analysis (tiered: >60% critical, >40% moderate)
- Top 3 sources cumulative impact calculation
- Underutilized sources identification (<5% contributors)
- Source diversification quality assessment
- Strategic recommendations based on patterns

**Hiring Trends Insights:**
- Volatility detection (peak vs average ratio analysis)
- Growth trend analysis (recent vs earlier periods with % change)
- Seasonality pattern recognition (quarterly comparisons)
- Gap analysis (identifies zero hiring months)
- Trend classification (growth/decline/stable)

**Time-to-Hire Insights:**
- Multi-tier categorization (Fast <30, Optimal 31-60, Acceptable 61-90, Slow 90+)
- Performance benchmarking against best practices
- Fast-fill pattern analysis and interpretation
- Critical bottleneck alerts (>30% slow fills)
- Pipeline health scoring (combined fast + optimal)
- Process efficiency recommendations

### Testing Results

Successfully tested with 2023 recruiting dataset (106 positions, 18 columns):

**Generated Output:**
- **Slide 1**: Professional title slide with centered layout
- **Slide 2**: Candidate Source Distribution (pie chart) + 6-8 insights
- **Slide 3**: Monthly Hiring Trends (line chart) + 5-7 insights
- **Slide 4**: Time-to-Hire Distribution (bar chart) + 7-9 insights
- **Slide 5**: Data Column Definitions (table with first 10 columns)

**File Details:**
- File size: ~62 KB
- Format: .pptx (PowerPoint)
- Upload location: AI Generated Reports folder
- Download: Automatic to local filesystem

### Key Insights Discovered by AI

**Example Patterns Automatically Detected:**

1. **Concentration Risk**: "LinkedIn is the dominant source at 63.2%" + "⚠️ High concentration risk - over 60% from single source"

2. **Volatility**: "⚠️ High volatility detected - peak is 2.8x average" (discovered December 2022 spike)

3. **Growth Trends**: "📈 Strong growth: 45% increase in recent months" (comparing Q1 vs Q4)

4. **Seasonality**: "🔄 Seasonal pattern detected - consider planning for peaks" (quarterly analysis)

5. **Process Efficiency**: "🎯 Focus area: More slow fills than optimal - streamline interview process"

6. **Pipeline Health**: "📊 53.8% filled within 60 days - healthy pipeline"

### Professional Formatting Features

✅ **Headers**: Colored header bar (professional blue) on every content slide  
✅ **No Overlap**: Charts and text properly spaced (charts left, insights right)  
✅ **Bullet Points**: Multiple insights formatted as clean bullet lists  
✅ **Color Scheme**: Consistent professional blue and orange accents  
✅ **Legends**: All charts include legends for clarity  
✅ **Centered Title**: Professional title slide with centered layout  
✅ **Table Headers**: Bold, colored headers in definition tables  
✅ **Emoji Icons**: Visual indicators for severity (⚠️ warning, 🔴 critical, ✅ good, 📈 trend)

### Benefits Achieved

- 🧠 **Intelligent Analysis**: AI discovers patterns humans might miss
- 📊 **Executive-Ready**: Professional formatting suitable for C-level presentations
- ⚡ **Automated**: Generates complete presentation in ~15 seconds
- 🎯 **Actionable**: Provides specific recommendations, not just statistics
- 📈 **Data-Driven**: All insights based on actual data patterns
- 🔄 **Repeatable**: Can regenerate anytime with fresh data
- 💼 **Business Value**: Replaces hours of manual slide creation

### Current MCP Tools (Phase 1 + Phase 2)

#### **Analytics Tools**
1. **`Validate_Excel_Data_Quality`**: Data quality validation and issue identification
2. **`Calculate_HR_Metrics`**: Comprehensive HR metrics calculation
3. **`Generate_Chart_Data`**: Chart data generation for external tools
4. **`Analyze_HR_File_Complete`**: Complete analysis with recommendations

#### **Visualization & Export Tools**
5. **`Create_Excel_With_Charts`**: Excel file generation with embedded charts (Phase 1)
6. **`Create_PowerPoint_Report`**: PowerPoint presentation with AI insights (Phase 2) ⭐ **NEW**

### Usage Example

```python
# Generate PowerPoint report with AI insights
result = await create_powerpoint_report_tool(
    file_name="2023 Recruiting Dataset.xlsx",
    folder_name="Data",  # Optional, defaults to "Data"
    output_folder="AI Generated Reports",  # Optional
    presentation_title="2023 Recruiting Performance Analysis"  # Optional
)

# Returns:
{
    "success": True,
    "message": "PowerPoint presentation created: report_2023 Recruiting Dataset.pptx",
    "file_name": "report_2023 Recruiting Dataset.pptx",
    "slides_created": 5,
    "download_info": "File uploaded to SharePoint and ready for download"
}
```

### Dependencies (Updated)

**Core Requirements:**
- `pandas>=2.0.0` - Data processing and analysis
- `numpy>=1.24.0` - Statistical calculations
- `openpyxl>=3.1.0` - Excel file creation and chart embedding
- `python-pptx>=0.6.21` - PowerPoint file creation and formatting ⭐ **ADDED**

**Removed Dependencies:**
- ~~`plotly>=5.17.0`~~ - No longer needed after HTML dashboard removal

### Future Enhancements (Optional)

If additional capabilities are needed:
1. **Chart Customization**: Color schemes, custom fonts, data point markers
2. **Additional Chart Types**: Scatter plots, combo charts, pivot charts
3. **Multi-Sheet Reports**: Separate sheets for different analysis types
4. **Template System**: Predefined templates for different report types
5. **Custom Branding**: Company logos, custom color schemes
6. **Slide Transitions**: Animated transitions and effects
7. **Speaker Notes**: Auto-generated presenter notes with talking points

The MCP server is now **fully compatible with Windsurf** and ready for production use with both Excel and PowerPoint report generation! 🎯✅

## 18. Context Management System Integration

### Overview

The SharePoint MCP server includes a comprehensive context management system that provides rich business context to all tools, enabling more intelligent and context-aware responses.

### Context Files and Categories

The system organizes context into 4 categories across 6 markdown files:

#### **1. Columns Context** (`columns`)
- **Files:** `column_definitions.md`
- **Content:** Excel column definitions for HR/recruiting data
- **Parsed:** 73+ column definitions automatically extracted
- **Used By:** All Excel analysis tools

#### **2. Metrics Context** (`metrics`)
- **Files:** `metrics_definitions.md`
- **Content:** HR metrics, KPIs, and analytics guidance
- **Used By:** Analytics and reporting tools

#### **3. Business Context** (`business`)
- **Files:** `company_overview.md`, `engineering_overview.md`
- **Content:** Company mission, values, strategy, culture
- **Used By:** Report generation, analysis tools

#### **4. Recruiting Context** (`recruiting`)
- **Files:** `hiring_guide.md`, `career_path.md`
- **Content:** Hiring processes, career frameworks
- **Used By:** HR metrics, validation tools

#### **5. Roles Context** (`roles`)
- **Files:** 
  - `software_engineer_role_description.md` - Software engineering career ladder (L1-L6)
  - `data_management_role_description.md` - Data engineering career paths
  - `engineering_leadership_role_description.md` - Engineering management track
  - `ml_ds_da_role_description.md` - ML/DS/DA role descriptions (references PDFs)
- **Content:** Role descriptions, career progression, expectations by level
- **Used By:** HR analytics, recruiting analysis, workforce planning tools
- **Note:** ML/DS/DA roles reference PDF files in the same context folder for detailed descriptions

### Tool-Context Mapping

Each tool automatically receives relevant context:

| Tool | Context Categories | What It Gets |
|------|-------------------|--------------|
| **Analyze_HR_File_Complete** | columns, business, metrics, roles | Column defs + company values + KPIs + role descriptions |
| **Calculate_HR_Metrics** | columns, metrics, recruiting, roles | Column defs + metrics + hiring guidelines + role info |
| **Validate_Excel_Data_Quality** | columns, recruiting, roles | Column defs + hiring best practices + role expectations |
| **Create_PowerPoint_Report** | columns, business, metrics, recruiting, roles | All context for comprehensive reports |
| **Create_Excel_With_Charts** | columns, metrics | Column defs + chart guidance |
| **Generate_Chart_Data** | columns, metrics | Column defs + visualization guidance |

### New MCP Tools for Context Access

#### **1. Get_Context_Summary**
Get summary of all loaded context files and categories

**Returns:**
```json
{
  "success": true,
  "summary": {
    "total_files": 10,
    "total_characters": 92391,
    "categories": ["columns", "metrics", "business", "recruiting", "roles"],
    "column_definitions_count": 73
  }
}
```

#### **2. Search_All_Context**
Search across all context files for specific information

**Parameters:**
- `search_term` (string) - The term to search for
- `categories` (list, optional) - Categories to search in

**Example:**
```python
# Search for information about diversity
results = context_manager.search_context('diversity', categories=['business', 'recruiting'])
```

### How Context Works

#### **Automatic Context Loading**
```python
# On server startup
context_manager = ContextManager()
# Automatically loads all .md files from context/ directory
# Categorizes them based on CONTEXT_CATEGORIES mapping
# Parses column definitions for quick lookup
```

#### **Tool Context Injection**
```python
# When a tool is called
context = context_manager.get_context_for_tool('Analyze_HR_File_Complete')
# Returns combined context from: columns, business, metrics categories
# AI receives this context to inform its responses
```

#### **Column Definition Matching**
```python
# When analyzing Excel files
matching_defs = context_manager.get_matching_columns(df.columns.tolist())
# Automatically matches Excel columns to definitions
# Enhances analysis with column context
```

### Context File Standards

All context files follow these standards:

**Format:**
- ✅ Markdown (.md) format
- ✅ Clear headings (`#`, `##`, `###`)
- ✅ Bullet points for lists
- ✅ Bold for emphasis (`**text**`)
- ✅ Tables for structured data

**Content:**
- ✅ No sensitive financial data
- ✅ No personal names
- ✅ No external URLs
- ✅ No internal document references
- ✅ No vendor/partner names
- ✅ No specific interview questions

### Testing Context Integration

Run the test script to verify context loading:

```bash
source venv/bin/activate
python test_context_integration.py
```

**Expected output:**
```
✓ Total files loaded: 10
✓ Total characters: 92,391
✓ Categories: columns, metrics, business, recruiting, roles
✓ Column definitions: 73
✓ Context for tools: Working
✓ Column matching: Working
```

### Usage Example: Context-Aware Analysis

**Before Context:**
> "Average time-to-hire is 75 days"

**With Context:**
> "⚠️ Average time-to-hire is 75 days, which exceeds the company target of 60 days by 25%. Based on company values emphasizing efficiency and the recruiting guidelines, recommend reviewing interview scheduling and hiring manager availability."

### Benefits

**For Users:**
- 🎯 More intelligent, context-aware responses
- 📊 Better recommendations aligned with company goals
- 🔍 Deeper insights from data analysis
- 📝 More comprehensive reports

**For Developers:**
- 🛠️ Easy to add new context files
- 🔧 Simple tool-context mapping
- 📦 Modular and maintainable
- 🧪 Easy to test and extend

**For the Business:**
- 💼 AI understands company values and strategy
- 📈 Recommendations aligned with business goals
- 🎓 Knowledge embedded in the system
- 🔒 Sensitive data properly protected

### Adding New Context Files

To add new context files:

1. **Create markdown file** in `src/mcp_sharepoint/context/`
2. **Add to category mapping** in `context_manager.py`:
   ```python
   CONTEXT_CATEGORIES = {
       'columns': ['column_definitions.md'],
       'metrics': ['metrics_definitions.md'],
       'business': ['company_overview.md', 'engineering_overview.md', 'your_new_file.md'],
       'recruiting': ['hiring_guide.md', 'career_path.md']
   }
   ```
3. **Update tool mapping** if needed:
   ```python
   TOOL_CONTEXT_MAP = {
       'Your_Tool_Name': ['columns', 'business', 'your_new_category'],
   }
   ```
4. **Restart server** - Context loads automatically on startup

### Context Files Location

All context files are stored in:
```
src/mcp_sharepoint/context/
├── column_definitions.md                      # Excel column definitions
├── metrics_definitions.md                     # HR metrics and KPIs
├── company_overview.md                        # Company mission and values
├── engineering_overview.md                    # Engineering culture
├── hiring_guide.md                           # Hiring processes
├── career_path.md                            # Career frameworks
├── software_engineer_role_description.md      # Software engineering career ladder
├── data_management_role_description.md        # Data engineering career paths
├── engineering_leadership_role_description.md # Engineering management track
├── ml_ds_da_role_description.md              # ML/DS/DA role descriptions
└── Position-Description-*.pdf                 # Detailed role descriptions (12 PDFs)
```

### Migration Notes

The old `context_helper.py` has been removed. All tools now use `context_manager.py`:

**What Changed:**
- ✅ `context_helper.get_column_definition()` → `context_manager.get_column_definition()`
- ✅ `context_helper.get_all_definitions()` → `context_manager.get_all_column_definitions()`
- ✅ `context_helper.get_matching_columns()` → `context_manager.get_matching_columns()`
- ✅ `context_helper.search_definitions()` → `context_manager.search_context()`

**What's New:**
- ✅ `context_manager.get_context_for_tool()` - Get context for specific tool
- ✅ `context_manager.get_context_by_category()` - Get all context in a category
- ✅ `context_manager.get_context_summary()` - Get statistics
- ✅ `context_manager.search_context()` - Search across all context

### Status

**Context System:** ✅ **Production Ready**  
**Files Loaded:** 10 markdown files + 12 PDF files  
**Categories:** 5 (columns, metrics, business, recruiting, roles)  
**Column Definitions:** 73 parsed definitions  
**Integration:** Complete

## 19. Training Data Generation (One-Time Setup)

### Overview

For testing and demonstration purposes, a training data generator script was created to populate SharePoint with realistic HR recruiting data. This is a **one-time setup** that creates sample Excel files for testing the MCP server's analytics capabilities.

### Quick Start

Generate training data with a single command:

```bash
# Activate environment and run
source venv/bin/activate && python generate_training_data.py
```

### What Gets Generated

The script automatically creates and uploads 4 Excel files to SharePoint:

| File | Records | Columns | Purpose |
|------|---------|---------|---------|
| **Filled Positions** | 75-100 | 17 | Candidate hiring data |
| **Time In Step Q3** | 75-100 | 27 | Pipeline time tracking |
| **Recruiting Report** | 75-100 | 19 | Funnel metrics |
| **All Departments** | 75-100 | 24 | Department analytics |

### Features

- ✅ **Realistic Data**: Names, departments, dates, metrics all randomly generated
- ✅ **Automatic Upload**: Files uploaded directly to SharePoint Data folder
- ✅ **Record Counts**: Filenames include record counts (e.g., "Filled Positions (90 records).xlsx")
- ✅ **Proper Formatting**: Excel files with headers, data types, and formatting
- ✅ **No Duplicates**: Each run generates fresh data

### When to Use

**Use training data generation when:**
- 🧪 Testing the MCP server for the first time
- 📊 Demonstrating analytics capabilities
- 🎓 Training users on the system
- 🔍 Validating new features

**Skip if:**
- ✅ You already have real HR data in SharePoint
- ✅ You've previously run the generator and have test data

### Expected Output

```
================================================================================
TRAINING DATA GENERATOR FOR HR ANALYTICS
================================================================================

🔄 Generating training datasets...

📊 Generating File 1: Filled Positions (90 records)...
   ✅ Generated 90 records

📊 Generating File 2: Time In Step Q3 (97 records)...
   ✅ Generated 97 records

📊 Generating File 3: Recruiting Report (86 records)...
   ✅ Generated 86 records

📊 Generating File 4: All Departments Report (91 records)...
   ✅ Generated 91 records

================================================================================
✅ ALL FILES GENERATED SUCCESSFULLY!
Total files: 4
Total records: 364
Upload location: SharePoint Data folder
================================================================================
```

### Detailed Documentation

For complete details about the training data generator, see:
- **`QUICK_START_TRAINING_DATA.md`** - Quick reference guide
- **`TRAINING_DATA_GENERATOR_README.md`** - Comprehensive documentation including:
  - Column definitions for each file
  - Data generation logic
  - Customization options
  - Troubleshooting

### Notes

- **One-time setup**: Only needs to be run once to create test data
- **Safe to re-run**: Will create new files with different record counts
- **Requires SharePoint access**: Uses same credentials as MCP server
- **Data folder**: Files are uploaded to the "Data" folder in SharePoint
- **Random data**: Each run generates different realistic data

### Status

**Training Data Generator:** ✅ **Available**  
**Purpose:** Testing and demonstration  
**Usage:** One-time setup (optional)
