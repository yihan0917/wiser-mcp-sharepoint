# How AI Agents Use the File Naming Convention

## Overview

This document explains how AI agents automatically discover and use the file naming convention to locate files in SharePoint without requiring explicit file names from users.

## The Problem We Solved

**Before:** When a user asked "analyze Q3 training data", the AI agent would:
- Not know which file to use
- Ask the user for the exact file name
- Or fail to complete the task

**After:** The AI agent can:
- Read the file naming convention from MCP resources
- Understand the naming pattern
- List files in SharePoint
- Match files based on the user's request
- Automatically select the correct file

## How It Works

### 1. **File Naming Convention Context**

The convention is defined in:
```
src/mcp_sharepoint/context/file_naming_convention.md
```

Convention structure:
```
{content}_{source}_{time_period}_{year}.xlsx
```

Example files:
- `Engineering_Hiring_Pipeline_SmartRecruiter_Q3_2024.xlsx`
- `Leadership_Development_Training_Manual_Q1_2025.xlsx`
- `New_Hire_Onboarding_SmartRecruiter_WK25_2024.xlsx`

### 2. **MCP Resources Expose Context**

The context is exposed as **MCP Resources** that AI agents can read:

```python
@mcp.resource("context://file-naming-convention")
def get_file_naming_convention() -> str:
    """
    File naming convention for Excel files in SharePoint.
    Read this BEFORE attempting to analyze data when the user
    provides a generic request.
    """
    return context_manager.get_context_by_category('file_naming')
```

Available resources:
- `context://file-naming-convention` - File naming convention
- `context://column-definitions` - Excel column definitions
- `context://metrics-definitions` - HR metrics definitions
- `context://all-context` - Complete context (all categories)
- `context://summary` - Overview of available context

### 3. **AI Agent Workflow**

When a user makes a generic request like **"analyze Q3 training data"**:

#### Step 1: AI reads the file naming convention
```
AI → Read resource: context://file-naming-convention
```

The AI learns:
- Convention structure: `{content}_{source}_{time_period}_{year}.xlsx`
- How to construct search patterns
- Examples of valid file names

#### Step 2: AI lists files in SharePoint
```
AI → Call tool: List_SharePoint_Documents(folder_name="Data")
```

Returns files like:
- `Engineering_Hiring_Pipeline_SmartRecruiter_Q3_2024.xlsx`
- `Leadership_Development_Training_Manual_Q1_2025.xlsx`
- `Q3_Performance_Reviews_Manual_June_2024.xlsx`

#### Step 3: AI matches files based on user request
The AI parses "analyze Q3 training data":
- Content: "training" → matches `*Training*`
- Time period: "Q3" → matches `*Q3*`
- Pattern: `*Training*Q3*.xlsx`

Finds match: `Leadership_Development_Training_Manual_Q1_2025.xlsx`

#### Step 4: AI calls the analysis tool
```
AI → Call tool: Analyze_HR_File_Complete(
    folder_name="Data",
    file_name="Leadership_Development_Training_Manual_Q1_2025.xlsx"
)
```

## Example User Interactions

### Example 1: Quarterly Training Data
**User:** "Show me Q3 training metrics"

**AI Process:**
1. Reads `context://file-naming-convention`
2. Lists files in SharePoint
3. Searches for `*Training*Q3*.xlsx`
4. Finds: `Leadership_Development_Training_Manual_Q3_2024.xlsx`
5. Analyzes the file

### Example 2: Monthly Performance Reviews
**User:** "Analyze June performance reviews"

**AI Process:**
1. Reads `context://file-naming-convention`
2. Lists files in SharePoint
3. Searches for `*Performance*Review*June*.xlsx`
4. Finds: `Q3_Performance_Reviews_Manual_June_2024.xlsx`
5. Analyzes the file

### Example 3: Weekly Onboarding Data
**User:** "Check week 25 onboarding data"

**AI Process:**
1. Reads `context://file-naming-convention`
2. Lists files in SharePoint
3. Searches for `*Onboarding*WK25*.xlsx`
4. Finds: `New_Hire_Onboarding_SmartRecruiter_WK25_2024.xlsx`
5. Analyzes the file

### Example 4: Ambiguous Request
**User:** "Analyze hiring data"

**AI Process:**
1. Reads `context://file-naming-convention`
2. Lists files in SharePoint
3. Searches for `*Hiring*.xlsx`
4. Finds multiple matches:
   - `Engineering_Hiring_Pipeline_SmartRecruiter_Q3_2024.xlsx`
   - `Software_Engineer_Hiring_Manual_WK48_2025.xlsx`
5. **Asks user to clarify:** "I found 2 hiring files. Which one would you like to analyze?"

## Context Manager vs MCP Resources

### Understanding the Difference

The SharePoint MCP server uses **two complementary systems** to provide context:

#### **Context Manager** (Internal Tool Context)
- **Purpose**: Provides context to tool implementation code (Python functions)
- **Visibility**: Internal only - AI agents cannot see this
- **When Used**: During tool execution (after the tool is called)
- **Access Method**: `context_manager.get_context_for_tool('ToolName')`

**Example:**
```python
# Inside tools.py
def analyze_hr_file_complete(folder_name: str, file_name: str):
    # Tool gets context internally
    tool_context = context_manager.get_context_for_tool('Analyze_HR_File_Complete')
    # Uses column definitions, business context, role descriptions
    # AI agent never sees this - it's used by the tool's Python code
```

#### **MCP Resources** (External AI Agent Context)
- **Purpose**: Provides context to AI agents before they call tools
- **Visibility**: External - AI agents can discover and read these
- **When Used**: Before tool calls (helps AI make decisions)
- **Access Method**: AI reads via MCP protocol (e.g., `context://file-naming-convention`)

**Example:**
```python
# In resources.py
@mcp.resource("context://file-naming-convention")
def get_file_naming_convention() -> str:
    # Exposes context to AI agents
    return context_manager.get_context_by_category('file_naming')
```

### How They Work Together

```
User: "Analyze Q3 training data"
    ↓
┌─────────────────────────────────────────┐
│ AI Agent (External)                     │
│ - Reads MCP Resource                    │
│   context://file-naming-convention      │
│ - Learns file naming pattern            │
│ - Lists files in SharePoint             │
│ - Finds matching file                   │
│ - Decides to call tool                  │
└─────────────────────────────────────────┘
    ↓
    Calls: Analyze_HR_File_Complete(...)
    ↓
┌─────────────────────────────────────────┐
│ Tool Implementation (Internal)          │
│ - Uses Context Manager                  │
│ - Gets column definitions               │
│ - Gets business context                 │
│ - Gets role descriptions                │
│ - Generates intelligent analysis        │
└─────────────────────────────────────────┘
    ↓
    Returns results to AI Agent
```

### Why Both Are Needed

| Problem | Solution |
|---------|----------|
| AI doesn't know file naming convention | **MCP Resources** expose it |
| AI doesn't know which file to use | **MCP Resources** provide guidance |
| Tool needs to understand column meanings | **Context Manager** provides definitions |
| Tool needs business context for insights | **Context Manager** provides company info |

### Key Insight

**MCP Resources actually use the Context Manager:**
```python
@mcp.resource("context://file-naming-convention")
def get_file_naming_convention() -> str:
    # MCP Resource calls Context Manager
    return context_manager.get_context_by_category('file_naming')
```

The Context Manager is the **source of truth** (loads all context files), and MCP Resources are the **external interface** that exposes selected context to AI agents.

**Analogy:**
- **Context Manager** = Your company's internal knowledge base
- **MCP Resources** = Public documentation that customers can read

Both use the same underlying knowledge, but serve different audiences!

## Technical Implementation

### Context Manager
`src/mcp_sharepoint/context_manager.py`

Loads all context files at startup:
```python
CONTEXT_CATEGORIES = {
    'file_naming': ['file_naming_convention.md'],
    'columns': ['column_definitions.md'],
    'metrics': ['metrics_definitions.md'],
    # ... other categories
}
```

### MCP Resources
`src/mcp_sharepoint/resources.py`

Exposes context as MCP resources:
```python
@mcp.resource("context://file-naming-convention")
def get_file_naming_convention() -> str:
    return context_manager.get_context_by_category('file_naming')
```

### Tool Context Mapping
Tools that need file naming context:
```python
TOOL_CONTEXT_MAP = {
    'Analyze_HR_File_Complete': ['file_naming', 'columns', 'metrics', ...],
    'List_SharePoint_Documents': ['file_naming'],
    'Get_Document_Content': ['file_naming', 'columns'],
    # ... other tools
}
```

## Benefits

### For Users
- ✅ No need to remember exact file names
- ✅ Natural language requests work
- ✅ Faster workflow
- ✅ Less friction

### For AI Agents
- ✅ Clear guidance on file naming patterns
- ✅ Ability to construct search patterns
- ✅ Handle ambiguity gracefully
- ✅ Provide better user experience

### For Developers
- ✅ Standardized file naming across organization
- ✅ Easier file management
- ✅ Better file discoverability
- ✅ Scalable system

## Testing the Resources

You can test that resources are properly exposed:

```bash
# Start the MCP server
cd /Users/yihan/Documents/sharepoint\ mcp/wiser-mcp-sharepoint
source venv/bin/activate
python -m mcp_sharepoint.server
```

In your MCP client (like Claude Desktop), the AI can now:
1. List available resources
2. Read `context://file-naming-convention`
3. Use the convention to find files
4. Complete tasks without explicit file names

## Best Practices

### For File Creators
1. Always follow the naming convention strictly
2. Use specific, meaningful content descriptors
3. Include the appropriate time period (week/month/quarter)
4. Always include the year

### For AI Agents
1. **Always read** `context://file-naming-convention` when user provides generic requests
2. **List files** before attempting to match
3. **Use fuzzy matching** when exact matches aren't found
4. **Ask for clarification** when multiple files match equally
5. **Suggest alternatives** when no files match

### For Users
1. Use natural language - the AI will understand
2. Include time periods when relevant (Q3, June, week 25)
3. Specify content type (training, hiring, performance)
4. The AI will ask for clarification if needed

## Troubleshooting

### AI doesn't find the file
- Check that the file follows the naming convention
- Verify the file is in the correct SharePoint folder
- Try being more specific in your request

### Multiple files match
- The AI will list all matches and ask you to choose
- Be more specific about time period or content type

### No files match
- The AI will suggest similar files
- Check if files exist in SharePoint
- Verify file naming follows the convention

## Future Enhancements

Potential improvements:
- Add fuzzy matching for typos
- Support for date ranges (Q1-Q3, Jan-Jun)
- Auto-suggest file names based on partial input
- File naming validation tool
- Automatic file renaming for non-compliant files
