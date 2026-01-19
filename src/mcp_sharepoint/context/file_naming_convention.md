# File Naming Convention for Excel Files

## Purpose
This document defines the standardized naming convention for Excel files stored in SharePoint. This convention enables AI agents to automatically identify and select appropriate files for analysis based on user requests without requiring explicit file names.

## Convention Structure

```
{content}_{source}_{time_period}_{year}.xlsx
```

### Components (in order):

1. **content** (REQUIRED)
   - A specific, meaningful phrase representing the main purpose/content of the Excel file
   - Use descriptive, distinctive terms that clearly identify the data type
   - Good examples: "Engineering_Hiring_Pipeline", "Q3_Performance_Reviews", "Leadership_Development_Training", "Voluntary_Exit_Interviews"
   - Avoid generic terms: Instead of "Training" use "Leadership_Development_Training" or "Technical_Skills_Training"
   - Use underscores for multi-word content

2. **source** (REQUIRED)
   - Indicates where the data originated
   - Valid values:
     - `Manual` - Files created manually by staff
     - `SmartRecruiter` - Files downloaded from SmartRecruiter platform
     - Other platform names as needed (e.g., `Workday`, `BambooHR`, `ADP`)

3. **time_period** (CONDITIONAL - use ONE of the following if applicable):
   - **Week**: `WK##` - For data covering a specific fiscal week (e.g., `WK25`, `WK48`)
   - **Month**: Full month name - For data covering a specific month (e.g., `January`, `June`, `December`)
   - **Quarter**: `Q#` - For data covering a specific quarter (e.g., `Q1`, `Q2`, `Q3`, `Q4`)
   - **Omit if**: File contains data spanning multiple periods or the entire year

4. **year** (REQUIRED)
   - Four-digit year (e.g., `2024`, `2025`)

5. **extension** (REQUIRED)
   - Always `.xlsx` for Excel files

## Examples

### Weekly Data
```
New_Hire_Onboarding_SmartRecruiter_WK25_2024.xlsx
Software_Engineer_Hiring_Manual_WK48_2025.xlsx
```

### Monthly Data
```
Q3_Performance_Reviews_Manual_June_2024.xlsx
Engineering_Headcount_Workday_December_2025.xlsx
Voluntary_Exit_Interviews_Manual_March_2024.xlsx
```

### Quarterly Data
```
Engineering_Hiring_Pipeline_SmartRecruiter_Q3_2024.xlsx
Leadership_Development_Training_Manual_Q1_2025.xlsx
Employee_Retention_Analysis_Manual_Q4_2024.xlsx
```

### Annual or Multi-Period Data
```
All_Open_Positions_SmartRecruiter_2024.xlsx
Company_Wide_Performance_Review_Manual_2025.xlsx
Department_Headcount_Trends_Manual_2024.xlsx
```

## AI Agent Usage Guidelines

### Automatic File Detection
When a user makes a request without specifying a file name, the AI agent should:

1. **Parse the request** for key indicators:
   - Content type: "training", "hiring", "performance", "headcount", etc.
   - Time period: "Q3", "third quarter", "June", "week 25", "2024", etc.
   - Data source: "SmartRecruiter data", "manual report", etc.

2. **Construct search patterns** based on the request:
   - If user asks about "Q3 2024 training": Look for files matching `*Training*Q3*2024.xlsx`
   - If user asks about "June hiring data": Look for files matching `*Hiring*June*.xlsx`
   - If user asks about "SmartRecruiter data for 2024": Look for files matching `*SmartRecruiter*2024.xlsx`

3. **Prioritize matches**:
   - Exact matches on all components (content + source + time_period + year)
   - Partial matches on content + time_period + year
   - Broader matches on content + year
   - Ask for clarification if multiple files match equally

### Example User Requests and File Detection

| User Request | Expected File Pattern | Example Matches |
|--------------|----------------------|-----------------|
| "Analyze Q3 leadership training data" | `*Leadership*Training*Q3*.xlsx` | `Leadership_Development_Training_Manual_Q1_2025.xlsx` |
| "Show me June performance review metrics" | `*Performance*Review*June*.xlsx` | `Q3_Performance_Reviews_Manual_June_2024.xlsx` |
| "Review week 25 onboarding data" | `*Onboarding*WK25*.xlsx` | `New_Hire_Onboarding_SmartRecruiter_WK25_2024.xlsx` |
| "Analyze 2024 SmartRecruiter hiring data" | `*Hiring*SmartRecruiter*2024.xlsx` | `Engineering_Hiring_Pipeline_SmartRecruiter_Q3_2024.xlsx` |
| "Check manual retention reports for Q4" | `*Retention*Manual*Q4*.xlsx` | `Employee_Retention_Analysis_Manual_Q4_2024.xlsx` |

### Handling Ambiguity

If multiple files match the user's request:
1. List all matching files
2. Ask the user to clarify which specific file(s) to analyze
3. Provide context about each file (size, last modified date, etc.)

If no files match the user's request:
1. List available files that partially match
2. Suggest the closest alternatives
3. Ask if the user wants to analyze a different file or time period

## Best Practices

### For File Creators
- Always follow the naming convention strictly
- Use consistent terminology for content types
- Ensure the year is always the last component before `.xlsx`
- Use underscores, not spaces or hyphens, as separators

### For AI Agents
- Always check the file list before making assumptions
- Use fuzzy matching when exact matches aren't found
- Consider the current date when inferring years (e.g., "Q3" likely means current year)
- Validate file contents match the expected structure before analysis

## Migration Notes

For existing files that don't follow this convention:
1. The AI agent should still attempt to work with them
2. Suggest renaming to the standard convention
3. Document any non-standard files for future reference
