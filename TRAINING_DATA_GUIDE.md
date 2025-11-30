# Training Data Generator Guide

Generate realistic HR recruiting data for testing and development.

## 🚀 Quick Start

```bash
# One command to generate and upload 4 Excel files
source venv/bin/activate && python generate_training_data.py
```

**What happens:**
1. ✅ Generates 4 Excel files (75-100 records each)
2. ✅ Uploads to SharePoint Data folder
3. ✅ Names files with record counts

## 📊 Generated Files

| File | Records | Columns | Purpose |
|------|---------|---------|---------|
| **Filled Positions** | 75-100 | 17 | Candidate hiring data |
| **Time In Step Q3** | 75-100 | 27 | Pipeline time tracking |
| **Recruiting Report** | 75-100 | 19 | Funnel metrics |
| **All Departments** | 75-100 | 24 | Department analytics |

## 📋 Prerequisites

**Dependencies** (should already be installed):
```bash
pip install pandas openpyxl msal requests python-dotenv
```

**Environment** - Ensure `.env` file contains:
```env
SHP_ID_APP=your-app-id
SHP_ID_APP_SECRET=your-app-secret
SHP_TENANT_ID=your-tenant-id
```

## 🎯 Data Characteristics

### File 1: Filled Positions
- **Columns**: Department, Job Code, Employment Type, Job Title, Dates, Candidate Info, Sources, Agency Fees
- **Features**: 10 departments, 15 job titles, realistic agency fees (15% of positions)
- **Date Range**: 2023-2024

### File 2: Time In Step Q3
- **Columns**: Time tracking for each stage (New, In-Review, Interview, Offer), Job details, Personnel
- **Features**: Q3 2023 focus, realistic time-in-stage (1-20 days per stage)
- **Pipeline**: Complete hiring process tracking

### File 3: Recruiting Report
- **Columns**: Job details, Funnel metrics (Applicants, Screens, Interviews, Offers), Sources
- **Features**: Time-to-hire 20-150 days, 50-1,000 applicants per position, realistic conversion rates
- **Geography**: 10 countries, diverse locations

### File 4: All Departments Report
- **Columns**: Department analytics, Application states, Time metrics, Multiple positions
- **Features**: Complete funnel tracking, realistic rejection rates, 30-120 days time-to-hire
- **Structure**: 1-3 positions per requisition

## 🔧 Customization

### Fixed Record Count
```python
# In main() function:
local_path, upload_name = generate_file1_filled_positions(num_records=100)
```

### Different Date Range
```python
# In generator functions:
creation_date = random_date(datetime(2024, 1, 1), datetime(2024, 12, 31))
```

### Higher Application Volumes
```python
# In generate_file3_recruiting_report():
applicants = random.randint(100, 2000)  # Instead of 50-1000
```

### Modify Data Pools
```python
# Edit global variables:
DEPARTMENTS = ['Engineering', 'Sales', 'Marketing', ...]
JOB_TITLES = ['Software Engineer', 'Product Manager', ...]
SOURCES = ['LinkedIn', 'Indeed', 'Internal Referral', ...]
```

## 📁 Output

**Local Files** (temporary):
- `/tmp/generated_file1.xlsx` → `/tmp/generated_file4.xlsx`

**SharePoint Upload** (permanent):
- `Training_Filled_Positions_XX_records.xlsx`
- `Training_Time_In_Step_Q3_XX_records.xlsx`
- `Training_Recruiting_Report_XX_records.xlsx`
- `Training_All_Depts_Report_XX_records.xlsx`

## 🐛 Troubleshooting

### ❌ ModuleNotFoundError
```bash
source venv/bin/activate
pip install pandas openpyxl msal requests python-dotenv
```

### ❌ Authentication Failed
- Check `.env` file has correct credentials
- Verify SharePoint write permissions
- Confirm TENANT_ID, CLIENT_ID, CLIENT_SECRET

### ❌ Upload Failed
- Check internet connection
- Verify SharePoint `Data` folder exists
- Confirm file permissions in SharePoint

### ❌ Import Errors
```bash
# Ensure virtual environment is activated
source venv/bin/activate
```

## 🎯 Use Cases

- ✅ **Testing** - Test MCP analytics tools with fresh data
- ✅ **Demos** - Generate realistic data for presentations  
- ✅ **Training** - Create datasets for AI model training
- ✅ **Development** - Develop features with varied data
- ✅ **QA** - Quality assurance with different scenarios

## 📊 Expected Output

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

## 🔐 Security Notes

- Never commit `.env` file to version control
- Keep SharePoint credentials secure
- Uses Microsoft Graph API with OAuth2 authentication
- Access tokens are temporary and expire after use

## 📝 Data Pools

**Personnel**: 10 departments, 7 recruiters, 6 hiring managers  
**Positions**: 15 job titles, 4 employment types, 5 hire reasons  
**Geography**: 10 countries, 10 cities, 8 US states  
**Sourcing**: 8 sources, 15% agency usage, realistic fees ($5K-$15K)

---

**Status**: ✅ Available  
**Purpose**: Testing and demonstration  
**Usage**: One-time setup (optional)  
**Last Updated**: November 2024
