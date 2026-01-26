# Training Data Generator for HR Analytics

This script generates realistic random training data for HR recruiting Excel files and uploads them to SharePoint.

## 📋 Overview

The `generate_training_data.py` script creates 4 Excel files with randomized but realistic HR recruiting data:

1. **Filled Positions** - Candidate and hiring data for filled positions
2. **Time In Step Q3** - Time tracking through recruiting pipeline stages
3. **Recruiting Report** - Comprehensive recruiting metrics and funnel data
4. **All Departments Report** - Department-level recruiting analytics

## 🚀 Quick Start

### Prerequisites

Make sure you have the required dependencies installed:

```bash
# Activate virtual environment
source venv/bin/activate

# Dependencies should already be installed from main project
# If not, install them:
pip install pandas openpyxl msal requests python-dotenv
```

### Environment Setup

Ensure your `.env` file contains the required SharePoint credentials:

```env
SHP_ID_APP=your-app-id
SHP_ID_APP_SECRET=your-app-secret
SHP_TENANT_ID=your-tenant-id
```

### Run the Script

```bash
# From the project root directory
python generate_training_data.py
```

## 📊 Generated Data

### File 1: Filled Positions (75-100 records)

**Columns (17):**
- Application field: Department Name
- Internal Job Code
- Application field: Employment Type
- Job Status
- Application field: Job Title
- Job Creation Date
- Application field: Reason for Hire
- Hired Position Incumbent Name
- Currency
- Candidate First and Last Name
- Candidate Location
- Application field: Start Date - Global Use
- Hired Position Actual Start Date
- Candidate Source
- Latest Offer Field: Referral Bonus Amount
- Latest Offer Field: Agency Fee
- Application field: Agency Fee

**Data Characteristics:**
- Dates: 2023-2024
- Departments: 10 different departments
- Job Titles: 15 different roles
- Sources: LinkedIn, Indeed, Referrals, etc.
- Agency Fees: 15% of positions use agencies ($5K-$15K)

### File 2: Time In Step Q3 (75-100 records)

**Columns (27):**
- Time tracking for each application state (New, In-Review, Interview, Offer)
- Job details (Title, Location, Department, Country, City)
- Personnel (Executives, Hiring Managers)
- IDs (Job ID, Approval ID, Ref ID)
- Dates (Creation Date, Filled Date)

**Data Characteristics:**
- Q3 2023 focus (July-September)
- Realistic time-in-stage metrics (1-20 days per stage)
- Complete hiring pipeline tracking

### File 3: Recruiting Report (75-100 records)

**Columns (19):**
- Job Title, Job Code, Recruiter, Hiring Manager
- Employment Type, Country, Location
- Dates (Opened, Closed, Total Days Open)
- Funnel metrics (Applicants, Screens, Interviews, Offers)
- Hire Name, Source, Agency Fees

**Data Characteristics:**
- Time-to-hire: 20-150 days
- Application volumes: 50-1,000 per position
- Realistic conversion rates (2-8% screen rate)
- Geographic diversity across 10 countries

### File 4: All Departments Report (75-100 records)

**Columns (24):**
- Job Ref ID, Department, Number of Positions
- Reason for Hire, Job Status
- Recruiters, Hiring Managers, Country
- Time-to-Hire and Time-to-Start metrics
- Application state distribution (In-Review, Interview, Withdrawn, Rejected, Offered, Hired)
- Time in each job status

**Data Characteristics:**
- Complete application funnel tracking
- Realistic rejection/withdrawal rates
- Time-to-hire: 30-120 days
- Multiple positions per requisition (1-3)

## 🎯 Data Pools Used

### Personnel
- **Departments**: Engineering, Sales, Marketing, HR, Finance, Operations, Product, Customer Success, IT, Legal
- **Recruiters**: 7 different recruiters
- **Hiring Managers**: 6 different hiring managers

### Positions
- **Job Titles**: 15 different roles (Software Engineer, Product Manager, Sales Rep, etc.)
- **Employment Types**: Full-time, Part-time, Contract, Intern
- **Hire Reasons**: New headcount, Replacement, Backfill, Growth, Expansion

### Geography
- **Countries**: 10 countries (US, France, Canada, Australia, Mexico, UK, Germany, Poland, India, Singapore)
- **Cities**: 10 major cities
- **US States**: 8 states (CA, NY, TX, FL, IL, WA, MA, CO)

### Sourcing
- **Sources**: LinkedIn, Indeed, Internal Referral, Agency, Company Website, Glassdoor, Internal Promotion, Employee Referral
- **Agency Usage**: 15% of positions
- **Agency Fees**: $5,000 - $15,000
- **Referral Bonuses**: $1,000 - $3,000

## 🔧 Customization

### Adjust Number of Records

Edit the `main()` function to specify exact record counts:

```python
# Instead of random.randint(75, 100), use a fixed number:
local_path, upload_name = generate_file1_filled_positions(num_records=100)
```

### Modify Data Pools

Edit the global variables at the top of the script:

```python
DEPARTMENTS = ['Engineering', 'Sales', 'Marketing', ...]  # Add/remove departments
JOB_TITLES = ['Software Engineer', ...]  # Add/remove job titles
SOURCES = ['LinkedIn', 'Indeed', ...]  # Add/remove sources
```

### Change Date Ranges

Modify the `random_date()` calls in each generator function:

```python
# Example: Change to 2024 data only
creation_date = random_date(datetime(2024, 1, 1), datetime(2024, 12, 31))
```

### Adjust Metrics

Modify the calculation logic in each generator:

```python
# Example: Increase application volumes
applicants = random.randint(100, 2000)  # Instead of 50-1000

# Example: Change conversion rates
recruiter_screens = int(applicants * random.uniform(0.05, 0.10))  # Instead of 0.02-0.08
```

## 📁 Output

### Local Files
Generated files are temporarily saved to `/tmp/`:
- `/tmp/generated_file1.xlsx`
- `/tmp/generated_file2.xlsx`
- `/tmp/generated_file3.xlsx`
- `/tmp/generated_file4.xlsx`

### SharePoint Upload
Files are automatically uploaded to your SharePoint `Data` folder with names:
- `Training_Filled_Positions_XX_records.xlsx`
- `Training_Time_In_Step_Q3_XX_records.xlsx`
- `Training_Recruiting_Report_XX_records.xlsx`
- `Training_All_Depts_Report_XX_records.xlsx`

(Where XX is the number of records generated)

## 🐛 Troubleshooting

### Authentication Errors

If you get authentication errors:
1. Check your `.env` file has correct credentials
2. Verify the credentials have SharePoint write permissions
3. Ensure TENANT_ID, CLIENT_ID, and CLIENT_SECRET are correct

### Upload Failures

If files generate but don't upload:
1. Check internet connection
2. Verify SharePoint site is accessible
3. Confirm the `Data` folder exists in SharePoint
4. Check file permissions in SharePoint

### Module Not Found Errors

If you get import errors:
```bash
# Make sure virtual environment is activated
source venv/bin/activate

# Install missing dependencies
pip install pandas openpyxl msal requests python-dotenv
```

## 🔐 Security Notes

- Never commit the `.env` file to version control
- Keep your SharePoint credentials secure
- The script uses Microsoft Graph API with OAuth2 authentication
- Access tokens are temporary and expire after use

## 📝 Use Cases

This script is perfect for:
- ✅ **Testing** - Test your MCP analytics tools with fresh data
- ✅ **Demos** - Generate realistic data for presentations
- ✅ **Training** - Create datasets for training AI models
- ✅ **Development** - Develop new features with varied data
- ✅ **QA** - Quality assurance testing with different scenarios

## 🎓 Learning Resources

To understand the data structure better:
1. Review `column_definitions.md` for column meanings
2. Check the original Excel files in SharePoint Data folder
3. Run analytics on generated data to see patterns
4. Compare generated vs. real data metrics

## 📞 Support

For issues or questions:
1. Check the main `SETUP_GUIDE.md`
2. Review SharePoint MCP server documentation
3. Verify environment configuration
4. Check SharePoint permissions

---

**Last Updated**: November 2025  
**Version**: 1.0  
**Author**: Wiser Solutions
