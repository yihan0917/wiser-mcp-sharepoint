# Excel Column Definitions Dictionary

This file contains definitions for Excel columns that may appear in SharePoint files. When analyzing Excel files, Windsurf will automatically provide context for any columns found in this dictionary.

## How to Add New Definitions

To add new column definitions, simply add a new line in this format:
```markdown
- **Column Name**: Description of what this column contains or represents
```

The system will automatically recognize these columns in any Excel file and provide the context to Windsurf during analysis.

## Recruiting & HR Column Definitions

### Job and Position Information
- **Job Title**: Title of the role or position
- **Job Code**: Unique identifier of position by department
- **Job Ref ID**: Unique identifier of position by department
- **Internal Job Code**: Unique identifier by department
- **Job ID**: Unique identifier by department
- **Job Approval ID**: Unique identifier by department
- **Department**: Department of the opening
- **Department Org Field Value**: Department classification
- **Number of Positions**: How many roles will be filled
- **Employment Type**: Type of employee (full-time, part-time, intern, contractor, etc.)
- **Job Status**: Current status of role (approved, open, sourcing, interviewing, offer, filled)
- **Reason for Hire Request**: New headcount, replacement for resigned employee, backfill for terminated employee

### Location and Geography
- **Country**: Country role is approved to hire in
- **Job Country**: Country role is approved to be hired in
- **Job Location**: Country/location of opening
- **Job City**: City of role, if needs to be in specific location
- **Location**: Location of offered candidate
- **Candidate Location**: Location of candidate

### Personnel and Roles
- **Recruiter**: TA Specialist managing the opening
- **Recruiters**: TA Specialist managing the opening
- **Hiring Manager**: Hiring manager of role/decision maker
- **Hiring Managers**: Hiring manager/decision maker
- **Executives**: Executive assigned to the department
- **Hire Name**: Name of offered candidate
- **Candidate First and Last Name**: Hired candidate's full name
- **Hired Position Incumbent Name**: Name of employee who left

### Dates and Timing
- **Date Opened**: Date the position is posted
- **Date Closed**: Date offer letter is signed
- **Job Creation Date**: Date role was opened in SmartRecruiter, after approval
- **Start Date**: Start date of offered candidate
- **Hired Position Actual Start Date**: Start date of hired candidate
- **Job Status: FILLED Date**: Date role was filled
- **Number of Days Open**: Difference between date role is opened and offer is signed
- **Average time to hire**: Average time for candidate to sign offer
- **Average time to start**: Average time for candidate to start

### Application Process and Metrics
- **Total Number of Applications**: Total number of inbound applications
- **Number of Applications**: Total number of applications
- **Number of Recruiter Screens**: Number of calls recruiter conducted
- **Number of Hiring Manager Screens**: Number of calls hiring manager conducted
- **Number of Final Round Interviews**: Number of interviews that the full interview panel conducted
- **Number of Offers Made**: Number of offers made to get the offer accepted

### Application Status Tracking
- **Number of Applications In Status: In-Review/Resume Review**: Total number of applications in this stage in SmartRecruiter
- **Number of Applications In State: Interview**: Total number of applications in this stage in SmartRecruiter
- **Number of Applications In Status: Interview/Final Interview**: Total number of applications in this stage in SmartRecruiter
- **Number of Applications In State: Withdrawn**: Total number of candidates that backed out at some point in the process
- **Number of Applications In State: Rejected**: Total number of resumes that were rejected at some stage in the process
- **Number of Applications In State: Offered**: Total number of candidates who were presented with an offer
- **Number of Applications In State: Hired**: Total number of candidates who accepted an offer

### Time in Application Stages
- **Time In Application State: New**: Number of days resume sat in this stage
- **Time in Application Status: In-Review/Submitted to Manager**: Number of days resume sat in this stage
- **Time in Application Status: In-Review/Scheduling Recruiter Screen**: Number of days resume sat in this stage
- **Time in Application Status: In-Review/Recruiter Screen**: Number of days resume sat in this stage
- **Time in Application Status: Interview/Hiring Manager Interview**: Number of days resume sat in this stage
- **Time in Application Status: Interview/Technical Interview**: Number of days resume sat in this stage
- **Time in Application Status: Interview/Final Interview**: Number of days resume sat in this stage
- **Time in Application Status: Offer/Offer Pending**: Number of days resume sat in this stage
- **Time in Application Status: Offer/Offer Accepted**: Number of days resume sat in this stage
- **Time In Application State: Offered**: Number of days resume sat in this stage
- **Time in Job Status: SOURCING**: Number of days that a candidate sat in the sourcing stage
- **Time in Job Status: INTERVIEW**: Number of days that a candidate sat in the interview stage

### Sourcing and Costs
- **Source**: Where the candidate applied/sourced from (LinkedIn, website, internal referral, etc.)
- **Candidate Source**: Source of candidate (LinkedIn, website, etc.)
- **Agency Fee/Referral Bonus**: Fee if candidate was through agency or if a referral bonus was paid to internal employee
- **Latest Offer Field: Referral Bonus Amount**: Amount of referral bonus, if applicable
- **Latest Offer Field: Agency Fee**: Amount of agency fee, if applicable
- **Application field: Agency Fee**: Amount of agency fee, if applicable
- **Currency**: Currency type for financial fields

### Application Field Prefixes
*Note: "Application field:" prefix indicates data fields from candidate applications*
- **Application field: Department Name**: Department from application
- **Application field: Employment Type**: Employment type from application
- **Application field: Job Title**: Title of role from application
- **Application field: Reason for Hire**: Reason from application
- **Application field: Start Date - Global Use**: Start date from application
- **Application field: Job Location**: Location from application

## HR Analytics & Reporting Metrics

When generating reports from HR/recruiting data, consider these key metrics and breakdowns. Feel free to suggest additional meaningful metrics based on the available data.

### Core Hiring Metrics
- **Total Hires**: Overall hiring volume
  - By region/country
  - By department
  - By employment type
  - By time period (monthly, quarterly, yearly)

- **Time to Hire**: Recruitment efficiency metrics
  - Average time to hire (days from job posting to offer acceptance)
  - Average time to start (days from offer acceptance to start date)
  - Time in each application stage
  - By region/country and department

### Hiring Effectiveness
- **Offer Acceptance Rate**: Percentage of offers accepted
- **Application Conversion Rates**: 
  - Resume review to interview rate
  - Interview to offer rate
  - Overall application to hire rate
- **Hiring Success by Source**: Which sources yield the best candidates

### Cost Analysis
- **Cost per Hire**: Total recruiting costs divided by number of hires
- **Agency Fees**: Total external recruiting costs
- **Referral Bonus Costs**: Internal referral program costs
- **Cost by Source**: Recruiting costs broken down by candidate source

### Current Pipeline
- **Open Positions**: Current hiring needs
  - By region/country
  - By department
  - By priority/urgency
  - Time positions have been open

### Candidate Experience
- **Application Withdrawal Rate**: Candidates who withdrew from process
- **Time in Each Stage**: How long candidates spend in each step
- **Rejection Reasons**: Why candidates were not selected

### Source Effectiveness
- **Candidate Source Distribution**: Where hires come from
  - LinkedIn, company website, referrals, agencies, etc.
  - Source quality (conversion rates by source)
  - Cost effectiveness by source

### Diversity & Inclusion Metrics
- **Gender Distribution**: 
  - By region/country
  - By department
  - By role level
- **Ethnicity Distribution**:
  - By region/country
  - By department
  - By role level
- **Diversity in Pipeline**: Representation at each hiring stage

### Operational Metrics
- **Recruiter Performance**: Metrics per recruiter
  - Number of hires
  - Time to fill
  - Quality of hires
- **Hiring Manager Engagement**: Interview participation and feedback
- **Process Efficiency**: Bottlenecks and delays in hiring process

### Trend Analysis
- **Seasonal Patterns**: Hiring trends by time of year
- **Department Growth**: Which areas are expanding
- **Turnover Impact**: Replacement vs. new headcount hiring
- **Market Competitiveness**: Offer acceptance rates and time to hire trends