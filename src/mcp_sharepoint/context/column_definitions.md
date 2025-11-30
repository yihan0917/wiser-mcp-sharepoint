---
FILE: column_definitions.md
PURPOSE: Define Excel column meanings for HR and recruiting data analysis
DEPARTMENT: HR/Recruiting
APPLIES_TO: HR analysts, recruiting team, data analysts
KEY_SECTIONS: Recruiter Reports, Personnel Data, Application Tracking, Timing Metrics, Sourcing, Costs
COMMON_SEARCHES: column definitions, Excel columns, HR data, recruiting metrics, job title, hiring manager, time to hire, source, application tracking, personnel data
RELATED_FILES: metrics_definitions.md, hiring_guide.md
LAST_UPDATED: 2024-11-30
---

# Excel Column Definitions Dictionary

This file contains definitions for Excel columns that may appear in SharePoint files. When analyzing Excel files, the system will automatically provide context for any columns found in this dictionary.

---

## Recruiter Report Columns (Manual Report)

- **Job Title** - Title of role
- **Job Code** - Unique identifier of position by department
- **Recruiter** - TA Specialist managing the opening
- **Hiring Manager** - Hiring manager of role/decision maker
- **Employment Type** - Type of employee (full time, intern, etc.)
- **Country** - Country role is approved to hire in
- **Date Opened** - Date the position is posted
- **Date Closed** - Date offer letter is signed
- **Number of Days Open** - Difference between date role is opened and offer is signed
- **Total Number of Applications** - Total number of in-bound applications
- **Number of Recruiter Screens** - Number of calls recruiter did
- **Number of Hiring Manager Screens** - Number of calls hiring manager did
- **Number of Final Round Interviews** - Number of interviews that the full interview panel did
- **Number of Offers Made** - Number of offers made to get the offer accepted
- **Start Date** - Start date of offered candidate
- **Location** - Location of offered candidate
- **Hire Name** - Name of offered candidate
- **Source** - Where the candidate applied/sourced from (LinkedIn, website, etc., internal referral)
- **Agency Fee/Referral Bonus** - Fee if candidate was through agency or if a referral bonus was paid to internal employee

---

## Filled Position (Candidate Data) (SmartRecruiter Report)

- **Job Ref ID** - Unique identifier of position by department
- **Department** - Department of opening
- **Number of Positions** - How many roles will be filled
- **Reason for Hire Request** - New headcount, replacement for resigned employee, backfill for terminated employee
- **Job Title** - Title of role
- **Job Status** - Approved, open, interviewing, filled
- **Recruiters** - TA Specialist managing the opening
- **Hiring Manager** - Hiring manager of the role/decision maker
- **Job Country** - Country role is approved to hire in
- **Job Creation Date** - Date role was opened in SmartRecruiter, after approval
- **Average Time to Hire** - Average time for candidate to sign offer
- **Average Time to Start** - Average time for candidate to start
- **Number of Applications** - Total number of applications
- **Number of Applications In Status: In-Review/Resume Review** - Total number of applications that are in this stage in SR
- **Number of Applications In State: Interview** - Total number of applications that are in this stage in SR
- **Number of Applications In Status: Interview/Final Interview** - Total number of applications that are in this stage in SR
- **Number of Applications In State: Withdrawn** - Total number of candidates that backed out at some point in the process
- **Number of Applications In State: Rejected** - Total number of resumes that were rejected at some stage in the process
- **Number of Applications In State: Offered** - Total number of candidates who were presented with an offer
- **Number of Applications In State: Hired** - Total number of candidates who accepted an offer
- **Time in Job Status: SOURCING** - Number of days that a candidate sat in the sourcing stage
- **Time in Job Status: INTERVIEW** - Number of days that a candidate sat in the interview stage

---

## Filled-Position (Candidate Data) (SmartRecruiter Report)

- **Application field: Department Name** - Department
- **Internal Job Code** - Unique identifier by department
- **Application field: Employment Type** - Employment type (intern, part-time, full-time, etc.)
- **Job Status** - Status of role (open, sourcing, interviewing, filled)
- **Application field: Job Title** - Title of role
- **Job Creation Date** - Date role was added to SR
- **Application field: Reason for Hire** - New Headcount, Backfill for terminated employee, replacement for resigned employee
- **Hired Position Incumbent Name** - Name of employee who left
- **Currency** - Currency type
- **Candidate First and Last Name** - Hired candidate's full name
- **Candidate Location** - Location of candidate
- **Application field: Start Date - Global Use** - Start date of hired candidate
- **Hired Position Actual Start Date** - Start date of hired candidate
- **Candidate Source** - Source of candidate (LinkedIn, website, etc.)
- **Latest Offer Field: Referral Bonus Amount** - Amount of referral bonus, if applicable
- **Latest Offer Field: Agency Fee** - Amount of agency fee, if applicable
- **Application field: Agency Fee** - Amount of agency fee, if applicable

---

## Global-Time-In-Step (SmartRecruiter Report)

**Note:** This report is by candidate versus by job

### Time Tracking Columns

- **Time In Application State: New** - Number of days resume sat in this stage
- **Time in Application Status: In-Review/Submitted to Manager** - Number of days resume sat in this stage
- **Time in Application Status: In-Review/Scheduling Recruiter Screen** - Number of days resume sat in this stage
- **Time in Application Status: In-Review/Recruiter Screen** - Number of days resume sat in this stage
- **Time in Application Status: Interview/Hiring Manager Interview** - Number of days resume sat in this stage
- **Time in Application Status: Interview/Technical Interview** - Number of days resume sat in this stage
- **Time in Application Status: Interview/Final Interview** - Number of days resume sat in this stage
- **Time in Application Status: Offer/Offer Pending** - Number of days resume sat in this stage
- **Time in Application Status: Offer/Offer Accepted** - Number of days resume sat in this stage
- **Time In Application State: Offered** - Number of days resume sat in this stage

### Job Information Columns

- **Application field: Job Title** - Job title
- **Application field: Job Location** - Location of opening
- **Application field: Department Name** - Department
- **Job Approval ID** - Unique identifier by department
- **Job Creation Date** - Date role was opened in SR
- **Department Org Field Value** - Department
- **Job Status: FILLED Date** - Date role was filled
- **Executives** - Executive assigned to the department
- **Hiring Managers** - Hiring manager/decision maker
- **Job ID** - Unique identifier by department
- **Job Location** - Country of opening
- **Job Ref ID** - Unique identifier by department
- **Job Title** - Title of the role
- **Job City** - City of role, if needs to be in specific location
- **Job Country** - Country role is approved to be hired in
- **Job Status** - Open, sourcing, interviewing, offer, filled 