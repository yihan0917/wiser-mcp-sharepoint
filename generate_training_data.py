#!/usr/bin/env python3
"""
Training Data Generator for SharePoint HR Analytics

This script generates realistic random training data for HR recruiting Excel files.
It reads the column structure from existing files in SharePoint and creates
new files with randomized but realistic data.

Usage:
    python generate_training_data.py

Requirements:
    - pandas
    - openpyxl
    - msal
    - requests
    - python-dotenv

Author: Wiser Solutions
Date: November 2025
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import random
import base64
import requests
from msal import ConfidentialClientApplication
from dotenv import load_dotenv
import os
import sys

# Load environment variables
load_dotenv()

# Configuration
CLIENT_ID = os.getenv('SHP_ID_APP')
CLIENT_SECRET = os.getenv('SHP_ID_APP_SECRET')
TENANT_ID = os.getenv('SHP_TENANT_ID')
SITE_ID = "wisersolutionsinc.sharepoint.com,99b9eb9e-666a-407b-bda2-b35e3ea8ceb1,c08228bf-ca6c-426c-a4dc-7d6bcd86b719"
DRIVE_ID = "b!nuu5mWpme0C9orNePqjOsb8ogsBsymxCpNx9a82Gtxmv3DWdneatRYo1CyYclCdc"

# Sample data pools
DEPARTMENTS = ['Engineering', 'Sales', 'Marketing', 'HR', 'Finance', 'Operations', 
               'Product', 'Customer Success', 'IT', 'Legal']

JOB_TITLES = ['Software Engineer', 'Senior Engineer', 'Product Manager', 'Sales Representative', 
              'Marketing Manager', 'HR Business Partner', 'Financial Analyst', 'Operations Manager', 
              'Customer Success Manager', 'Data Analyst', 'DevOps Engineer', 'UX Designer', 
              'Account Executive', 'Content Manager', 'Recruiter']

LOCATIONS = ['US', 'France', 'Canada', 'Australia', 'Mexico', 'UK', 'Germany', 'Poland', 'India', 'Singapore']
CITIES = ['New York', 'San Francisco', 'Paris', 'Toronto', 'Sydney', 'Mexico City', 'London', 'Berlin', 'Warsaw', 'Mumbai']
COUNTRIES = ['United States', 'France', 'Canada', 'Australia', 'Mexico', 'United Kingdom', 'Germany', 'Poland', 'India', 'Singapore']
STATES = ['CA', 'NY', 'TX', 'FL', 'IL', 'WA', 'MA', 'CO']

SOURCES = ['LinkedIn', 'Indeed', 'Internal Referral', 'Agency', 'Company Website', 
           'Glassdoor', 'Internal Promotion', 'Employee Referral']

RECRUITERS = ['Sarah Johnson', 'Mike Chen', 'Emily Rodriguez', 'David Kim', 
              'Lisa Anderson', 'James Wilson', 'Maria Garcia']

HIRING_MANAGERS = ['John Smith', 'Jennifer Lee', 'Robert Brown', 'Michelle Davis', 
                   'William Taylor', 'Amanda Martinez']

EMPLOYMENT_TYPES = ['Full-time', 'Part-time', 'Contract', 'Intern']
JOB_STATUSES = ['FILLED', 'SOURCING', 'INTERVIEW', 'OFFER', 'OPEN']
HIRE_REASONS = ['New headcount', 'Replacement', 'Backfill', 'Growth', 'Expansion']


def random_date(start_date, end_date):
    """Generate random date between start and end"""
    delta = end_date - start_date
    random_days = random.randint(0, delta.days)
    return start_date + timedelta(days=random_days)


def random_name():
    """Generate random person name"""
    first_names = ['John', 'Jane', 'Michael', 'Sarah', 'David', 'Emily', 'Robert', 
                   'Lisa', 'James', 'Maria', 'William', 'Jennifer', 'Richard', 'Amanda']
    last_names = ['Smith', 'Johnson', 'Williams', 'Brown', 'Jones', 'Garcia', 'Miller', 
                  'Davis', 'Rodriguez', 'Martinez', 'Wilson', 'Anderson', 'Taylor', 'Thomas']
    return f"{random.choice(first_names)} {random.choice(last_names)}"


def get_access_token():
    """Get Microsoft Graph API access token"""
    app = ConfidentialClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        client_credential=CLIENT_SECRET
    )
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    return result['access_token']


def upload_to_sharepoint(local_path, upload_name, access_token):
    """Upload file to SharePoint Data folder"""
    with open(local_path, 'rb') as f:
        file_content = f.read()
    
    upload_url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/drives/{DRIVE_ID}/root:/Data/{upload_name}:/content"
    
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    }
    
    response = requests.put(upload_url, headers=headers, data=file_content)
    
    if response.status_code in [200, 201]:
        return True, len(file_content)
    else:
        return False, response.text


def generate_file1_filled_positions(num_records=90):
    """Generate File 1: Filled Position (Candidate Data)"""
    print(f"\n📊 Generating File 1: Filled Positions ({num_records} records)...")
    
    data = []
    for i in range(num_records):
        creation_date = random_date(datetime(2023, 1, 1), datetime(2024, 12, 31))
        start_date = creation_date + timedelta(days=random.randint(30, 120))
        actual_start = start_date + timedelta(days=random.randint(-5, 10))
        
        source = random.choice(SOURCES)
        
        data.append({
            'Application field: Department Name': random.choice(DEPARTMENTS),
            'Internal Job Code': f"JOB-{random.randint(1000, 9999)}",
            'Application field: Employment Type': random.choice(EMPLOYMENT_TYPES),
            'Job Status': 'FILLED',
            'Application field: Job Title': random.choice(JOB_TITLES),
            'Job Creation Date': creation_date.strftime('%Y-%m-%d'),
            'Application field: Reason for Hire': random.choice(HIRE_REASONS),
            'Hired Position Incumbent Name': random_name() if random.random() > 0.7 else '',
            'Currency': 'USD',
            'Candidate First and Last Name': random_name(),
            'Candidate Location': random.choice(LOCATIONS),
            'Application field: Start Date - Global Use': start_date.strftime('%Y-%m-%d'),
            'Hired Position Actual Start Date': actual_start.strftime('%Y-%m-%d'),
            'Candidate Source': source,
            'Latest Offer Field: Referral Bonus Amount': random.choice([0, 1000, 2000, 3000]) if 'Referral' in source else 0,
            'Latest Offer Field: Agency Fee': random.choice([0, 5000, 8000, 12000, 15000]) if random.random() > 0.85 else 0,
            'Application field: Agency Fee': random.choice([0, 5000, 8000, 12000, 15000]) if random.random() > 0.85 else 0
        })
    
    df = pd.DataFrame(data)
    output_path = '/tmp/generated_file1.xlsx'
    df.to_excel(output_path, index=False)
    print(f"   ✅ Generated {num_records} records")
    return output_path, f'Training_Filled_Positions_{num_records}_records.xlsx'


def generate_file2_time_in_step(num_records=97):
    """Generate File 2: Global Time In Step - Q3 Report"""
    print(f"\n📊 Generating File 2: Time In Step Q3 ({num_records} records)...")
    
    data = []
    for i in range(num_records):
        creation_date = random_date(datetime(2023, 7, 1), datetime(2023, 9, 30))  # Q3
        filled_date = creation_date + timedelta(days=random.randint(30, 150))
        
        data.append({
            'Time In Application State: New': random.randint(1, 10),
            'Time in Application Status: In-Review/Submitted to Manager': random.randint(2, 15),
            'Time in Application Status: In-Review/Scheduling Recruiter Screen': random.randint(1, 7),
            'Time in Application Status: In-Review/Recruiter Screen': random.randint(3, 10),
            'Time in Application Status: Interview/Hiring Manager Interview': random.randint(5, 20),
            'Time in Application Status: Interview/Technical Interview': random.randint(3, 15),
            'Time in Application Status: Interview/Final Interview': random.randint(5, 15),
            'Time in Application Status: Offer/Offer Pending': random.randint(2, 10),
            'Time in Application Status: Offer/Offer Accepted': random.randint(1, 7),
            'Time In Application State: Offered': random.randint(3, 14),
            'Application field: Job Title': random.choice(JOB_TITLES),
            'Application field: Job Location': random.choice(LOCATIONS),
            'Application field: Department Name': random.choice(DEPARTMENTS),
            'Job Approval ID': f"APPR-{random.randint(10000, 99999)}",
            'Job Creation Date': creation_date.strftime('%Y-%m-%d'),
            'Department Org Field Value': random.choice(DEPARTMENTS),
            'Job Status: FILLED Date': filled_date.strftime('%Y-%m-%d'),
            'Executives': random.choice(HIRING_MANAGERS),
            'Hiring Managers': random.choice(HIRING_MANAGERS),
            'Job ID': f"JOB-{random.randint(1000, 9999)}",
            'Job Location': random.choice(LOCATIONS),
            'Job Location State': random.choice(STATES),
            'Job Ref ID': f"REF-{random.randint(100000, 999999)}",
            'Job Title': random.choice(JOB_TITLES),
            'Job City': random.choice(CITIES),
            'Job Country': random.choice(COUNTRIES),
            'Job Status': 'FILLED'
        })
    
    df = pd.DataFrame(data)
    output_path = '/tmp/generated_file2.xlsx'
    df.to_excel(output_path, index=False)
    print(f"   ✅ Generated {num_records} records")
    return output_path, f'Training_Time_In_Step_Q3_{num_records}_records.xlsx'


def generate_file3_recruiting_report(num_records=86):
    """Generate File 3: Recruiting Report Columns"""
    print(f"\n📊 Generating File 3: Recruiting Report ({num_records} records)...")
    
    data = []
    for i in range(num_records):
        date_opened = random_date(datetime(2023, 1, 1), datetime(2024, 6, 30))
        days_open = random.randint(20, 150)
        date_closed = date_opened + timedelta(days=days_open)
        start_date = date_closed + timedelta(days=random.randint(10, 30))
        
        applicants = random.randint(50, 1000)
        recruiter_screens = int(applicants * random.uniform(0.02, 0.08))
        manager_screens = int(recruiter_screens * random.uniform(0.3, 0.7))
        final_interviews = int(manager_screens * random.uniform(0.4, 0.8))
        offers = random.randint(1, 3)
        
        data.append({
            'Job Title ': random.choice(JOB_TITLES),
            'Job Code ': f"JC-{random.randint(1000, 9999)}",
            'Recruiter': random.choice(RECRUITERS),
            'Hiring Manger ': random.choice(HIRING_MANAGERS),
            'Employment type': random.choice(EMPLOYMENT_TYPES),
            'Country': random.choice(COUNTRIES),
            'Date Opened': date_opened.strftime('%Y-%m-%d'),
            'Date Closed': date_closed.strftime('%Y-%m-%d'),
            'Total Days Open ': days_open,
            'Total Number of Inbound Applicants': applicants,
            'Number of Recruiter Screen': recruiter_screens,
            'Number of Hiring Manager Screens': manager_screens,
            'Number of Final Roun Interviews': final_interviews,
            'Number of Offers Made': offers,
            'Start Date': start_date.strftime('%Y-%m-%d'),
            'Location': random.choice(LOCATIONS),
            'Hire Name': random_name(),
            'Source ': random.choice(SOURCES),
            'Agency Fee/Referral Bonus ': random.choice([0, 0, 0, 0, 5000, 8000, 12000, 15000])
        })
    
    df = pd.DataFrame(data)
    output_path = '/tmp/generated_file3.xlsx'
    df.to_excel(output_path, index=False)
    print(f"   ✅ Generated {num_records} records")
    return output_path, f'Training_Recruiting_Report_{num_records}_records.xlsx'


def generate_file4_all_depts(num_records=91):
    """Generate File 4: SR Report - All Departments"""
    print(f"\n📊 Generating File 4: All Departments Report ({num_records} records)...")
    
    data = []
    for i in range(num_records):
        creation_date = random_date(datetime(2023, 1, 1), datetime(2024, 12, 31))
        filled_date = creation_date + timedelta(days=random.randint(30, 120))
        time_to_hire = (filled_date - creation_date).days
        time_to_start = time_to_hire + random.randint(10, 30)
        
        num_applications = random.randint(30, 500)
        in_review = int(num_applications * random.uniform(0.1, 0.3))
        in_interview = int(num_applications * random.uniform(0.05, 0.15))
        final_interview = int(in_interview * random.uniform(0.3, 0.6))
        withdrawn = int(num_applications * random.uniform(0.05, 0.15))
        rejected = int(num_applications * random.uniform(0.4, 0.7))
        offered = random.randint(1, 3)
        hired = 1
        
        data.append({
            'Job Ref ID': f"REF-{random.randint(100000, 999999)}",
            'Department Org Field Value': random.choice(DEPARTMENTS),
            'Number of Positions': random.randint(1, 3),
            'Reason for Hire Request': random.choice(HIRE_REASONS),
            'Job Title': random.choice(JOB_TITLES),
            'Job Status': random.choice(JOB_STATUSES),
            'Recruiters': random.choice(RECRUITERS),
            'Hiring Managers': random.choice(HIRING_MANAGERS),
            'Job Country': random.choice(COUNTRIES),
            'Job Creation Date': creation_date.strftime('%Y-%m-%d'),
            'Job Status: FILLED Date': filled_date.strftime('%Y-%m-%d') if random.random() > 0.3 else '',
            'Average Time-To-Hire': time_to_hire if random.random() > 0.3 else '',
            'Average Time-To-Start': time_to_start if random.random() > 0.3 else '',
            'Number of Applications': num_applications,
            'Number of Applications In Status: In-Review/Resume Review': in_review,
            'Number of Applications In State: Interview': in_interview,
            'Number of Applications In Status: Interview/Final Interview': final_interview,
            'Number of Applications In State: Withdrawn': withdrawn,
            'Number of Applications In State: Rejected': rejected,
            'Number of Applications In State: Offered': offered,
            'Number of Applications In State: Hired': hired,
            'Time in Job Status: SOURCING': random.randint(5, 30),
            'Time in Job Status: INTERVIEW': random.randint(10, 45),
            'Unnamed: 23': ''
        })
    
    df = pd.DataFrame(data)
    output_path = '/tmp/generated_file4.xlsx'
    df.to_excel(output_path, index=False)
    print(f"   ✅ Generated {num_records} records")
    return output_path, f'Training_All_Depts_Report_{num_records}_records.xlsx'


def main():
    """Main execution function"""
    print("=" * 80)
    print("TRAINING DATA GENERATOR FOR HR ANALYTICS")
    print("=" * 80)
    print("\n🔄 Generating training datasets...")
    
    # Generate all files
    files_to_upload = []
    
    try:
        # File 1: Filled Positions
        local_path, upload_name = generate_file1_filled_positions(num_records=random.randint(75, 100))
        files_to_upload.append((local_path, upload_name))
        
        # File 2: Time In Step
        local_path, upload_name = generate_file2_time_in_step(num_records=random.randint(75, 100))
        files_to_upload.append((local_path, upload_name))
        
        # File 3: Recruiting Report
        local_path, upload_name = generate_file3_recruiting_report(num_records=random.randint(75, 100))
        files_to_upload.append((local_path, upload_name))
        
        # File 4: All Departments
        local_path, upload_name = generate_file4_all_depts(num_records=random.randint(75, 100))
        files_to_upload.append((local_path, upload_name))
        
        print("\n" + "=" * 80)
        print("✅ ALL FILES GENERATED SUCCESSFULLY!")
        print("=" * 80)
        
        # Upload to SharePoint
        print("\n📤 Uploading to SharePoint...")
        print("=" * 80)
        
        access_token = get_access_token()
        
        for local_path, upload_name in files_to_upload:
            success, result = upload_to_sharepoint(local_path, upload_name, access_token)
            
            if success:
                print(f"✅ {upload_name}")
                print(f"   Size: {result:,} bytes")
            else:
                print(f"❌ {upload_name}: Upload failed")
                print(f"   {result[:200]}")
        
        print("\n" + "=" * 80)
        print("✅ UPLOAD COMPLETE!")
        print("=" * 80)
        print("\n📁 Files uploaded to SharePoint Data folder")
        print("🎯 Ready for HR analytics and training!")
        
    except Exception as e:
        print(f"\n❌ Error: {str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
