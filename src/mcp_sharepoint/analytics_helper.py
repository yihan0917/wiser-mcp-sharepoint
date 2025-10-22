"""
Analytics helper for HR data analysis and validation
"""
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from typing import Dict, List, Any, Optional, Tuple
from .common import logger
import re
import io

class HRAnalytics:
    """Helper class for HR data analytics and validation"""
    
    def __init__(self):
        self.date_columns = [
            'Date Opened', 'Date Closed', 'Start Date', 'Job Creation Date',
            'Job Status: FILLED Date', 'Hired Position Actual Start Date'
        ]
        self.numeric_columns = [
            'Total Days Open', 'Number of Days Open', 'Inbound Applicants', 
            'Total Number of Applications', 'Number of Applications',
            'Recruiter Screen', 'Number of Recruiter Screens',
            'Hiring Manager Screen', 'Number of Hiring Manager Screens',
            'Final Round Interviews', 'Number of Final Round Interviews',
            'Offers', 'Number of Offers Made', 'Agency Fee'
        ]
        self.required_columns = [
            'Job Title', 'Role Name', 'Recruiter', 'Hiring Manager'
        ]
    
    def parse_excel_content(self, content: str) -> pd.DataFrame:
        """Parse Excel content string into pandas DataFrame"""
        try:
            lines = content.split('\n')
            data_lines = []
            headers = None
            
            for line in lines:
                if line.startswith('HEADERS:'):
                    # Extract headers
                    headers_text = line[8:].strip()
                    headers = [col.strip() for col in headers_text.split('|')]
                elif line.startswith('---'):
                    continue
                elif line.strip() and not line.startswith('===') and headers:
                    # Data row
                    values = [val.strip() for val in line.split('|')]
                    if len(values) == len(headers):
                        data_lines.append(values)
            
            if headers and data_lines:
                df = pd.DataFrame(data_lines, columns=headers)
                return self._clean_dataframe(df)
            else:
                return pd.DataFrame()
                
        except Exception as e:
            logger.error(f"Error parsing Excel content: {e}")
            return pd.DataFrame()
    
    def _clean_dataframe(self, df: pd.DataFrame) -> pd.DataFrame:
        """Clean and standardize DataFrame"""
        try:
            # Convert date columns
            for col in df.columns:
                if any(date_col in col for date_col in self.date_columns):
                    df[col] = pd.to_datetime(df[col], errors='coerce')
            
            # Convert numeric columns
            for col in df.columns:
                if any(num_col in col for num_col in self.numeric_columns):
                    df[col] = pd.to_numeric(df[col], errors='coerce')
            
            # Clean text columns
            text_columns = df.select_dtypes(include=['object']).columns
            for col in text_columns:
                df[col] = df[col].astype(str).str.strip()
                df[col] = df[col].replace('', np.nan)
            
            return df
            
        except Exception as e:
            logger.error(f"Error cleaning DataFrame: {e}")
            return df
    
    def validate_data_quality(self, df: pd.DataFrame) -> Dict[str, Any]:
        """Validate data quality and identify issues"""
        try:
            issues = {
                'missing_data': {},
                'data_anomalies': {},
                'format_issues': {},
                'summary': {}
            }
            
            total_rows = len(df)
            
            # Check missing data
            for col in df.columns:
                missing_count = df[col].isnull().sum()
                missing_pct = (missing_count / total_rows) * 100 if total_rows > 0 else 0
                
                if missing_pct > 0:
                    issues['missing_data'][col] = {
                        'count': int(missing_count),
                        'percentage': round(missing_pct, 2)
                    }
            
            # Check for anomalies in numeric data
            numeric_cols = df.select_dtypes(include=[np.number]).columns
            for col in numeric_cols:
                if col in df.columns and not df[col].isnull().all():
                    values = df[col].dropna()
                    if len(values) > 0:
                        q1 = values.quantile(0.25)
                        q3 = values.quantile(0.75)
                        iqr = q3 - q1
                        lower_bound = q1 - 1.5 * iqr
                        upper_bound = q3 + 1.5 * iqr
                        
                        outliers = values[(values < lower_bound) | (values > upper_bound)]
                        if len(outliers) > 0:
                            issues['data_anomalies'][col] = {
                                'outlier_count': len(outliers),
                                'outlier_values': outliers.tolist()[:10],  # First 10 outliers
                                'expected_range': f"{lower_bound:.1f} - {upper_bound:.1f}"
                            }
            
            # Check date consistency
            if 'Date Opened' in df.columns and 'Date Closed' in df.columns:
                date_issues = df[
                    (df['Date Opened'].notna()) & 
                    (df['Date Closed'].notna()) & 
                    (df['Date Opened'] > df['Date Closed'])
                ]
                if len(date_issues) > 0:
                    issues['format_issues']['invalid_date_sequence'] = {
                        'count': len(date_issues),
                        'description': 'Date Opened is after Date Closed'
                    }
            
            # Summary
            issues['summary'] = {
                'total_rows': total_rows,
                'total_columns': len(df.columns),
                'columns_with_missing_data': len(issues['missing_data']),
                'columns_with_anomalies': len(issues['data_anomalies']),
                'format_issues_found': len(issues['format_issues'])
            }
            
            return issues
            
        except Exception as e:
            logger.error(f"Error validating data quality: {e}")
            return {'error': str(e)}
    
    def calculate_hiring_metrics(self, df: pd.DataFrame) -> Dict[str, Any]:
        """Calculate key hiring metrics"""
        try:
            metrics = {}
            
            # Basic counts
            metrics['total_positions'] = len(df)
            metrics['filled_positions'] = len(df[df['Date Closed'].notna()])
            metrics['open_positions'] = len(df[df['Date Closed'].isna()])
            
            # Time to hire metrics
            if 'Total Days Open' in df.columns:
                time_data = df['Total Days Open'].dropna()
                if len(time_data) > 0:
                    metrics['time_to_hire'] = {
                        'average_days': round(time_data.mean(), 1),
                        'median_days': round(time_data.median(), 1),
                        'min_days': int(time_data.min()),
                        'max_days': int(time_data.max()),
                        'percentile_75': round(time_data.quantile(0.75), 1),
                        'percentile_90': round(time_data.quantile(0.90), 1)
                    }
            
            # Application metrics
            if 'Inbound Applicants' in df.columns:
                app_data = df['Inbound Applicants'].dropna()
                if len(app_data) > 0:
                    metrics['application_metrics'] = {
                        'total_applications': int(app_data.sum()),
                        'average_per_position': round(app_data.mean(), 1),
                        'median_per_position': round(app_data.median(), 1)
                    }
            
            # Source analysis
            if 'Source' in df.columns:
                source_counts = df['Source'].value_counts()
                metrics['source_distribution'] = {
                    'top_sources': source_counts.head(10).to_dict(),
                    'total_sources': len(source_counts)
                }
            
            # Location analysis
            if 'Location' in df.columns:
                location_counts = df['Location'].value_counts()
                metrics['location_distribution'] = {
                    'by_location': location_counts.head(10).to_dict(),
                    'total_locations': len(location_counts)
                }
            
            # Department analysis
            if 'Department' in df.columns:
                dept_counts = df['Department'].value_counts()
                metrics['department_distribution'] = {
                    'by_department': dept_counts.head(10).to_dict(),
                    'total_departments': len(dept_counts)
                }
            
            # Cost analysis
            if 'Agency Fee' in df.columns:
                cost_data = df['Agency Fee'].dropna()
                cost_data = cost_data[cost_data > 0]  # Only positive costs
                if len(cost_data) > 0:
                    metrics['cost_analysis'] = {
                        'total_agency_fees': round(cost_data.sum(), 2),
                        'average_fee': round(cost_data.mean(), 2),
                        'positions_with_fees': len(cost_data),
                        'percentage_with_fees': round((len(cost_data) / len(df)) * 100, 1)
                    }
            
            return metrics
            
        except Exception as e:
            logger.error(f"Error calculating hiring metrics: {e}")
            return {'error': str(e)}
    
    def generate_chart_data(self, df: pd.DataFrame, chart_type: str, breakdown_by: str = None) -> Dict[str, Any]:
        """Generate data formatted for charts and visualizations"""
        try:
            chart_data = {'chart_type': chart_type, 'data': {}}
            
            if chart_type == 'hiring_trends':
                # Monthly hiring trends
                if 'Date Closed' in df.columns:
                    filled_positions = df[df['Date Closed'].notna()].copy()
                    filled_positions['Month'] = filled_positions['Date Closed'].dt.to_period('M')
                    monthly_counts = filled_positions['Month'].value_counts().sort_index()
                    
                    chart_data['data'] = {
                        'labels': [str(month) for month in monthly_counts.index],
                        'values': monthly_counts.values.tolist(),
                        'title': 'Monthly Hiring Trends'
                    }
            
            elif chart_type == 'source_effectiveness':
                # Source distribution pie chart
                if 'Source' in df.columns:
                    source_counts = df['Source'].value_counts().head(10)
                    chart_data['data'] = {
                        'labels': source_counts.index.tolist(),
                        'values': source_counts.values.tolist(),
                        'title': 'Candidate Source Distribution'
                    }
            
            elif chart_type == 'time_to_hire_distribution':
                # Time to hire histogram
                if 'Total Days Open' in df.columns:
                    time_data = df['Total Days Open'].dropna()
                    if len(time_data) > 0:
                        # Create bins
                        bins = [0, 30, 60, 90, 120, 180, float('inf')]
                        labels = ['0-30 days', '31-60 days', '61-90 days', '91-120 days', '121-180 days', '180+ days']
                        
                        hist_data = pd.cut(time_data, bins=bins, labels=labels, right=False)
                        hist_counts = hist_data.value_counts()
                        
                        chart_data['data'] = {
                            'labels': hist_counts.index.tolist(),
                            'values': hist_counts.values.tolist(),
                            'title': 'Time to Hire Distribution'
                        }
            
            elif chart_type == 'department_hiring':
                # Department hiring bar chart
                if 'Department' in df.columns:
                    dept_counts = df['Department'].value_counts().head(10)
                    chart_data['data'] = {
                        'labels': dept_counts.index.tolist(),
                        'values': dept_counts.values.tolist(),
                        'title': 'Hiring by Department'
                    }
            
            elif chart_type == 'conversion_funnel':
                # Hiring funnel analysis
                funnel_data = {}
                if 'Inbound Applicants' in df.columns:
                    funnel_data['Applications'] = df['Inbound Applicants'].sum()
                if 'Recruiter Screen' in df.columns:
                    funnel_data['Recruiter Screens'] = df['Recruiter Screen'].sum()
                if 'Hiring Manager Screen' in df.columns:
                    funnel_data['Hiring Manager Screens'] = df['Hiring Manager Screen'].sum()
                if 'Final Round Interviews' in df.columns:
                    funnel_data['Final Interviews'] = df['Final Round Interviews'].sum()
                if 'Offers' in df.columns:
                    funnel_data['Offers'] = df['Offers'].sum()
                
                chart_data['data'] = {
                    'labels': list(funnel_data.keys()),
                    'values': list(funnel_data.values()),
                    'title': 'Hiring Process Conversion Funnel'
                }
            
            return chart_data
            
        except Exception as e:
            logger.error(f"Error generating chart data: {e}")
            return {'error': str(e)}

# Global instance
hr_analytics = HRAnalytics()
