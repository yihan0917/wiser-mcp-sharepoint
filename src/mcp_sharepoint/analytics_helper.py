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
        # Patterns for auto-detection (no hardcoded column names)
        self.date_patterns = ['date', 'created', 'opened', 'closed', 'filled', 'start']
        self.time_patterns = ['time', 'days', 'duration']
        self.role_patterns = ['title', 'role', 'position', 'job']
        self.location_patterns = ['location', 'country', 'city', 'state', 'region']
        self.person_patterns = ['recruiter', 'manager', 'hiring', 'executive']
        self.department_patterns = ['department', 'org', 'team', 'division']
    
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
        """Clean and standardize DataFrame with intelligent type detection"""
        try:
            # Auto-detect and convert date columns
            for col in df.columns:
                col_lower = col.lower()
                if any(pattern in col_lower for pattern in self.date_patterns):
                    df[col] = pd.to_datetime(df[col], errors='coerce')
            
            # Auto-detect and convert numeric columns
            for col in df.columns:
                # Try to convert to numeric if it looks like numbers
                if df[col].dtype == 'object':
                    try:
                        converted = pd.to_numeric(df[col], errors='coerce')
                        # If more than 50% successfully converted, treat as numeric
                        if converted.notna().sum() / len(df) > 0.5:
                            df[col] = converted
                    except:
                        pass
            
            # Clean text columns
            text_columns = df.select_dtypes(include=['object']).columns
            for col in text_columns:
                df[col] = df[col].astype(str).str.strip()
                df[col] = df[col].replace('', np.nan)
                df[col] = df[col].replace('nan', np.nan)
            
            return df
            
        except Exception as e:
            logger.error(f"Error cleaning DataFrame: {e}")
            return df
    
    def detect_column_types(self, df: pd.DataFrame) -> Dict[str, List[str]]:
        """Intelligently detect column types based on content and names"""
        column_types = {
            'date_columns': [],
            'time_step_columns': [],
            'numeric_columns': [],
            'role_columns': [],
            'location_columns': [],
            'person_columns': [],
            'department_columns': [],
            'categorical_columns': [],
            'id_columns': []
        }
        
        for col in df.columns:
            col_lower = col.lower()
            
            # Date columns
            if df[col].dtype in ['datetime64[ns]', 'datetime64']:
                column_types['date_columns'].append(col)
            elif any(pattern in col_lower for pattern in self.date_patterns):
                column_types['date_columns'].append(col)
            
            # Time/duration columns (numeric)
            elif 'time' in col_lower and df[col].dtype in ['int64', 'float64']:
                column_types['time_step_columns'].append(col)
            
            # Role/title columns
            elif any(pattern in col_lower for pattern in self.role_patterns):
                column_types['role_columns'].append(col)
            
            # Location columns
            elif any(pattern in col_lower for pattern in self.location_patterns):
                column_types['location_columns'].append(col)
            
            # Person columns (recruiter, manager, etc.)
            elif any(pattern in col_lower for pattern in self.person_patterns):
                column_types['person_columns'].append(col)
            
            # Department columns
            elif any(pattern in col_lower for pattern in self.department_patterns):
                column_types['department_columns'].append(col)
            
            # ID columns
            elif 'id' in col_lower or 'ref' in col_lower:
                column_types['id_columns'].append(col)
            
            # Numeric columns (not already categorized)
            elif df[col].dtype in ['int64', 'float64']:
                column_types['numeric_columns'].append(col)
            
            # Categorical columns (text with limited unique values)
            elif df[col].dtype == 'object':
                unique_ratio = df[col].nunique() / len(df)
                if unique_ratio < 0.5:  # Less than 50% unique values
                    column_types['categorical_columns'].append(col)
        
        return column_types
    
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
        """Calculate key hiring metrics dynamically based on available columns"""
        try:
            metrics = {}
            column_types = self.detect_column_types(df)
            
            # Basic counts
            metrics['total_records'] = len(df)
            
            # Time-based metrics (from time_step_columns or numeric columns with 'time'/'days')
            time_columns = column_types['time_step_columns']
            if time_columns:
                # Calculate total time across all time steps
                df['_total_time'] = df[time_columns].sum(axis=1, skipna=True)
                time_data = df['_total_time'].dropna()
                
                if len(time_data) > 0:
                    metrics['time_metrics'] = {
                        'average_total_time': round(time_data.mean(), 1),
                        'median_total_time': round(time_data.median(), 1),
                        'min_time': round(time_data.min(), 1),
                        'max_time': round(time_data.max(), 1),
                        'percentile_75': round(time_data.quantile(0.75), 1),
                        'percentile_90': round(time_data.quantile(0.90), 1)
                    }
                    
                    # Time by step
                    metrics['time_by_step'] = {}
                    for col in time_columns[:10]:  # Top 10 time steps
                        step_data = df[col].dropna()
                        if len(step_data) > 0:
                            metrics['time_by_step'][col] = {
                                'average': round(step_data.mean(), 1),
                                'median': round(step_data.median(), 1)
                            }
            
            # Role/Position analysis
            role_columns = column_types['role_columns']
            if role_columns:
                for col in role_columns[:3]:  # Top 3 role columns
                    role_counts = df[col].value_counts().head(15)
                    metrics[f'{col}_distribution'] = role_counts.to_dict()
            
            # Location analysis
            location_columns = column_types['location_columns']
            if location_columns:
                for col in location_columns[:3]:  # Top 3 location columns
                    loc_counts = df[col].value_counts().head(15)
                    metrics[f'{col}_distribution'] = loc_counts.to_dict()
            
            # Department analysis
            dept_columns = column_types['department_columns']
            if dept_columns:
                for col in dept_columns[:2]:  # Top 2 department columns
                    dept_counts = df[col].value_counts().head(10)
                    metrics[f'{col}_distribution'] = dept_counts.to_dict()
            
            # Date-based analysis
            date_columns = column_types['date_columns']
            if len(date_columns) >= 2:
                # Try to find creation and completion dates
                creation_col = next((col for col in date_columns if 'creat' in col.lower() or 'open' in col.lower()), None)
                completion_col = next((col for col in date_columns if 'close' in col.lower() or 'fill' in col.lower()), None)
                
                if creation_col and completion_col:
                    filled = df[df[completion_col].notna()]
                    metrics['status_metrics'] = {
                        'total_positions': len(df),
                        'filled_positions': len(filled),
                        'fill_rate': round((len(filled) / len(df)) * 100, 1) if len(df) > 0 else 0
                    }
            
            # Numeric columns analysis (excluding time columns)
            other_numeric = [col for col in column_types['numeric_columns'] if col not in time_columns]
            if other_numeric:
                metrics['numeric_summaries'] = {}
                for col in other_numeric[:5]:  # Top 5 numeric columns
                    num_data = df[col].dropna()
                    if len(num_data) > 0:
                        metrics['numeric_summaries'][col] = {
                            'sum': round(num_data.sum(), 1),
                            'average': round(num_data.mean(), 1),
                            'median': round(num_data.median(), 1)
                        }
            
            return metrics
            
        except Exception as e:
            logger.error(f"Error calculating hiring metrics: {e}")
            return {'error': str(e)}
    
    def generate_chart_data(self, df: pd.DataFrame, chart_type: str, breakdown_by: str = None) -> Dict[str, Any]:
        """Generate data formatted for charts and visualizations - dynamically detects available data"""
        try:
            chart_data = {'chart_type': chart_type, 'data': {}}
            column_types = self.detect_column_types(df)
            
            if chart_type == 'time_by_step':
                # Time spent in each recruiting step
                time_columns = column_types['time_step_columns']
                if time_columns:
                    step_averages = {}
                    for col in time_columns[:10]:  # Top 10 steps
                        avg_time = df[col].mean()
                        if not pd.isna(avg_time):
                            # Shorten column names for display
                            display_name = col.replace('Time in Application Status: ', '').replace('Time In Application State: ', '')
                            step_averages[display_name] = round(avg_time, 1)
                    
                    if step_averages:
                        # Sort by value descending
                        sorted_steps = dict(sorted(step_averages.items(), key=lambda x: x[1], reverse=True))
                        chart_data['data'] = {
                            'labels': list(sorted_steps.keys()),
                            'values': list(sorted_steps.values()),
                            'title': 'Average Time by Recruiting Step (Days)'
                        }
            
            elif chart_type == 'role_distribution':
                # Distribution by role/job title
                role_columns = column_types['role_columns']
                if role_columns:
                    col = role_columns[0]  # Use first role column
                    role_counts = df[col].value_counts().head(10)
                    chart_data['data'] = {
                        'labels': role_counts.index.tolist(),
                        'values': role_counts.values.tolist(),
                        'title': f'Top 10 Roles by Volume'
                    }
            
            elif chart_type == 'location_distribution':
                # Distribution by location
                location_columns = column_types['location_columns']
                if location_columns:
                    # Prefer country over city/state
                    col = next((c for c in location_columns if 'country' in c.lower()), location_columns[0])
                    loc_counts = df[col].value_counts().head(10)
                    chart_data['data'] = {
                        'labels': loc_counts.index.tolist(),
                        'values': loc_counts.values.tolist(),
                        'title': f'Hiring by Location'
                    }
            
            elif chart_type == 'department_distribution':
                # Distribution by department
                dept_columns = column_types['department_columns']
                if dept_columns:
                    col = dept_columns[0]
                    dept_counts = df[col].value_counts().head(10)
                    chart_data['data'] = {
                        'labels': dept_counts.index.tolist(),
                        'values': dept_counts.values.tolist(),
                        'title': 'Hiring by Department'
                    }
            
            elif chart_type == 'time_distribution':
                # Total time distribution (histogram)
                time_columns = column_types['time_step_columns']
                if time_columns:
                    df['_total_time'] = df[time_columns].sum(axis=1, skipna=True)
                    time_data = df['_total_time'].dropna()
                    
                    if len(time_data) > 0:
                        # Create bins
                        bins = [0, 30, 60, 90, 120, 180, float('inf')]
                        labels = ['0-30 days', '31-60 days', '61-90 days', '91-120 days', '121-180 days', '180+ days']
                        
                        hist_data = pd.cut(time_data, bins=bins, labels=labels, right=False)
                        hist_counts = hist_data.value_counts().sort_index()
                        
                        chart_data['data'] = {
                            'labels': hist_counts.index.tolist(),
                            'values': hist_counts.values.tolist(),
                            'title': 'Total Time-to-Hire Distribution'
                        }
            
            elif chart_type == 'monthly_trends':
                # Monthly trends from date columns
                date_columns = column_types['date_columns']
                if date_columns:
                    # Use first date column that looks like a completion date
                    date_col = next((col for col in date_columns if 'fill' in col.lower() or 'close' in col.lower()), date_columns[0])
                    
                    filled_positions = df[df[date_col].notna()].copy()
                    if len(filled_positions) > 0:
                        filled_positions['Month'] = pd.to_datetime(filled_positions[date_col]).dt.to_period('M')
                        monthly_counts = filled_positions['Month'].value_counts().sort_index()
                        
                        chart_data['data'] = {
                            'labels': [str(month) for month in monthly_counts.index],
                            'values': monthly_counts.values.tolist(),
                            'title': 'Monthly Hiring Trends'
                        }
            
            return chart_data
            
        except Exception as e:
            logger.error(f"Error generating chart data: {e}")
            return {'error': str(e)}

# Global instance
hr_analytics = HRAnalytics()
