"""
SharePoint MCP tools using Microsoft Graph API
"""
import base64, os, io
from functools import wraps
from typing import Optional, Dict, Any
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from .common import logger, mcp, ACCESS_TOKEN, SITE_ID, DRIVE_ID, make_graph_request
from .resources import list_folders, list_documents, get_document_content, download_document
from .context_helper import context_helper
from .analytics_helper import hr_analytics

# Helper functions
def _handle_sp_operation(func):
    """Decorator for SharePoint operations with error handling"""
    @wraps(func)
    async def wrapper(*args, **kwargs):
        try:
            return await func(*args, **kwargs)
        except Exception as e:
            logger.error(f"Error in {func.__name__}: {str(e)}")
            return {"success": False, "message": f"Operation failed: {str(e)}"}
    return wrapper

def _file_success_response(file_info: Dict[str, Any], message: str) -> Dict[str, Any]:
    """Standard success response for file operations"""
    return {
        "success": True,
        "message": message,
        "file": {
            "name": file_info.get('name'),
            "id": file_info.get('id'),
            "url": file_info.get('webUrl'),
            "size": file_info.get('size', 0)
        }
    }

def _folder_success_response(folder_info: Dict[str, Any], message: str) -> Dict[str, Any]:
    """Standard success response for folder operations"""
    return {
        "success": True,
        "message": message,
        "folder": {
            "name": folder_info.get('name'),
            "id": folder_info.get('id'),
            "url": folder_info.get('webUrl')
        }
    }

def _create_word_document(content: str) -> bytes:
    """Create a proper Word document from markdown-like content"""
    try:
        # Create a new Document
        doc = Document()
        
        # Split content into lines for processing
        lines = content.split('\n')
        i = 0
        
        while i < len(lines):
            line = lines[i].strip()
            
            if not line:
                # Add empty paragraph for spacing
                doc.add_paragraph()
                i += 1
                continue
            
            # Handle headings
            if line.startswith('# '):
                # Main title (level 0)
                title = doc.add_heading(line[2:], 0)
                title.alignment = WD_ALIGN_PARAGRAPH.CENTER
            elif line.startswith('## '):
                # Section heading (level 1)
                doc.add_heading(line[3:], 1)
            elif line.startswith('### '):
                # Subsection heading (level 2)
                doc.add_heading(line[4:], 2)
            elif line.startswith('---'):
                # Horizontal rule - add spacing
                doc.add_paragraph()
            elif line.startswith('**') and line.endswith('**'):
                # Bold text as separate paragraph
                para = doc.add_paragraph()
                para.add_run(line[2:-2]).bold = True
            elif '**' in line:
                # Mixed formatting in paragraph
                para = doc.add_paragraph()
                parts = line.split('**')
                for j, part in enumerate(parts):
                    if j % 2 == 0:
                        para.add_run(part)
                    else:
                        para.add_run(part).bold = True
            else:
                # Regular paragraph
                doc.add_paragraph(line)
            
            i += 1
        
        # Save to bytes
        doc_buffer = io.BytesIO()
        doc.save(doc_buffer)
        doc_buffer.seek(0)
        return doc_buffer.getvalue()
        
    except Exception as e:
        logger.error(f"Error creating Word document: {e}")
        # Fallback: create simple document with plain text
        doc = Document()
        doc.add_paragraph(content)
        doc_buffer = io.BytesIO()
        doc.save(doc_buffer)
        doc_buffer.seek(0)
        return doc_buffer.getvalue()

def _generate_recommendations(validation_results: dict, metrics: dict) -> list:
    """Generate actionable recommendations based on analysis results"""
    recommendations = []
    
    try:
        # Data quality recommendations
        if validation_results.get('summary', {}).get('columns_with_missing_data', 0) > 0:
            recommendations.append({
                "type": "data_quality",
                "priority": "medium",
                "message": f"Found missing data in {validation_results['summary']['columns_with_missing_data']} columns. Consider data cleanup.",
                "action": "Review and fill missing values for better analysis accuracy"
            })
        
        # Performance recommendations
        if 'time_to_hire' in metrics:
            avg_days = metrics['time_to_hire'].get('average_days', 0)
            if avg_days > 90:
                recommendations.append({
                    "type": "performance",
                    "priority": "high",
                    "message": f"Average time to hire is {avg_days} days, which may be too long",
                    "action": "Consider streamlining the hiring process to reduce time to hire"
                })
        
        # Cost recommendations
        if 'cost_analysis' in metrics:
            pct_with_fees = metrics['cost_analysis'].get('percentage_with_fees', 0)
            if pct_with_fees > 30:
                recommendations.append({
                    "type": "cost",
                    "priority": "medium",
                    "message": f"{pct_with_fees}% of hires involved agency fees",
                    "action": "Consider increasing direct sourcing to reduce agency costs"
                })
        
        # Source diversity recommendations
        if 'source_distribution' in metrics:
            top_sources = metrics['source_distribution'].get('top_sources', {})
            if len(top_sources) > 0:
                top_source_pct = max(top_sources.values()) / sum(top_sources.values()) * 100
                if top_source_pct > 60:
                    recommendations.append({
                        "type": "sourcing",
                        "priority": "medium",
                        "message": f"Over-reliance on single source ({top_source_pct:.1f}% from top source)",
                        "action": "Diversify candidate sourcing channels for better pipeline resilience"
                    })
        
        return recommendations
        
    except Exception as e:
        logger.error(f"Error generating recommendations: {e}")
        return []


# Basic tool implementations
@mcp.tool(name="List_SharePoint_Folders", description="List folders in the specified SharePoint directory or root if not specified")
async def list_folders_tool(parent_folder: Optional[str] = None):
    """List folders in the specified SharePoint directory or root if not specified"""
    return list_folders(parent_folder)

@mcp.tool(name="List_SharePoint_Documents", description="List all documents in a specified SharePoint folder")
async def list_documents_tool(folder_name: str):
    """List all documents in a specified SharePoint folder"""
    return list_documents(folder_name)

@mcp.tool(name="Get_Document_Content", description="Get content of a document in SharePoint")
async def get_document_content_tool(folder_name: str, file_name: str):
    """Get content of a document in SharePoint"""
    return get_document_content(folder_name, file_name)

@mcp.tool(name="Download_Document", description="Download a document from SharePoint to local filesystem")
@_handle_sp_operation
async def download_document_tool(folder_name: str, file_name: str, local_path: str):
    """Download a document from SharePoint to local filesystem with fallback support"""
    return download_document(folder_name, file_name, local_path)

# Graph API specific tools
@mcp.tool(name="Create_Folder", description="Create a new folder in the specified directory or root if not specified")
@_handle_sp_operation
async def create_folder(folder_name: str, parent_folder: Optional[str] = None):
    """Create a new folder using Graph API"""
    logger.info(f"Creating folder '{folder_name}' in {parent_folder or 'root directory'}")
    
    # Check for existing folder
    existing_folders = list_folders(parent_folder)
    if any(f["name"] == folder_name for f in existing_folders):
        return {"success": False, "message": f"Folder {folder_name} already exists"}
    
    try:
        if not parent_folder or parent_folder == "":
            endpoint = f"sites/{SITE_ID}/drives/{DRIVE_ID}/root/children"
        else:
            endpoint = f"sites/{SITE_ID}/drives/{DRIVE_ID}/root:/{parent_folder}:/children"
        
        folder_data = {
            "name": folder_name,
            "folder": {},
            "@microsoft.graph.conflictBehavior": "rename"
        }
        
        response = make_graph_request("POST", endpoint, folder_data)
        
        if response and response.status_code == 201:
            folder_info = response.json()
            return _folder_success_response(folder_info, f"Folder {folder_name} created successfully")
        else:
            return {"success": False, "message": f"Failed to create folder: {response.status_code if response else 'No response'}"}
            
    except Exception as e:
        return {"success": False, "message": f"Error creating folder: {str(e)}"}

@mcp.tool(name="Upload_Document", description="Upload a new file to a SharePoint directory")
@_handle_sp_operation
async def upload_document(folder_name: str, file_name: str, content: str, is_base64: bool = False):
    """Upload a new file using Graph API with special handling for Word documents"""
    logger.info(f"Uploading document {file_name} to folder {folder_name}")
    
    try:
        # Special handling for Word documents
        if file_name.lower().endswith('.docx'):
            file_content = _create_word_document(content)
            content_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        else:
            # Convert content for other file types
            file_content = base64.b64decode(content) if is_base64 else content.encode('utf-8')
            content_type = "text/plain" if file_name.lower().endswith('.txt') else "application/octet-stream"
        
        # Build endpoint
        if not folder_name or folder_name == "":
            endpoint = f"sites/{SITE_ID}/drives/{DRIVE_ID}/root:/{file_name}:/content"
        else:
            endpoint = f"sites/{SITE_ID}/drives/{DRIVE_ID}/root:/{folder_name}/{file_name}:/content"
        
        response = make_graph_request("PUT", endpoint, file_content, content_type)
        
        if response and response.status_code in [200, 201]:
            file_info = response.json()
            return _file_success_response(file_info, f"File {file_name} uploaded successfully")
        else:
            return {"success": False, "message": f"Failed to upload file: {response.status_code if response else 'No response'}"}
            
    except Exception as e:
        return {"success": False, "message": f"Error uploading file: {str(e)}"}

@mcp.tool(name="Delete_Document", description="Delete a document from a SharePoint directory")
@_handle_sp_operation
async def delete_document(folder_name: str, file_name: str):
    """Delete a document using Graph API"""
    logger.info(f"Deleting document {file_name} from folder {folder_name}")
    
    try:
        # Build endpoint
        if not folder_name or folder_name == "":
            endpoint = f"sites/{SITE_ID}/drives/{DRIVE_ID}/root:/{file_name}"
        else:
            endpoint = f"sites/{SITE_ID}/drives/{DRIVE_ID}/root:/{folder_name}/{file_name}"
        
        # Check if file exists
        check_response = make_graph_request("GET", endpoint)
        if not check_response or check_response.status_code != 200:
            return {"success": False, "message": f"File {file_name} does not exist"}
        
        # Delete the file
        response = make_graph_request("DELETE", endpoint)
        
        if response and response.status_code == 204:
            return {"success": True, "message": f"File {file_name} deleted successfully"}
        else:
            return {"success": False, "message": f"Failed to delete file: {response.status_code if response else 'No response'}"}
            
    except Exception as e:
        return {"success": False, "message": f"Error deleting file: {str(e)}"}

# Context and Column Definition Tools
@mcp.tool(name="Get_Column_Definition", description="Get definition for a specific Excel column")
async def get_column_definition_tool(column_name: str):
    """Get definition for a specific Excel column"""
    try:
        definition = context_helper.get_column_definition(column_name)
        if definition:
            return {
                "success": True,
                "column": column_name,
                "definition": definition
            }
        else:
            return {
                "success": False,
                "message": f"No definition found for column '{column_name}'"
            }
    except Exception as e:
        return {"success": False, "message": f"Error getting column definition: {str(e)}"}

@mcp.tool(name="Search_Column_Definitions", description="Search for column definitions containing a specific term")
async def search_column_definitions_tool(search_term: str):
    """Search for column definitions containing a specific term"""
    try:
        results = context_helper.search_definitions(search_term)
        return {
            "success": True,
            "search_term": search_term,
            "result_count": len(results),
            "results": results
        }
    except Exception as e:
        return {"success": False, "message": f"Error searching column definitions: {str(e)}"}

@mcp.tool(name="Get_All_Column_Definitions", description="Get all available column definitions")
async def get_all_column_definitions_tool():
    """Get all available column definitions"""
    try:
        definitions = context_helper.get_all_definitions()
        return {
            "success": True,
            "total_columns": len(definitions),
            "definitions": definitions
        }
    except Exception as e:
        return {"success": False, "message": f"Error getting all definitions: {str(e)}"}

@mcp.tool(name="Get_Matching_Columns", description="Get definitions for columns that match a list of column names")
async def get_matching_columns_tool(column_names: list):
    """Get definitions for columns that exist in the dictionary"""
    try:
        matching = context_helper.get_matching_columns(column_names)
        return {
            "success": True,
            "input_columns": column_names,
            "matched_count": len(matching),
            "matching_definitions": matching
        }
    except Exception as e:
        return {"success": False, "message": f"Error getting matching columns: {str(e)}"}

# HR Analytics and Data Quality Tools
@mcp.tool(name="Validate_Excel_Data_Quality", description="Validate data quality and identify issues in Excel HR data")
async def validate_excel_data_quality_tool(folder_name: str, file_name: str):
    """Validate data quality and identify issues in Excel HR data"""
    try:
        # Get the Excel content
        content_result = get_document_content(folder_name, file_name)
        if not content_result.get("success", True):
            return {"success": False, "message": "Failed to retrieve Excel file"}
        
        # Parse content into DataFrame
        df = hr_analytics.parse_excel_content(content_result.get("content", ""))
        if df.empty:
            return {"success": False, "message": "No data found in Excel file"}
        
        # Validate data quality
        validation_results = hr_analytics.validate_data_quality(df)
        
        return {
            "success": True,
            "file_name": file_name,
            "validation_results": validation_results
        }
        
    except Exception as e:
        return {"success": False, "message": f"Error validating data quality: {str(e)}"}

@mcp.tool(name="Calculate_HR_Metrics", description="Calculate comprehensive HR and recruiting metrics from Excel data")
async def calculate_hr_metrics_tool(folder_name: str, file_name: str):
    """Calculate comprehensive HR and recruiting metrics from Excel data"""
    try:
        # Get the Excel content
        content_result = get_document_content(folder_name, file_name)
        if not content_result.get("success", True):
            return {"success": False, "message": "Failed to retrieve Excel file"}
        
        # Parse content into DataFrame
        df = hr_analytics.parse_excel_content(content_result.get("content", ""))
        if df.empty:
            return {"success": False, "message": "No data found in Excel file"}
        
        # Calculate metrics
        metrics = hr_analytics.calculate_hiring_metrics(df)
        
        return {
            "success": True,
            "file_name": file_name,
            "metrics": metrics,
            "data_summary": {
                "total_rows": len(df),
                "total_columns": len(df.columns),
                "columns": list(df.columns)
            }
        }
        
    except Exception as e:
        return {"success": False, "message": f"Error calculating HR metrics: {str(e)}"}

@mcp.tool(name="Generate_Chart_Data", description="Generate data formatted for charts and visualizations")
async def generate_chart_data_tool(folder_name: str, file_name: str, chart_type: str, breakdown_by: Optional[str] = None):
    """Generate data formatted for charts and visualizations
    
    Chart types: hiring_trends, source_effectiveness, time_to_hire_distribution, 
                department_hiring, conversion_funnel
    """
    try:
        # Get the Excel content
        content_result = get_document_content(folder_name, file_name)
        if not content_result.get("success", True):
            return {"success": False, "message": "Failed to retrieve Excel file"}
        
        # Parse content into DataFrame
        df = hr_analytics.parse_excel_content(content_result.get("content", ""))
        if df.empty:
            return {"success": False, "message": "No data found in Excel file"}
        
        # Generate chart data
        chart_data = hr_analytics.generate_chart_data(df, chart_type, breakdown_by)
        
        return {
            "success": True,
            "file_name": file_name,
            "chart_type": chart_type,
            "chart_data": chart_data
        }
        
    except Exception as e:
        return {"success": False, "message": f"Error generating chart data: {str(e)}"}

@mcp.tool(name="Analyze_HR_File_Complete", description="Complete analysis of HR Excel file including data quality, metrics, and chart data")
async def analyze_hr_file_complete_tool(folder_name: str, file_name: str):
    """Perform complete analysis of HR Excel file including data quality, metrics, and suggested visualizations"""
    try:
        # Get the Excel content
        content_result = get_document_content(folder_name, file_name)
        if not content_result.get("success", True):
            return {"success": False, "message": "Failed to retrieve Excel file"}
        
        # Parse content into DataFrame
        df = hr_analytics.parse_excel_content(content_result.get("content", ""))
        if df.empty:
            return {"success": False, "message": "No data found in Excel file"}
        
        # Perform all analyses
        validation_results = hr_analytics.validate_data_quality(df)
        metrics = hr_analytics.calculate_hiring_metrics(df)
        
        # Generate multiple chart datasets
        chart_suggestions = []
        chart_types = ['hiring_trends', 'source_effectiveness', 'time_to_hire_distribution', 'department_hiring', 'conversion_funnel']
        
        for chart_type in chart_types:
            try:
                chart_data = hr_analytics.generate_chart_data(df, chart_type)
                if chart_data.get('data') and len(chart_data['data'].get('values', [])) > 0:
                    chart_suggestions.append(chart_data)
            except:
                continue  # Skip charts that can't be generated
        
        return {
            "success": True,
            "file_name": file_name,
            "data_summary": {
                "total_rows": len(df),
                "total_columns": len(df.columns),
                "columns": list(df.columns),
                "date_range": {
                    "earliest_date": str(df.select_dtypes(include=['datetime64']).min().min()) if not df.select_dtypes(include=['datetime64']).empty else None,
                    "latest_date": str(df.select_dtypes(include=['datetime64']).max().max()) if not df.select_dtypes(include=['datetime64']).empty else None
                }
            },
            "data_quality": validation_results,
            "hr_metrics": metrics,
            "suggested_charts": chart_suggestions,
            "recommendations": _generate_recommendations(validation_results, metrics)
        }
        
    except Exception as e:
        return {"success": False, "message": f"Error performing complete analysis: {str(e)}"}

