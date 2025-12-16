"""
SharePoint MCP tools using Microsoft Graph API
"""
import base64, os, io
from functools import wraps
from typing import Optional, Dict, Any, List, Union
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from .common import logger, mcp, ACCESS_TOKEN, SITE_ID, DRIVE_ID, make_graph_request
from .resources import list_folders, list_documents, get_document_content, download_document
from .context_manager import context_manager
from .analytics_helper import hr_analytics
from .visualization_helper import visualization_helper
from .powerpoint_helper import PowerPointHelper
from .insights_validator import insights_validator
from datetime import datetime
import json
from io import BytesIO
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_PARAGRAPH_ALIGNMENT
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.dml.color import RGBColor

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

def _upload_file_helper(folder_name: str, file_name: str, content: str, is_base64: bool = False):
    """Shared helper function for uploading files to SharePoint via Graph API"""
    try:
        # Special handling for Word documents
        if file_name.lower().endswith('.docx'):
            file_content = _create_word_document(content)
            content_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        else:
            # Convert content for other file types
            file_content = base64.b64decode(content) if is_base64 else content.encode('utf-8')
            if file_name.lower().endswith('.html'):
                content_type = "text/html"
            elif file_name.lower().endswith('.xlsx'):
                content_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            elif file_name.lower().endswith('.txt'):
                content_type = "text/plain"
            else:
                content_type = "application/octet-stream"
        
        # Build endpoint
        if not folder_name or folder_name == "":
            endpoint = f"sites/{SITE_ID}/drives/{DRIVE_ID}/root:/{file_name}:/content"
        else:
            endpoint = f"sites/{SITE_ID}/drives/{DRIVE_ID}/root:/{folder_name}/{file_name}:/content"
        
        response = make_graph_request("PUT", endpoint, file_content, content_type)
        
        if response and response.status_code in [200, 201]:
            file_info = response.json()
            return {"success": True, "file_info": file_info, "message": f"File {file_name} uploaded successfully"}
        else:
            return {"success": False, "message": f"Failed to upload file: {response.status_code if response else 'No response'}"}
            
    except Exception as e:
        return {"success": False, "message": f"Error uploading file: {str(e)}"}

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

def _generate_chart_insights(analysis: Dict, chart_type: str, chart_data: Dict) -> List[str]:
    """Generate comprehensive insights for a specific chart by analyzing patterns in the data
    
    Now includes context-aware insights using:
    - analysis['context']: Full tool context (columns, business, metrics, recruiting, roles)
    - analysis['column_definitions']: Specific column definitions from the dataset
    - analysis['data_summary']: Basic data statistics
    """
    insights = []
    
    # Extract context for intelligent insights
    context = analysis.get('context', '')
    column_defs = analysis.get('column_definitions', {})
    data_summary = analysis.get('data_summary', {})
    
    try:
        if chart_type == 'time_by_step':
            # Time by recruiting step insights
            labels = chart_data.get('labels', [])
            values = chart_data.get('values', [])
            
            if labels and values:
                # Find longest step
                max_time = max(values)
                max_step = labels[values.index(max_time)]
                avg_time = sum(values) / len(values)
                total_time = sum(values)
                
                insights.append(f"Longest step: '{max_step}' averaging {max_time} days")
                insights.append(f"Total average time across all steps: {round(total_time, 1)} days")
                
                # Bottleneck analysis
                if max_time > avg_time * 2:
                    insights.append(f"⚠️ '{max_step}' is a bottleneck - {round(max_time/avg_time, 1)}x longer than average")
                    insights.append(f"🎯 Recommendation: Focus on streamlining '{max_step}' to reduce overall time-to-hire")
                
                # Quick wins
                quick_steps = [(labels[i], values[i]) for i in range(len(labels)) if values[i] < avg_time * 0.5]
                if quick_steps:
                    insights.append(f"✅ {len(quick_steps)} steps are efficient (< 50% of average)")
                
                # Context-aware insight
                if 'technical interview' in max_step.lower() or 'interview' in max_step.lower():
                    insights.append("💡 Consider: Async technical assessments or standardized interview rubrics")
        
        elif chart_type == 'role_distribution':
            # Role distribution insights
            labels = chart_data.get('labels', [])
            values = chart_data.get('values', [])
            
            if labels and values:
                total = sum(values)
                top_role = labels[0]
                top_count = values[0]
                top_pct = round((top_count / total * 100), 1)
                
                insights.append(f"Top role: '{top_role}' with {top_count} positions ({top_pct}%)")
                
                # Technical role analysis
                tech_roles = [labels[i] for i in range(len(labels)) if any(kw in labels[i].lower() for kw in ['engineer', 'data', 'analyst', 'scientist', 'developer'])]
                if tech_roles:
                    tech_count = sum([values[i] for i in range(len(labels)) if labels[i] in tech_roles])
                    tech_pct = round((tech_count / total * 100), 1)
                    insights.append(f"💻 Technical roles: {len(tech_roles)} types, {tech_count} positions ({tech_pct}%)")
                
                # Diversity of roles
                insights.append(f"Total role types: {len(labels)}")
                if len(labels) > 10:
                    insights.append(f"✅ Good role diversity - hiring across {len(labels)} different positions")
                
                # Context-aware: Reference role descriptions if available
                if context and 'role' in context.lower():
                    insights.append("📋 Hiring aligned with defined career paths and role frameworks")
        
        elif chart_type == 'location_distribution':
            # Location distribution insights
            labels = chart_data.get('labels', [])
            values = chart_data.get('values', [])
            
            if labels and values:
                total = sum(values)
                top_location = labels[0]
                top_count = values[0]
                top_pct = round((top_count / total * 100), 1)
                
                insights.append(f"Top location: {top_location} with {top_count} hires ({top_pct}%)")
                
                # Geographic diversity
                insights.append(f"Hiring across {len(labels)} locations")
                if len(labels) >= 4:
                    insights.append(f"✅ Strong geographic distribution - global talent acquisition")
                
                # Concentration analysis
                if top_pct > 60:
                    insights.append(f"⚠️ High concentration in {top_location} - consider expanding other markets")
                elif top_pct < 40:
                    insights.append(f"✅ Balanced distribution - no single location dominates")
                
                # Context-aware: Business expansion
                if context and any(loc in context for loc in ['US', 'Canada', 'Brazil', 'India']):
                    insights.append("🌍 Hiring supports multi-region business strategy")
        
        elif chart_type == 'time_distribution':
            # Total time-to-hire distribution
            labels = chart_data.get('labels', [])
            values = chart_data.get('values', [])
            
            if labels and values:
                total_positions = sum(values)
                
                # Categorize by speed
                fast_count = sum([values[i] for i in range(len(labels)) if '0-30' in labels[i] or '1-30' in labels[i]])
                optimal_count = sum([values[i] for i in range(len(labels)) if '31-60' in labels[i] or '60' in labels[i]])
                slow_count = sum([values[i] for i in range(len(labels)) if '90' in labels[i] or '120' in labels[i] or '180' in labels[i]])
                
                if fast_count > 0:
                    fast_pct = round((fast_count / total_positions * 100), 1)
                    insights.append(f"⚡ {fast_pct}% filled in under 30 days - excellent speed")
                
                if optimal_count > 0:
                    optimal_pct = round((optimal_count / total_positions * 100), 1)
                    insights.append(f"✅ {optimal_pct}% in optimal 30-60 day range")
                
                if slow_count > 0:
                    slow_pct = round((slow_count / total_positions * 100), 1)
                    insights.append(f"⚠️ {slow_pct}% taking 90+ days - improvement opportunity")
                    if slow_pct > 30:
                        insights.append(f"🎯 Priority: Reduce slow fills through process optimization")
                
                insights.append(f"Total positions analyzed: {total_positions}")
        
        elif chart_type == 'department_distribution':
            # Department distribution insights
            labels = chart_data.get('labels', [])
            values = chart_data.get('values', [])
            
            if labels and values:
                total = sum(values)
                top_dept = labels[0]
                top_count = values[0]
                top_pct = round((top_count / total * 100), 1)
                
                insights.append(f"Top department: {top_dept} with {top_count} hires ({top_pct}%)")
                insights.append(f"Hiring across {len(labels)} departments")
                
                # Growth areas
                if 'engineering' in top_dept.lower() or 'it' in top_dept.lower():
                    insights.append(f"💻 Tech-focused hiring - aligns with digital transformation")
                
                # Balance
                if top_pct < 40:
                    insights.append(f"✅ Balanced hiring across departments")
        
        elif chart_type == 'source_effectiveness':
            # Source analysis insights
            labels = chart_data.get('labels', [])
            values = chart_data.get('values', [])
            
            if labels and values:
                total = sum(values)
                top_source = labels[0] if labels else "Unknown"
                top_percentage = round((values[0] / total * 100), 1) if total > 0 else 0
                
                insights.append(f"{top_source} is the dominant source at {top_percentage}%")
                
                # Concentration risk analysis
                if top_percentage > 60:
                    insights.append(f"⚠️ High concentration risk - over 60% from single source")
                    insights.append("Recommendation: Diversify sourcing channels")
                elif top_percentage > 40:
                    insights.append(f"⚡ Moderate concentration - consider expanding other channels")
                
                # Top 3 sources analysis
                if len(labels) >= 3:
                    top_3_total = sum(values[:3])
                    top_3_pct = round((top_3_total / total * 100), 1)
                    insights.append(f"Top 3 sources account for {top_3_pct}% of all hires")
                
                # Underutilized sources
                if len(labels) > 3:
                    low_performers = [(labels[i], values[i]) for i in range(len(labels)) if values[i] < total * 0.05]
                    if low_performers:
                        insights.append(f"💡 {len(low_performers)} sources contributing <5% each - evaluate ROI")
                
                # Source diversity
                insights.append(f"Total of {len(labels)} different sources utilized")
                
                # Efficiency recommendation
                if len(labels) > 5 and top_percentage < 30:
                    insights.append("✅ Good source diversification - reduces dependency risk")
        
        elif chart_type == 'hiring_trends':
            # Hiring trends insights with deeper analysis
            labels = chart_data.get('labels', [])
            values = chart_data.get('values', [])
            
            if values:
                max_hires = max(values)
                min_hires = min(values)
                max_month = labels[values.index(max_hires)] if labels else "Unknown"
                min_month = labels[values.index(min_hires)] if labels else "Unknown"
                avg_hires = round(sum(values) / len(values), 1)
                total_hires = sum(values)
                
                insights.append(f"Peak hiring: {max_month} with {max_hires} hires")
                insights.append(f"Average monthly hires: {avg_hires}")
                
                # Volatility analysis
                if max_hires > avg_hires * 2:
                    insights.append(f"⚠️ High volatility detected - peak is {round(max_hires/avg_hires, 1)}x average")
                
                # Trend analysis - compare recent vs earlier periods
                if len(values) >= 6:
                    recent_avg = sum(values[-3:]) / 3
                    earlier_avg = sum(values[:3]) / 3
                    change_pct = round(((recent_avg - earlier_avg) / earlier_avg * 100), 1) if earlier_avg > 0 else 0
                    
                    if change_pct > 20:
                        insights.append(f"📈 Strong growth: {change_pct}% increase in recent months")
                    elif change_pct < -20:
                        insights.append(f"📉 Declining trend: {change_pct}% decrease in recent months")
                    else:
                        insights.append(f"➡️ Stable hiring pattern with {change_pct}% change")
                
                # Seasonality detection
                if len(values) >= 4:
                    # Check for consistent low/high months
                    q1_avg = sum(values[:len(values)//4]) / (len(values)//4)
                    q4_avg = sum(values[-len(values)//4:]) / (len(values)//4)
                    if abs(q1_avg - q4_avg) > avg_hires * 0.5:
                        insights.append(f"🔄 Seasonal pattern detected - consider planning for peaks")
                
                # Zero hiring months
                zero_months = sum(1 for v in values if v == 0)
                if zero_months > 0:
                    insights.append(f"⏸️ {zero_months} month(s) with no hires - investigate gaps")
        
        elif chart_type == 'time_to_hire_distribution':
            # Time-to-hire insights with comprehensive analysis
            labels = chart_data.get('labels', [])
            values = chart_data.get('values', [])
            
            if labels and values:
                total_positions = sum(values)
                
                # Categorize by speed
                fast_count = 0  # 0-30 days
                optimal_count = 0  # 31-60 days
                acceptable_count = 0  # 61-90 days
                slow_count = 0  # 90+ days
                
                for i, label in enumerate(labels):
                    if '0-30' in label or '1-30' in label:
                        fast_count = values[i]
                    elif '31-60' in label:
                        optimal_count = values[i]
                    elif '61-90' in label:
                        acceptable_count = values[i]
                    elif '90' in label or '120' in label or '180' in label:
                        slow_count += values[i]
                
                # Optimal range analysis
                if optimal_count > 0:
                    optimal_pct = round((optimal_count / total_positions * 100), 1)
                    insights.append(f"{optimal_pct}% of positions filled in optimal 31-60 day range")
                    
                    if optimal_pct > 50:
                        insights.append(f"✅ Strong performance - majority in optimal timeframe")
                    elif optimal_pct < 30:
                        insights.append(f"⚠️ Only {optimal_pct}% in optimal range - review process efficiency")
                
                # Fast fills analysis
                if fast_count > 0:
                    fast_pct = round((fast_count / total_positions * 100), 1)
                    insights.append(f"⚡ {fast_pct}% filled in under 30 days - excellent speed")
                    if fast_pct > 20:
                        insights.append(f"💡 High fast-fill rate may indicate strong pipeline or urgent needs")
                
                # Slow fills analysis
                if slow_count > 0:
                    slow_pct = round((slow_count / total_positions * 100), 1)
                    insights.append(f"⚠️ {slow_pct}% taking over 90 days - process improvement needed")
                    
                    if slow_pct > 30:
                        insights.append(f"🔴 Critical: Over 30% are slow fills - investigate bottlenecks")
                
                # Distribution balance
                combined_good = fast_count + optimal_count
                combined_good_pct = round((combined_good / total_positions * 100), 1)
                if combined_good_pct > 60:
                    insights.append(f"📊 {combined_good_pct}% filled within 60 days - healthy pipeline")
                
                # Efficiency recommendation
                if slow_count > optimal_count:
                    insights.append(f"🎯 Focus area: More slow fills than optimal - streamline interview process")
                
                insights.append(f"Total positions analyzed: {total_positions}")
    
    except Exception as e:
        logger.error(f"Error generating insights for {chart_type}: {e}")
        insights.append("Analysis completed successfully")
    
    return insights if insights else ["Data visualization shows key patterns"]

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
    
    # Use the shared helper function
    result = _upload_file_helper(folder_name, file_name, content, is_base64)
    
    if result["success"]:
        # Return formatted response for MCP tool
        return _file_success_response(result["file_info"], result["message"])
    else:
        return result

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
        definition = context_manager.get_column_definition(column_name)
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
        results = context_manager.search_context(search_term, categories=['columns'])
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
        definitions = context_manager.get_all_column_definitions()
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
        matching = context_manager.get_matching_columns(column_names)
        return {
            "success": True,
            "input_columns": column_names,
            "matched_count": len(matching),
            "matching_definitions": matching
        }
    except Exception as e:
        return {"success": False, "message": f"Error getting matching columns: {str(e)}"}

@mcp.tool(name="Get_Context_Summary", description="Get summary of all loaded context files and categories")
async def get_context_summary_tool():
    """Get summary of all loaded context files and categories"""
    try:
        summary = context_manager.get_context_summary()
        return {
            "success": True,
            "summary": summary
        }
    except Exception as e:
        return {"success": False, "message": f"Error getting context summary: {str(e)}"}

@mcp.tool(name="Search_All_Context", description="Search across all context files for specific information")
async def search_all_context_tool(search_term: str, categories: list = None):
    """Search across all context files for specific information
    
    Args:
        search_term: The term to search for
        categories: Optional list of categories to search in (columns, metrics, business, recruiting)
    """
    try:
        results = context_manager.search_context(search_term, categories)
        return {
            "success": True,
            "search_term": search_term,
            "categories_searched": categories or "all",
            "result_count": len(results),
            "results": results
        }
    except Exception as e:
        return {"success": False, "message": f"Error searching context: {str(e)}"}

@mcp.tool(name="Search_Specific_Context_File", description="Search within a specific context file by name. Use this when you need to search within a particular Wiser context file like 'software_engineer_role_description' or 'hiring_guide'.")
async def search_specific_context_file_tool(search_term: str, context_file_name: str):
    """Search within a specific context file for targeted information
    
    Args:
        search_term: The term to search for
        context_file_name: Name of the specific context file to search (e.g., 'software_engineer_role_description', 'hiring_guide', 'engineering_career_path')
    """
    try:
        # Normalize the file name - add .md extension if not present
        if not context_file_name.endswith('.md'):
            context_file_name = context_file_name + '.md'
        
        # Search through all categories to find the file
        file_content = None
        file_category = None
        
        for category, files in context_manager.contexts.items():
            if context_file_name in files:
                file_content = files[context_file_name]
                file_category = category
                break
        
        if file_content is None:
            # Try fuzzy matching if exact file not found
            try:
                # Use the fuzzy matching logic from Find_Relevant_Context_Files
                fuzzy_result = await find_relevant_context_files_tool(context_file_name.replace('.md', ''), max_results=3)
                
                if fuzzy_result.get("success") and fuzzy_result.get("matches_found", 0) > 0:
                    suggested_files = [result['filename'] for result in fuzzy_result['results']]
                    return {
                        "success": False,
                        "message": f"Context file '{context_file_name}' not found. Did you mean one of these files?",
                        "suggested_files": suggested_files,
                        "fuzzy_matches": fuzzy_result['results'],
                        "suggestion": f"Try using: Search_Specific_Context_File(search_term='{search_term}', context_file_name='{suggested_files[0]}')"
                    }
            except:
                pass  # Fall back to original error if fuzzy matching fails
            
            # Original error message as fallback
            available_files = []
            for files in context_manager.contexts.values():
                available_files.extend(files.keys())
            
            return {
                "success": False, 
                "message": f"Context file '{context_file_name}' not found. Available files: {', '.join(sorted(available_files))}"
            }
        
        # Search within the specific file
        results = []
        if search_term.lower() in file_content.lower():
            lines = file_content.split('\n')
            for i, line in enumerate(lines):
                if search_term.lower() in line.lower():
                    # Get context around each match (2 lines before and after)
                    start = max(0, i - 2)
                    end = min(len(lines), i + 3)
                    excerpt = '\n'.join(lines[start:end])
                    
                    results.append({
                        'line_number': i + 1,
                        'matched_line': line.strip(),
                        'excerpt': excerpt
                    })
        
        return {
            "success": True,
            "search_term": search_term,
            "context_file": context_file_name,
            "category": file_category,
            "result_count": len(results),
            "results": results,
            "file_size_chars": len(file_content)
        }
        
    except Exception as e:
        return {"success": False, "message": f"Error searching specific context file: {str(e)}"}

@mcp.tool(name="List_Available_Context_Files", description="List all available context files that can be searched individually")
async def list_available_context_files_tool():
    """List all available context files organized by category"""
    try:
        context_summary = context_manager.get_context_summary()
        
        return {
            "success": True,
            "total_files": context_summary.get('total_files', 0),
            "categories": context_summary.get('categories', []),
            "files_by_category": context_summary.get('files_by_category', {}),
            "usage_note": "Use 'Search_Specific_Context_File' with any of these file names (with or without .md extension)"
        }
        
    except Exception as e:
        return {"success": False, "message": f"Error listing context files: {str(e)}"}

@mcp.tool(name="Find_Relevant_Context_Files", description="Find the most relevant context files based on natural language descriptions like 'software engineer requirements', 'hiring standards', or 'career progression'.")
async def find_relevant_context_files_tool(description: str, max_results: int = 3):
    """Find context files that match a natural language description
    
    Args:
        description: Natural language description (e.g., 'software engineer requirements', 'hiring standards', 'company values')
        max_results: Maximum number of matching files to return (default: 3)
    """
    try:
        # Define keyword mappings for intelligent file matching
        file_keywords = {
            # Role-related files
            'software_engineer_role_description.md': [
                'software engineer', 'developer', 'programming', 'coding', 'technical role',
                'engineer requirements', 'software development', 'backend', 'frontend', 'full stack'
            ],
            'data_management_role_description.md': [
                'data engineer', 'data management', 'database', 'data pipeline', 'etl',
                'data architecture', 'data platform', 'data infrastructure'
            ],
            'engineering_leadership_role_description.md': [
                'engineering manager', 'tech lead', 'engineering leadership', 'team lead',
                'engineering director', 'vp engineering', 'cto', 'management'
            ],
            'ml_ds_da_role_description.md': [
                'machine learning', 'data scientist', 'data analyst', 'ml engineer',
                'analytics', 'ai', 'artificial intelligence', 'data science', 'statistics'
            ],
            'in_store_operations_role_description.md': [
                'operations', 'retail', 'store operations', 'field operations',
                'user support', 'retail intelligence', 'data validation'
            ],
            
            # Business and culture files
            'company_overview.md': [
                'company', 'mission', 'values', 'culture', 'vision', 'business',
                'organization', 'company culture', 'corporate values'
            ],
            'engineering_overview.md': [
                'engineering culture', 'tech stack', 'engineering practices', 'development process',
                'engineering strategy', 'technology', 'platform', 'architecture'
            ],
            
            # Hiring and career files
            'hiring_guide.md': [
                'hiring', 'recruitment', 'interview', 'hiring process', 'recruiting',
                'candidate', 'interview process', 'hiring standards', 'recruitment process'
            ],
            'engineering_career_path.md': [
                'career path', 'promotion', 'levels', 'career progression', 'advancement',
                'l1', 'l2', 'l3', 'l4', 'l5', 'l6', 'l7', 'l8', 'career framework',
                'job levels', 'seniority', 'career ladder'
            ],
            
            # Data and metrics files
            'column_definitions.md': [
                'columns', 'data fields', 'excel columns', 'data definitions',
                'field definitions', 'data dictionary', 'column meanings'
            ],
            'metrics_definitions.md': [
                'metrics', 'kpi', 'measurements', 'analytics', 'performance indicators',
                'hiring metrics', 'hr metrics', 'time to hire', 'cost per hire'
            ]
        }
        
        # Add specific position description files
        position_files = {
            'Position-Description-MLE1-DS1.md': ['ml engineer 1', 'data scientist 1', 'entry level ml', 'junior data scientist'],
            'Position-Description-DA1.md': ['data analyst 1', 'junior analyst', 'entry level analyst'],
            'Position-Description-MLE2-DS2.md': ['ml engineer 2', 'data scientist 2', 'mid level ml'],
            'Position-Description-DA2.md': ['data analyst 2', 'mid level analyst'],
            'Position-Description-SMLE-SDS.md': ['senior ml engineer', 'senior data scientist'],
            'Position-Description-SDA.md': ['senior data analyst', 'senior analyst'],
            'Position-Description-LSMLE-LSDS.md': ['lead senior ml engineer', 'lead senior data scientist'],
            'Position-Description-LSDA.md': ['lead senior data analyst', 'lead senior analyst'],
            'Position-Description-PMLE-PDS.md': ['principal ml engineer', 'principal data scientist'],
            'Position-Description-PDA.md': ['principal data analyst', 'principal analyst'],
            'Position-Description-SEMLDS Manager.md': ['ml manager', 'data science manager', 'engineering manager ml'],
            'Position-Description-Analytics Manager.md': ['analytics manager', 'data analytics manager']
        }
        
        # Combine all keyword mappings
        all_keywords = {**file_keywords, **position_files}
        
        # Score files based on keyword matches
        file_scores = {}
        description_lower = description.lower()
        
        for filename, keywords in all_keywords.items():
            score = 0
            matched_keywords = []
            
            for keyword in keywords:
                if keyword.lower() in description_lower:
                    # Give higher scores for exact matches and longer keywords
                    keyword_score = len(keyword.split()) * 2 if keyword.lower() == description_lower else len(keyword.split())
                    score += keyword_score
                    matched_keywords.append(keyword)
            
            if score > 0:
                file_scores[filename] = {
                    'score': score,
                    'matched_keywords': matched_keywords
                }
        
        # Sort by score and get top results
        sorted_files = sorted(file_scores.items(), key=lambda x: x[1]['score'], reverse=True)
        top_files = sorted_files[:max_results]
        
        # Get file categories and prepare results
        results = []
        for filename, match_info in top_files:
            # Find which category this file belongs to
            file_category = None
            for category, files in context_manager.contexts.items():
                if filename in files:
                    file_category = category
                    break
            
            results.append({
                'filename': filename,
                'category': file_category,
                'relevance_score': match_info['score'],
                'matched_keywords': match_info['matched_keywords'],
                'file_exists': filename in [f for files in context_manager.contexts.values() for f in files.keys()]
            })
        
        if not results:
            # If no matches found, suggest using the general search
            return {
                "success": True,
                "description": description,
                "matches_found": 0,
                "suggestion": "No specific files matched your description. Try using 'Search_All_Context' to search across all files, or use 'List_Available_Context_Files' to see all available files.",
                "results": []
            }
        
        return {
            "success": True,
            "description": description,
            "matches_found": len(results),
            "results": results,
            "usage_note": f"Use 'Search_Specific_Context_File' with the filename from the top result: '{results[0]['filename']}'"
        }
        
    except Exception as e:
        return {"success": False, "message": f"Error finding relevant context files: {str(e)}"}

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

@mcp.tool(name="Calculate_HR_Metrics", description="Calculate comprehensive HR and recruiting metrics from Excel data. Returns pre-calculated metrics as REFERENCE, plus raw data summary and column context to enable AI-driven insights beyond pre-defined metrics.")
async def calculate_hr_metrics_tool(folder_name: str, file_name: str):
    """Calculate comprehensive HR and recruiting metrics from Excel data
    
    Returns both pre-calculated metrics (as reference/context) AND raw data summaries
    to allow AI to discover additional insights and patterns not captured by pre-defined metrics.
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
        
        # Calculate pre-defined metrics (as reference)
        metrics = hr_analytics.calculate_hiring_metrics(df)
        
        # Detect column types for AI context
        column_types = hr_analytics.detect_column_types(df)
        
        # Get relevant context from context manager
        tool_context = context_manager.get_context_for_tool('Calculate_HR_Metrics')
        
        # Get column definitions for columns in this dataset
        column_definitions = {}
        for col in df.columns:
            col_def = context_manager.get_column_definition(col)
            if col_def:
                column_definitions[col] = col_def
        
        # Generate data preview (first few rows as dict for AI to analyze)
        data_preview = df.head(10).to_dict('records') if len(df) > 0 else []
        
        # Statistical summaries for numeric columns
        numeric_summaries = {}
        for col in df.select_dtypes(include=['int64', 'float64']).columns:
            numeric_summaries[col] = {
                'mean': float(df[col].mean()) if df[col].notna().any() else None,
                'median': float(df[col].median()) if df[col].notna().any() else None,
                'std': float(df[col].std()) if df[col].notna().any() else None,
                'min': float(df[col].min()) if df[col].notna().any() else None,
                'max': float(df[col].max()) if df[col].notna().any() else None,
                'q25': float(df[col].quantile(0.25)) if df[col].notna().any() else None,
                'q75': float(df[col].quantile(0.75)) if df[col].notna().any() else None
            }
        
        # Categorical distributions for AI to analyze
        categorical_distributions = {}
        for col in df.select_dtypes(include=['object']).columns:
            value_counts = df[col].value_counts().head(20)
            categorical_distributions[col] = value_counts.to_dict()
        
        return {
            "success": True,
            "file_name": file_name,
            
            # Pre-calculated metrics (REFERENCE - AI can use these as starting point)
            "pre_calculated_metrics": metrics,
            
            # Raw data for AI to analyze and discover new insights
            "data_summary": {
                "total_rows": len(df),
                "total_columns": len(df.columns),
                "columns": list(df.columns),
                "column_types": column_types,
                "data_preview": data_preview,
                "numeric_summaries": numeric_summaries,
                "categorical_distributions": categorical_distributions
            },
            
            # Context to guide AI analysis
            "context": {
                "column_definitions": column_definitions,
                "business_context": tool_context[:2000] if tool_context else None,  # Truncate for token efficiency
                "analysis_guidance": "Use pre_calculated_metrics as reference. Analyze data_summary to discover additional insights, patterns, correlations, and anomalies not captured by standard metrics. Consider business context when making recommendations."
            }
        }
        
    except Exception as e:
        return {"success": False, "message": f"Error calculating HR metrics: {str(e)}"}

@mcp.tool(name="Analyze_HR_File_Complete", description="Complete analysis of HR Excel file including data quality, metrics, and chart data. Returns structured data + raw data summaries to enable AI to generate custom insights beyond pre-defined analysis.")
async def analyze_hr_file_complete_tool(folder_name: str, file_name: str):
    """Perform complete analysis of HR Excel file including data quality, metrics, and suggested visualizations.
    
    Returns pre-calculated analysis as REFERENCE, plus raw data and context to allow AI to:
    - Discover patterns and correlations not in pre-defined metrics
    - Generate custom insights based on specific data characteristics
    - Make context-aware recommendations using business knowledge
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
        
        # Get context for AI-driven insights
        tool_context = context_manager.get_context_for_tool('Analyze_HR_File_Complete')
        column_types = hr_analytics.detect_column_types(df)
        
        # Get column definitions
        column_definitions = {}
        for col in df.columns:
            col_def = context_manager.get_column_definition(col)
            if col_def:
                column_definitions[col] = col_def
        
        # Data preview for AI analysis
        data_preview = df.head(15).to_dict('records') if len(df) > 0 else []
        
        # Statistical summaries
        numeric_summaries = {}
        for col in df.select_dtypes(include=['int64', 'float64']).columns:
            numeric_summaries[col] = {
                'mean': float(df[col].mean()) if df[col].notna().any() else None,
                'median': float(df[col].median()) if df[col].notna().any() else None,
                'std': float(df[col].std()) if df[col].notna().any() else None,
                'min': float(df[col].min()) if df[col].notna().any() else None,
                'max': float(df[col].max()) if df[col].notna().any() else None
            }
        
        # Categorical distributions
        categorical_distributions = {}
        for col in df.select_dtypes(include=['object']).columns:
            value_counts = df[col].value_counts().head(15)
            categorical_distributions[col] = value_counts.to_dict()
        
        return {
            "success": True,
            "file_name": file_name,
            
            # Basic data info
            "data_summary": {
                "total_rows": len(df),
                "total_columns": len(df.columns),
                "columns": list(df.columns),
                "column_types": column_types,
                "date_range": {
                    "earliest_date": str(df.select_dtypes(include=['datetime64']).min().min()) if not df.select_dtypes(include=['datetime64']).empty else None,
                    "latest_date": str(df.select_dtypes(include=['datetime64']).max().max()) if not df.select_dtypes(include=['datetime64']).empty else None
                },
                "data_preview": data_preview,
                "numeric_summaries": numeric_summaries,
                "categorical_distributions": categorical_distributions
            },
            
            # Pre-calculated analysis (REFERENCE for AI)
            "pre_calculated_analysis": {
                "data_quality": validation_results,
                "hr_metrics": metrics,
                "suggested_charts": chart_suggestions,
                "basic_recommendations": _generate_recommendations(validation_results, metrics)
            },
            
            # Context for AI-driven insights
            "context": {
                "column_definitions": column_definitions,
                "business_context": tool_context[:3000] if tool_context else None,
                "analysis_guidance": (
                    "The pre_calculated_analysis provides standard metrics as REFERENCE. "
                    "Use data_summary to discover additional insights:\n"
                    "- Identify patterns, trends, and correlations not captured by standard metrics\n"
                    "- Analyze distributions and outliers in numeric_summaries\n"
                    "- Examine categorical_distributions for hiring patterns\n"
                    "- Consider business_context and column_definitions when interpreting data\n"
                    "- Generate custom recommendations based on specific data characteristics\n"
                    "- Think creatively about what the data reveals beyond pre-defined metrics"
                )
            }
        }
        
    except Exception as e:
        return {"success": False, "message": f"Error performing complete analysis: {str(e)}"}

@mcp.tool(name="Create_PowerPoint_Report", 
description="""Create professional PowerPoint presentation with comprehensive visualizations AND AI-generated insights.

This enhanced tool combines:
1. Automatic chart generation with data-driven visualizations (time analysis, role distribution, location trends, etc.)
2. AI-generated custom insight slides (recommendations, comparisons, strategic analysis)

The tool automatically generates standard analytical charts from the data, then optionally adds AI-crafted insight slides for deeper strategic recommendations. This provides both comprehensive data visualization AND intelligent business insights in a single presentation.

Optional AI insights can include:
- 'metric': Highlight key metrics with large values
- 'comparison': Side-by-side comparisons
- 'recommendation': Actionable recommendations with rationale
- 'analysis': Detailed analysis with bullet points
- 'chart_with_analysis': Charts with insights
- 'table': Data tables with insights
- 'two_column': Two-column layouts

IMPORTANT - AI Insights Format:
- The ai_insights parameter accepts BOTH JSON string and JSON array formats
- MCP framework auto-parses JSON arrays, so both formats work correctly
- Follow the template structure in docs/AI_INSIGHTS_TEMPLATE.md for proper formatting
- Each insight must have: {"title": "...", "data": {"type": "...", ...}}

Use this tool when you want both automated data visualizations AND custom AI insights in one presentation.""")
async def create_powerpoint_report_tool(
    file_name: str, 
    folder_name: Optional[str] = None, 
    output_folder: Optional[str] = None, 
    presentation_title: Optional[str] = None,
    ai_insights: Optional[Union[str, list]] = None  # JSON string or list of custom AI insight slides
):
    """Create professional PowerPoint presentation from HR data analysis with optional AI insights
    
    Args:
        file_name: Excel file to analyze
        folder_name: Source folder (default: "Recruiting Data")
        output_folder: Output folder (default: "AI Generated Reports")
        presentation_title: Presentation title (default: "HR Recruiting Analytics Report")
        ai_insights: Optional JSON string or list with custom AI insight slides to add after charts
                    Accepts both formats (MCP framework auto-parses JSON arrays):
                    - JSON string: "[{...}, {...}]"
                    - JSON array: [{...}, {...}]
                    See docs/AI_INSIGHTS_TEMPLATE.md for complete format specification
                    Format: [{"title": "...", "data": {"type": "metric|comparison|...", ...}}]
    """
    try:
        # Set default folders if not specified
        if folder_name is None:
            folder_name = "Recruiting Data"
        if output_folder is None:
            output_folder = "AI Generated Reports"
        
        # Get the Excel content from source folder
        content_result = get_document_content(folder_name, file_name)
        if not content_result.get("success", True):
            return {"success": False, "message": "Failed to retrieve Excel file"}
        
        # Parse content into DataFrame
        df = hr_analytics.parse_excel_content(content_result.get("content", ""))
        if df.empty:
            return {"success": False, "message": "No data found in Excel file"}
        
        # Get relevant context for this tool (columns, business, metrics, recruiting, roles)
        tool_context = context_manager.get_context_for_tool('Create_PowerPoint_Report')
        
        # Get column definitions for columns in the dataset
        columns_in_data = df.columns.tolist()
        column_definitions = {}
        for col in columns_in_data:
            definition = context_manager.get_column_definition(col)
            if definition:
                column_definitions[col] = definition
        
        # Perform analysis with context
        validation_results = hr_analytics.validate_data_quality(df)
        metrics = hr_analytics.calculate_hiring_metrics(df)
        analysis = {
            'validation': validation_results, 
            'metrics': metrics,
            'context': tool_context,
            'column_definitions': column_definitions,
            'data_summary': {
                'total_rows': len(df),
                'total_columns': len(df.columns),
                'columns': columns_in_data
            }
        }
        
        # Create PowerPoint presentation
        ppt = PowerPointHelper()
        
        # Title slide
        if presentation_title is None:
            presentation_title = "HR Recruiting Analytics Report"
        ppt.create_title_slide(presentation_title)
        
        # Chart slides with insights - dynamically generated based on available data
        chart_types = [
            {
                'type': 'time_by_step',
                'title': 'Time Spent in Each Recruiting Step',
                'insights_key': 'time_analysis'
            },
            {
                'type': 'role_distribution',
                'title': 'Hiring by Role',
                'insights_key': 'role_analysis'
            },
            {
                'type': 'location_distribution',
                'title': 'Hiring by Location',
                'insights_key': 'location_analysis'
            },
            {
                'type': 'time_distribution',
                'title': 'Total Time-to-Hire Distribution',
                'insights_key': 'time_distribution'
            },
            {
                'type': 'department_distribution',
                'title': 'Hiring by Department',
                'insights_key': 'department_analysis'
            }
        ]
        
        slides_created = 0
        for chart_config in chart_types:
            try:
                # Generate chart data
                chart_data_result = hr_analytics.generate_chart_data(df, chart_config['type'])
                chart_data = chart_data_result.get('data', {})
                
                if not chart_data.get('labels') or not chart_data.get('values'):
                    continue
                
                # Generate insights for this chart
                insights = _generate_chart_insights(analysis, chart_config['type'], chart_data)
                
                # Create slide
                ppt.create_chart_slide(
                    chart_config['title'],
                    chart_data,
                    chart_config['type'],
                    insights
                )
                slides_created += 1
                
            except Exception as e:
                logger.error(f"Error creating slide for {chart_config['type']}: {e}")
                continue
        
        # Add AI-generated insight slides if provided
        ai_slides_created = 0
        if ai_insights:
            try:
                import json
                
                # Handle both string and list inputs (MCP framework may auto-parse JSON)
                if isinstance(ai_insights, str):
                    insights_data = json.loads(ai_insights)
                elif isinstance(ai_insights, list):
                    insights_data = ai_insights
                else:
                    logger.warning(f"AI insights must be a JSON string or list, got {type(ai_insights)}, skipping AI slides")
                    insights_data = None
                
                if insights_data and isinstance(insights_data, list):
                    # VALIDATE AI INSIGHTS STRUCTURE AGAINST TEMPLATE
                    is_valid, validation_errors, valid_insights = insights_validator.validate_insights(insights_data)
                    
                    if not is_valid:
                        # Log detailed validation report
                        validation_report = insights_validator.format_validation_report(
                            validation_errors, 
                            len(valid_insights), 
                            len(insights_data)
                        )
                        logger.error(validation_report)
                        
                        # If no valid insights, return error
                        if len(valid_insights) == 0:
                            return {
                                "success": False,
                                "message": "AI insights validation failed - no valid insights to create slides",
                                "validation_errors": validation_errors,
                                "hint": "Please review docs/AI_INSIGHTS_TEMPLATE.md for correct structure"
                            }
                        
                        # Log warning but continue with valid insights
                        logger.warning(f"Proceeding with {len(valid_insights)} valid insights out of {len(insights_data)} total")
                    else:
                        logger.info(f"✓ All {len(insights_data)} AI insights validated successfully")
                    
                    # Create slides from VALIDATED insights only
                    for insight_def in valid_insights:
                        try:
                            insight_title = insight_def.get('title', 'Insight')
                            insight_data = insight_def.get('data', {})
                            
                            # Use PowerPointHelper's create_insight_slide method
                            # which supports: metric, comparison, recommendation, analysis,
                            # chart_with_analysis, table, two_column
                            ppt.create_insight_slide(insight_title, insight_data)
                            ai_slides_created += 1
                            
                        except Exception as e:
                            logger.error(f"Error creating AI insight slide '{insight_def.get('title', 'unknown')}': {e}")
                            continue
                    
                    logger.info(f"Added {ai_slides_created} AI-generated insight slides")
                    
            except json.JSONDecodeError as e:
                logger.error(f"Invalid JSON format for ai_insights: {e}. See docs/AI_INSIGHTS_TEMPLATE.md for proper format")
            except Exception as e:
                logger.error(f"Error processing AI insights: {e}. See docs/AI_INSIGHTS_TEMPLATE.md for proper format")
        
        # Save PowerPoint to bytes
        pptx_bytes = ppt.save_to_bytes()
        
        # Upload to SharePoint (output folder)
        pptx_filename = f"report_{file_name.replace('.xlsx', '')}.pptx"
        upload_result = _upload_file_helper(output_folder, pptx_filename, base64.b64encode(pptx_bytes).decode(), is_base64=True)
        
        total_slides = slides_created + 1 + ai_slides_created  # +1 for title slide
        
        result = {
            "success": True,
            "message": f"PowerPoint presentation created: {pptx_filename}",
            "file_name": pptx_filename,
            "slides_created": total_slides,
            "chart_slides": slides_created,
            "download_info": "File uploaded to SharePoint and ready for download"
        }
        
        if ai_slides_created > 0:
            result["ai_insight_slides"] = ai_slides_created
            result["message"] += f" (includes {ai_slides_created} AI insight slides)"
        
        return result
        
    except Exception as e:
        return {"success": False, "message": f"Error creating PowerPoint presentation: {str(e)}"}


# Helper functions

def _get_theme_colors(theme: str) -> dict:
    """Return color scheme for theme"""
    themes = {
        "professional": {
            "primary": RGBColor(68, 114, 196),    # Blue
            "secondary": RGBColor(237, 125, 49),   # Orange
            "accent": RGBColor(165, 165, 165),     # Gray
            "text": RGBColor(0, 0, 0),             # Black
            "background": RGBColor(255, 255, 255)  # White
        },
        "modern": {
            "primary": RGBColor(0, 176, 240),      # Cyan
            "secondary": RGBColor(255, 192, 0),    # Yellow
            "accent": RGBColor(112, 48, 160),      # Purple
            "text": RGBColor(64, 64, 64),          # Dark gray
            "background": RGBColor(255, 255, 255)
        },
        "minimal": {
            "primary": RGBColor(50, 50, 50),       # Charcoal
            "secondary": RGBColor(150, 150, 150),  # Gray
            "accent": RGBColor(200, 200, 200),     # Light gray
            "text": RGBColor(0, 0, 0),
            "background": RGBColor(255, 255, 255)
        },
        "default": {
            "primary": RGBColor(68, 114, 196),
            "secondary": RGBColor(237, 125, 49),
            "accent": RGBColor(165, 165, 165),
            "text": RGBColor(0, 0, 0),
            "background": RGBColor(255, 255, 255)
        }
    }
    return themes.get(theme, themes["default"])


def _create_title_slide(prs, title: str, theme_colors: dict, subtitle: str = ""):
    """Create title slide"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank layout
    
    # Title
    title_box = slide.shapes.add_textbox(Inches(1), Inches(2.5), Inches(8), Inches(1.5))
    title_frame = title_box.text_frame
    title_frame.text = title
    title_para = title_frame.paragraphs[0]
    title_para.alignment = PP_PARAGRAPH_ALIGNMENT.CENTER
    title_para.font.size = Pt(44)
    title_para.font.bold = True
    title_para.font.color.rgb = theme_colors["primary"]
    
    # Subtitle if provided
    if subtitle:
        subtitle_box = slide.shapes.add_textbox(Inches(1), Inches(4.5), Inches(8), Inches(0.8))
        subtitle_frame = subtitle_box.text_frame
        subtitle_frame.text = subtitle
        subtitle_para = subtitle_frame.paragraphs[0]
        subtitle_para.alignment = PP_PARAGRAPH_ALIGNMENT.CENTER
        subtitle_para.font.size = Pt(20)
        subtitle_para.font.color.rgb = theme_colors["secondary"]


def _create_content_slide(prs, slide_def: dict, theme_colors: dict):
    """Create content slide with title and bullet points"""
    slide = prs.slides.add_slide(prs.slide_layouts[5])  # Title only layout
    
    # Title
    title = slide.shapes.title
    title.text = slide_def.get('title', 'Untitled')
    title.text_frame.paragraphs[0].font.color.rgb = theme_colors["primary"]
    
    # Content
    content = slide_def.get('content', '')
    left = Inches(0.8)
    top = Inches(1.8)
    width = Inches(8.4)
    height = Inches(5)
    
    text_box = slide.shapes.add_textbox(left, top, width, height)
    tf = text_box.text_frame
    tf.word_wrap = True
    
    # Handle different content formats
    if isinstance(content, list):
        # List of bullet points
        for i, item in enumerate(content):
            if isinstance(item, dict):
                # Structured bullet with level
                text = item.get('text', '')
                level = item.get('level', 0)
            else:
                text = str(item)
                level = 0
            
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = text
            p.level = level
            p.font.size = Pt(18 - (level * 2))
            p.space_after = Pt(12)
    else:
        # Plain text content
        tf.text = str(content)
        tf.paragraphs[0].font.size = Pt(18)


def _create_two_column_slide(prs, slide_def: dict, theme_colors: dict):
    """Create two-column slide"""
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    
    # Title
    title = slide.shapes.title
    title.text = slide_def.get('title', 'Untitled')
    title.text_frame.paragraphs[0].font.color.rgb = theme_colors["primary"]
    
    # Left column
    left_content = slide_def.get('left_column', {})
    left_title = left_content.get('title', '')
    left_points = left_content.get('content', [])
    
    left_box = slide.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(4), Inches(5))
    tf = left_box.text_frame
    tf.word_wrap = True
    
    if left_title:
        p = tf.paragraphs[0]
        p.text = left_title
        p.font.size = Pt(20)
        p.font.bold = True
        p.font.color.rgb = theme_colors["primary"]
        p.space_after = Pt(12)
    
    for item in (left_points if isinstance(left_points, list) else [left_points]):
        p = tf.add_paragraph() if left_title else (tf.paragraphs[0] if item == left_points[0] else tf.add_paragraph())
        p.text = f"• {item}" if not str(item).startswith('•') else str(item)
        p.font.size = Pt(16)
        p.space_after = Pt(8)
    
    # Right column
    right_content = slide_def.get('right_column', {})
    right_title = right_content.get('title', '')
    right_points = right_content.get('content', [])
    
    right_box = slide.shapes.add_textbox(Inches(5.2), Inches(1.8), Inches(4), Inches(5))
    tf = right_box.text_frame
    tf.word_wrap = True
    
    if right_title:
        p = tf.paragraphs[0]
        p.text = right_title
        p.font.size = Pt(20)
        p.font.bold = True
        p.font.color.rgb = theme_colors["secondary"]
        p.space_after = Pt(12)
    
    for item in (right_points if isinstance(right_points, list) else [right_points]):
        p = tf.add_paragraph() if right_title else (tf.paragraphs[0] if item == right_points[0] else tf.add_paragraph())
        p.text = f"• {item}" if not str(item).startswith('•') else str(item)
        p.font.size = Pt(16)
        p.space_after = Pt(8)


def _create_chart_slide(prs, slide_def: dict, theme_colors: dict):
    """Create slide with chart"""
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    
    # Title
    title = slide.shapes.title
    title.text = slide_def.get('title', 'Untitled')
    title.text_frame.paragraphs[0].font.color.rgb = theme_colors["primary"]
    
    # Chart data
    chart_data_def = slide_def.get('chart_data', {})
    chart_type = chart_data_def.get('type', 'bar').lower()
    
    # Prepare chart data
    chart_data = CategoryChartData()
    chart_data.categories = chart_data_def.get('categories', [])
    
    datasets = chart_data_def.get('datasets', [])
    for dataset in datasets:
        series_name = dataset.get('label', 'Series')
        series_values = dataset.get('data', [])
        chart_data.add_series(series_name, series_values)
    
    # Determine chart type
    chart_type_map = {
        'bar': XL_CHART_TYPE.BAR_CLUSTERED,
        'column': XL_CHART_TYPE.COLUMN_CLUSTERED,
        'line': XL_CHART_TYPE.LINE,
        'pie': XL_CHART_TYPE.PIE,
        'area': XL_CHART_TYPE.AREA
    }
    xl_chart_type = chart_type_map.get(chart_type, XL_CHART_TYPE.COLUMN_CLUSTERED)
    
    # Add chart
    x, y, cx, cy = Inches(1), Inches(2), Inches(8), Inches(4.5)
    chart = slide.shapes.add_chart(xl_chart_type, x, y, cx, cy, chart_data).chart
    
    # Chart formatting
    if len(datasets) <= 1:
        chart.has_legend = False
    
    chart_title = chart_data_def.get('title', '')
    if chart_title:
        chart.chart_title.text_frame.text = chart_title


def _create_table_slide(prs, slide_def: dict, theme_colors: dict):
    """Create slide with table"""
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    
    # Title
    title = slide.shapes.title
    title.text = slide_def.get('title', 'Untitled')
    title.text_frame.paragraphs[0].font.color.rgb = theme_colors["primary"]
    
    # Table data
    table_data = slide_def.get('table_data', {})
    headers = table_data.get('headers', [])
    rows = table_data.get('rows', [])
    
    if not headers or not rows:
        return
    
    # Create table
    cols = len(headers)
    table_rows = len(rows) + 1  # +1 for header
    
    x, y = Inches(1), Inches(2)
    cx, cy = Inches(8), Inches(4.5)
    
    table = slide.shapes.add_table(table_rows, cols, x, y, cx, cy).table
    
    # Set headers
    for col_idx, header in enumerate(headers):
        cell = table.cell(0, col_idx)
        cell.text = str(header)
        cell.text_frame.paragraphs[0].font.bold = True
        cell.text_frame.paragraphs[0].font.size = Pt(14)
        cell.fill.solid()
        cell.fill.fore_color.rgb = theme_colors["primary"]
        cell.text_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
    
    # Set data rows
    for row_idx, row in enumerate(rows, start=1):
        for col_idx, value in enumerate(row):
            cell = table.cell(row_idx, col_idx)
            cell.text = str(value)
            cell.text_frame.paragraphs[0].font.size = Pt(12)


def _create_image_slide(prs, slide_def: dict, theme_colors: dict):
    """Create slide with image"""
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    
    # Title
    title = slide.shapes.title
    title.text = slide_def.get('title', 'Untitled')
    title.text_frame.paragraphs[0].font.color.rgb = theme_colors["primary"]
    
    # Image handling would go here
    # Note: Requires additional logic for base64 or URL image loading
    # Placeholder for now
    left = Inches(2)
    top = Inches(2.5)
    width = Inches(6)
    height = Inches(4)
    
    text_box = slide.shapes.add_textbox(left, top, width, height)
    tf = text_box.text_frame
    tf.text = "[Image placeholder - image support requires additional implementation]"
    tf.paragraphs[0].alignment = PP_PARAGRAPH_ALIGNMENT.CENTER
    tf.paragraphs[0].font.size = Pt(16)
    tf.paragraphs[0].font.color.rgb = theme_colors["accent"]


def _parse_markdown_to_slides(markdown_content: str) -> list:
    """Parse Markdown content into slide definitions"""
    slides = []
    current_slide = None
    
    lines = markdown_content.split('\n')
    
    for line in lines:
        line = line.strip()
        
        # H1 = new slide title
        if line.startswith('# '):
            if current_slide:
                slides.append(current_slide)
            current_slide = {
                'type': 'content',
                'title': line[2:].strip(),
                'content': []
            }
        
        # H2 = subtitle or section
        elif line.startswith('## ') and current_slide:
            current_slide['content'].append({'text': line[3:].strip(), 'level': 0})
        
        # Bullet points
        elif (line.startswith('- ') or line.startswith('* ')) and current_slide:
            current_slide['content'].append({'text': line[2:].strip(), 'level': 1})
        
        # Regular text
        elif line and current_slide and not line.startswith('#'):
            if current_slide['content'] and isinstance(current_slide['content'][-1], str):
                current_slide['content'][-1] += ' ' + line
            else:
                current_slide['content'].append(line)
    
    if current_slide:
        slides.append(current_slide)
    
    return slides


def _parse_html_to_slides(html_content: str) -> list:
    """Parse HTML content into slide definitions"""
    # Basic HTML parsing - would need BeautifulSoup for production
    # This is a simplified version
    slides = []
    
    # Look for section tags or h1 tags as slide boundaries
    import re
    
    # Split by <h1> tags
    sections = re.split(r'<h1[^>]*>(.*?)</h1>', html_content, flags=re.IGNORECASE | re.DOTALL)
    
    for i in range(1, len(sections), 2):
        if i < len(sections):
            title = re.sub(r'<[^>]+>', '', sections[i]).strip()
            content_html = sections[i+1] if i+1 < len(sections) else ''
            
            # Extract bullet points
            bullets = re.findall(r'<li[^>]*>(.*?)</li>', content_html, flags=re.IGNORECASE | re.DOTALL)
            content = [re.sub(r'<[^>]+>', '', bullet).strip() for bullet in bullets]
            
            if not content:
                # Extract paragraphs
                paras = re.findall(r'<p[^>]*>(.*?)</p>', content_html, flags=re.IGNORECASE | re.DOTALL)
                content = [re.sub(r'<[^>]+>', '', para).strip() for para in paras]
            
            slides.append({
                'type': 'content',
                'title': title,
                'content': content if content else ['No content']
            })
    
    return slides
