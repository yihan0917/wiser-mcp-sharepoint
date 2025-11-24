"""
SharePoint MCP tools using Microsoft Graph API
"""
import base64, os, io
from functools import wraps
from typing import Optional, Dict, Any, List
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from .common import logger, mcp, ACCESS_TOKEN, SITE_ID, DRIVE_ID, make_graph_request
from .resources import list_folders, list_documents, get_document_content, download_document
from .context_manager import context_manager
from .analytics_helper import hr_analytics
from .visualization_helper import visualization_helper
from .powerpoint_helper import PowerPointHelper

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

# Visualization and Export Tools
@mcp.tool(name="Create_Excel_With_Charts", description="Create Excel file with embedded charts and download link")
async def create_excel_with_charts_tool(folder_name: str, file_name: str, chart_types: Optional[List[str]] = None):
    """Create Excel file with embedded charts from HR data"""
    try:
        # Get the Excel content
        content_result = get_document_content(folder_name, file_name)
        if not content_result.get("success", True):
            return {"success": False, "message": "Failed to retrieve Excel file"}
        
        # Parse content into DataFrame
        df = hr_analytics.parse_excel_content(content_result.get("content", ""))
        if df.empty:
            return {"success": False, "message": "No data found in Excel file"}
        
        # Generate chart configurations
        if not chart_types:
            chart_types = ['source_effectiveness', 'hiring_trends', 'time_to_hire_distribution']
        
        chart_configs = []
        for chart_type in chart_types:
            try:
                chart_data = hr_analytics.generate_chart_data(df, chart_type)
                if chart_data.get('data') and len(chart_data['data'].get('values', [])) > 0:
                    chart_configs.append(chart_data)
            except:
                continue
        
        # Create Excel with charts
        excel_bytes = visualization_helper.create_excel_with_charts(df, chart_configs, file_name)
        
        # Upload to SharePoint
        excel_filename = f"charts_{file_name.replace('.xlsx', '')}_analysis.xlsx"
        upload_result = _upload_file_helper(folder_name, excel_filename, base64.b64encode(excel_bytes).decode(), is_base64=True)
        
        return {
            "success": True,
            "message": f"Excel file with charts created: {excel_filename}",
            "file_name": excel_filename,
            "charts_included": len(chart_configs),
            "chart_types": [config.get('chart_type') for config in chart_configs],
            "download_info": "File uploaded to SharePoint and ready for download"
        }
        
    except Exception as e:
        return {"success": False, "message": f"Error creating Excel with charts: {str(e)}"}

@mcp.tool(name="Create_PowerPoint_Report", 
description="Create professional PowerPoint presentation with charts, insights, and data definitions. This tool creates basic PowerPoint reports with standard charts. For comprehensive, detailed presentations with custom visualizations, Claude should generate the presentation manually and then upload it to SharePoint.")
async def create_powerpoint_report_tool(file_name: str, folder_name: Optional[str] = None, output_folder: Optional[str] = None, presentation_title: Optional[str] = None):
    """Create professional PowerPoint presentation from HR data analysis"""
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
        
        # Save PowerPoint to bytes
        pptx_bytes = ppt.save_to_bytes()
        
        # Upload to SharePoint (output folder)
        pptx_filename = f"report_{file_name.replace('.xlsx', '')}.pptx"
        upload_result = _upload_file_helper(output_folder, pptx_filename, base64.b64encode(pptx_bytes).decode(), is_base64=True)
        
        return {
            "success": True,
            "message": f"PowerPoint presentation created: {pptx_filename}",
            "file_name": pptx_filename,
            "slides_created": slides_created + 1,  # +1 for title slide
            "download_info": "File uploaded to SharePoint and ready for download"
        }
        
    except Exception as e:
        return {"success": False, "message": f"Error creating PowerPoint presentation: {str(e)}"}




