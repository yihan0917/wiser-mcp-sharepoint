"""
Enhanced Context Manager for MCP Server
Loads and manages multiple context files for intelligent tool responses
"""
import os
import re
from typing import Dict, Any, Optional, List
from .common import logger

class ContextManager:
    """Manages multiple context files and provides context for MCP tools"""
    
    # Define context categories and their associated files
    CONTEXT_CATEGORIES = {
        'columns': ['column_definitions.md'],
        'metrics': ['metrics_definitions.md'],
        'business': ['company_overview.md', 'engineering_overview.md'],
        'recruiting': ['hiring_guide.md', 'career_path.md'],
        'roles': [
            # Software Engineering roles
            'software_engineer_role_description.md',
            # Data Management roles
            'data_management_role_description.md',
            # Engineering Leadership roles
            'engineering_leadership_role_description.md',
            # ML/DS/DA roles - Index
            'ml_ds_da_role_description.md',
            # ML/DS/DA roles - Detailed descriptions
            'Position-Description-MLE1-DS1.md',
            'Position-Description-DA1.md',
            'Position-Description-MLE2-DS2.md',
            'Position-Description-DA2.md',
            'Position-Description-SMLE-SDS.md',
            'Position-Description-SDA.md',
            'Position-Description-LSMLE-LSDS.md',
            'Position-Description-LSDA.md',
            'Position-Description-PMLE-PDS.md',
            'Position-Description-PDA.md',
            'Position-Description-SEMLDS Manager.md',
            'Position-Description-Analytics Manager.md'
        ]
    }
    
    # Map tools to the context categories they need
    TOOL_CONTEXT_MAP = {
        'Analyze_HR_File_Complete': ['columns', 'business', 'metrics', 'roles'],
        'Calculate_HR_Metrics': ['columns', 'metrics', 'recruiting', 'roles'],
        'Validate_Excel_Data_Quality': ['columns', 'recruiting', 'roles'],
        'Create_PowerPoint_Report': ['columns', 'business', 'metrics', 'recruiting', 'roles'],
        'Create_Excel_With_Charts': ['columns', 'metrics'],
        'Generate_Chart_Data': ['columns', 'metrics'],
        'Get_Document_Content': ['columns'],
        'List_SharePoint_Documents': [],
        'List_SharePoint_Folders': [],
        'Upload_Document': [],
        'Download_Document': [],
        'Delete_Document': [],
        'Create_Folder': []
    }
    
    def __init__(self):
        self.contexts = {}
        self.column_definitions = {}
        self._load_all_contexts()
    
    def _load_all_contexts(self):
        """Load all context files from the context directory"""
        try:
            # Get the context directory
            current_dir = os.path.dirname(__file__)
            context_dir = os.path.join(current_dir, 'context')
            
            if not os.path.exists(context_dir):
                logger.warning(f"Context directory not found at {context_dir}")
                return
            
            # Load each context file
            for category, files in self.CONTEXT_CATEGORIES.items():
                self.contexts[category] = {}
                for filename in files:
                    filepath = os.path.join(context_dir, filename)
                    if os.path.exists(filepath):
                        with open(filepath, 'r', encoding='utf-8') as f:
                            content = f.read()
                            self.contexts[category][filename] = content
                            logger.info(f"Loaded context: {filename} ({len(content)} chars)")
                            
                            # Special handling for column definitions
                            if filename == 'column_definitions.md':
                                self.column_definitions = self._parse_column_definitions(content)
                    else:
                        logger.warning(f"Context file not found: {filepath}")
            
            logger.info(f"Context Manager initialized with {len(self.contexts)} categories")
            
        except Exception as e:
            logger.error(f"Error loading contexts: {e}")
    
    def _parse_column_definitions(self, content: str) -> Dict[str, str]:
        """Parse markdown content to extract column definitions"""
        try:
            definitions = {}
            lines = content.split('\n')
            
            for line in lines:
                line = line.strip()
                
                # Look for column definitions in format: - **Column Name** - Definition
                if line.startswith('- **') and '**' in line[4:]:
                    # Extract column name and definition
                    match = re.match(r'- \*\*(.*?)\*\*\s*[-–]\s*(.*)', line)
                    if match:
                        column_name = match.group(1).strip()
                        definition = match.group(2).strip()
                        definitions[column_name] = definition
            
            logger.info(f"Parsed {len(definitions)} column definitions")
            return definitions
            
        except Exception as e:
            logger.error(f"Error parsing column definitions: {e}")
            return {}
    
    def get_context_for_tool(self, tool_name: str) -> str:
        """Get relevant context for a specific tool"""
        try:
            categories = self.TOOL_CONTEXT_MAP.get(tool_name, [])
            if not categories:
                return ""
            
            context_parts = []
            for category in categories:
                if category in self.contexts:
                    for filename, content in self.contexts[category].items():
                        context_parts.append(f"=== {filename} ===\n{content}\n")
            
            return "\n".join(context_parts)
            
        except Exception as e:
            logger.error(f"Error getting context for tool {tool_name}: {e}")
            return ""
    
    def get_context_by_category(self, category: str) -> str:
        """Get all context for a specific category"""
        try:
            if category not in self.contexts:
                return ""
            
            context_parts = []
            for filename, content in self.contexts[category].items():
                context_parts.append(f"=== {filename} ===\n{content}\n")
            
            return "\n".join(context_parts)
            
        except Exception as e:
            logger.error(f"Error getting context for category {category}: {e}")
            return ""
    
    def get_all_context(self) -> Dict[str, Dict[str, str]]:
        """Get all loaded context"""
        return self.contexts.copy()
    
    def get_column_definition(self, column_name: str) -> Optional[str]:
        """Get definition for a specific column"""
        return self.column_definitions.get(column_name)
    
    def get_matching_columns(self, column_names: List[str]) -> Dict[str, str]:
        """Get definitions for columns that exist in the dictionary"""
        matching = {}
        for column_name in column_names:
            if column_name in self.column_definitions:
                matching[column_name] = self.column_definitions[column_name]
        return matching
    
    def get_all_column_definitions(self) -> Dict[str, str]:
        """Get all available column definitions"""
        return self.column_definitions.copy()
    
    def search_context(self, search_term: str, categories: Optional[List[str]] = None) -> List[Dict[str, Any]]:
        """Search across all context for a specific term"""
        try:
            results = []
            search_categories = categories if categories else list(self.contexts.keys())
            
            for category in search_categories:
                if category not in self.contexts:
                    continue
                
                for filename, content in self.contexts[category].items():
                    if search_term.lower() in content.lower():
                        # Find the line containing the search term
                        lines = content.split('\n')
                        for i, line in enumerate(lines):
                            if search_term.lower() in line.lower():
                                # Get context around the match
                                start = max(0, i - 2)
                                end = min(len(lines), i + 3)
                                excerpt = '\n'.join(lines[start:end])
                                
                                results.append({
                                    'category': category,
                                    'file': filename,
                                    'line': i + 1,
                                    'excerpt': excerpt
                                })
                                break  # Only first match per file
            
            return results
            
        except Exception as e:
            logger.error(f"Error searching context: {e}")
            return []
    
    def get_context_summary(self) -> Dict[str, Any]:
        """Get summary statistics about loaded context"""
        try:
            total_files = sum(len(files) for files in self.contexts.values())
            total_chars = sum(
                sum(len(content) for content in files.values())
                for files in self.contexts.values()
            )
            
            return {
                'total_files': total_files,
                'total_characters': total_chars,
                'categories': list(self.contexts.keys()),
                'column_definitions_count': len(self.column_definitions),
                'files_by_category': {
                    cat: list(files.keys()) 
                    for cat, files in self.contexts.items()
                }
            }
            
        except Exception as e:
            logger.error(f"Error getting context summary: {e}")
            return {}
    
    def reload_contexts(self):
        """Reload all context files (useful for development)"""
        self.contexts = {}
        self.column_definitions = {}
        self._load_all_contexts()
    
    def enhance_excel_text(self, content: str, sheets_dict: Dict[str, Any] = None) -> str:
        """Enhance Excel text content with column definitions"""
        try:
            lines = content.split('\n')
            enhanced_lines = []
            
            for line in lines:
                enhanced_lines.append(line)
                
                # Check if this line contains headers
                if line.startswith('HEADERS:'):
                    # Extract column names from the headers line
                    headers_text = line[8:].strip()  # Remove 'HEADERS: '
                    column_names = [col.strip() for col in headers_text.split('|')]
                    
                    # Find matching definitions
                    matching_definitions = self.get_matching_columns(column_names)
                    
                    if matching_definitions:
                        enhanced_lines.append("\n--- COLUMN DEFINITIONS ---")
                        for column, definition in matching_definitions.items():
                            enhanced_lines.append(f"• {column}: {definition}")
                        enhanced_lines.append("--- END COLUMN DEFINITIONS ---\n")
            
            return '\n'.join(enhanced_lines)
            
        except Exception as e:
            logger.error(f"Error enhancing Excel text: {e}")
            return content

# Global instance
context_manager = ContextManager()
