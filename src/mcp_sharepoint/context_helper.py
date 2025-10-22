"""
Context helper for providing Excel column definitions from a general dictionary
"""
import os
import re
from typing import Dict, Any, Optional, List
from .common import logger

class ContextHelper:
    """Helper class for providing context about Excel columns from a general dictionary"""
    
    def __init__(self):
        self.column_definitions = self._load_column_definitions()
    
    def _load_column_definitions(self) -> Dict[str, str]:
        """Load column definitions from Markdown file"""
        try:
            # Get the directory of this file
            current_dir = os.path.dirname(__file__)
            definitions_path = os.path.join(current_dir, 'column_definitions.md')
            
            if os.path.exists(definitions_path):
                with open(definitions_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                return self._parse_column_definitions(content)
            else:
                logger.warning(f"Column definitions file not found at {definitions_path}")
                return {}
        except Exception as e:
            logger.error(f"Error loading column definitions: {e}")
            return {}
    
    def _parse_column_definitions(self, content: str) -> Dict[str, str]:
        """Parse markdown content to extract column definitions"""
        try:
            definitions = {}
            lines = content.split('\n')
            
            for line in lines:
                line = line.strip()
                
                # Look for column definitions (- **Column**: Definition)
                if line.startswith('- **') and '**:' in line:
                    match = re.match(r'- \*\*(.*?)\*\*:\s*(.*)', line)
                    if match:
                        column_name = match.group(1).strip()
                        definition = match.group(2).strip()
                        definitions[column_name] = definition
            
            return definitions
            
        except Exception as e:
            logger.error(f"Error parsing column definitions: {e}")
            return {}
    
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
    
    def search_definitions(self, search_term: str) -> List[Dict[str, str]]:
        """Search for column definitions containing a term"""
        try:
            results = []
            for column_name, definition in self.column_definitions.items():
                if (search_term.lower() in column_name.lower() or 
                    search_term.lower() in definition.lower()):
                    results.append({
                        'column': column_name,
                        'definition': definition
                    })
            return results
        except Exception as e:
            logger.error(f"Error searching definitions: {e}")
            return []
    
    def get_all_definitions(self) -> Dict[str, str]:
        """Get all available column definitions"""
        return self.column_definitions.copy()

# Global instance
context_helper = ContextHelper()
