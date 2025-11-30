"""
Enhanced Context Manager for MCP Server
Loads and manages multiple context files for intelligent tool responses
"""
import os
import re
import hashlib
import yaml
from functools import lru_cache
from collections import defaultdict
from typing import Dict, Any, Optional, List
from .common import logger

class ContextManager:
    """Manages multiple context files and provides context for MCP tools"""
    
    # Define context categories and their associated files
    CONTEXT_CATEGORIES = {
        'index': ['_context_index.md'],  # Quick reference guide
        'columns': ['column_definitions.md'],
        'metrics': ['metrics_definitions.md'],
        'business': ['company_overview.md', 'engineering_overview.md'],
        'recruiting': ['hiring_guide.md', 'engineering_career_path.md'],
        'dept_engineering': [
            # Software Engineering roles
            'software_engineer_role_description.md',
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
        ],
        'dept_data_management': [
            # Data Management roles (may be separate dept or under Engineering)
            'data_management_role_description.md'
        ],
        'dept_operations': [
            # In-Store Operations roles (Operations department)
            'in_store_operations_role_description.md'
        ]
    }
    
    # Map tools to the context categories they need
    TOOL_CONTEXT_MAP = {
        'Analyze_HR_File_Complete': ['index', 'columns', 'business', 'metrics', 'dept_engineering', 'dept_data_management', 'dept_operations'],
        'Calculate_HR_Metrics': ['index', 'columns', 'metrics', 'recruiting', 'dept_engineering', 'dept_data_management', 'dept_operations'],
        'Validate_Excel_Data_Quality': ['index', 'columns', 'recruiting', 'dept_engineering', 'dept_data_management', 'dept_operations'],
        'Create_PowerPoint_Report': ['index', 'columns', 'business', 'metrics', 'recruiting', 'dept_engineering', 'dept_data_management', 'dept_operations'],
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
        self.metadata_index = {}  # Store file metadata
        self.column_definitions = {}
        self.file_hashes = {}  # Track file changes
        self.search_index = defaultdict(set)  # Fast term lookup
        self.context_cache = {}  # Cache assembled contexts
        self._load_all_contexts()
        self._build_search_index()
    
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
                            
                            # Parse metadata and content
                            metadata, body = self._parse_file_with_metadata(content, filename)
                            self.contexts[category][filename] = body
                            if metadata:
                                self.metadata_index[filename] = metadata
                            
                            logger.info(f"Loaded context: {filename} ({len(body)} chars)")
                            
                            # Special handling for column definitions
                            if filename == 'column_definitions.md':
                                self.column_definitions = self._parse_column_definitions(body)
                    else:
                        logger.warning(f"Context file not found: {filepath}")
            
            logger.info(f"Context Manager initialized with {len(self.contexts)} categories")
            
        except Exception as e:
            logger.error(f"Error loading contexts: {e}")
    
    def _parse_file_with_metadata(self, content: str, filename: str) -> tuple[Optional[Dict], str]:
        """Parse YAML frontmatter and return metadata + body content"""
        try:
            if content.startswith('---\n'):
                parts = content.split('---\n', 2)
                if len(parts) >= 3:
                    # Parse YAML metadata
                    metadata_yaml = parts[1]
                    body = parts[2].strip()
                    try:
                        metadata = yaml.safe_load(metadata_yaml)
                        logger.info(f"Parsed metadata for {filename}: {list(metadata.keys()) if metadata else 'None'}")
                        return metadata, body
                    except yaml.YAMLError as e:
                        logger.warning(f"Invalid YAML in {filename}: {e}")
                        return None, content
            
            # No metadata found, return original content
            return None, content
            
        except Exception as e:
            logger.error(f"Error parsing metadata in {filename}: {e}")
            return None, content
    
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
    
    def _build_search_index(self):
        """Build inverted index for fast term lookups"""
        try:
            self.search_index.clear()
            
            for category, files in self.contexts.items():
                for filename, content in files.items():
                    # Extract searchable terms (words, phrases)
                    terms = self._extract_search_terms(content)
                    for term in terms:
                        self.search_index[term.lower()].add((category, filename))
            
            logger.info(f"Built search index with {len(self.search_index)} terms")
            
        except Exception as e:
            logger.error(f"Error building search index: {e}")
    
    def _extract_search_terms(self, content: str) -> List[str]:
        """Extract searchable terms from content using intelligent pattern detection"""
        terms = []
        
        # Extract headers (## Header, ### Header) - these are always important
        headers = re.findall(r'^#{1,6}\s+(.+)$', content, re.MULTILINE)
        terms.extend(headers)
        
        # Extract bold terms (**term**) - explicitly highlighted content
        bold_terms = re.findall(r'\*\*([^*]+)\*\*', content)
        terms.extend(bold_terms)
        
        # Extract level indicators (L1, L2, etc.)
        levels = re.findall(r'\bL\d+\b', content, re.IGNORECASE)
        terms.extend(levels)
        
        # Extract capitalized phrases (likely job titles, proper nouns)
        # Matches: "Software Engineer", "Data Analyst", "Machine Learning", etc.
        capitalized_phrases = re.findall(r'\b[A-Z][a-z]+(?:\s+[A-Z][a-z]+){0,3}\b', content)
        terms.extend(capitalized_phrases)
        
        # Extract words ending in common role suffixes
        role_suffixes = re.findall(r'\b\w+(?:Engineer|Manager|Analyst|Developer|Scientist|Director|Specialist|Coordinator|Administrator|Consultant|Architect)\b', content, re.IGNORECASE)
        terms.extend(role_suffixes)
        
        # Extract hyphenated terms (Full-Stack, CI/CD, etc.)
        hyphenated = re.findall(r'\b\w+[-/]\w+(?:[-/]\w+)*\b', content)
        terms.extend(hyphenated)
        
        # Extract acronyms (2-5 uppercase letters)
        acronyms = re.findall(r'\b[A-Z]{2,5}\b', content)
        terms.extend(acronyms)
        
        # Clean, filter, and deduplicate
        cleaned_terms = []
        stop_words = {'the', 'and', 'or', 'but', 'in', 'on', 'at', 'to', 'for', 'of', 'with', 'by', 'is', 'are', 'was', 'were', 'be', 'been', 'being', 'have', 'has', 'had', 'do', 'does', 'did', 'will', 'would', 'could', 'should', 'may', 'might', 'must', 'can', 'this', 'that', 'these', 'those'}
        
        for term in terms:
            cleaned = term.strip()
            # Filter: length > 2, not a stop word, not already included
            if (len(cleaned) > 2 and 
                cleaned.lower() not in stop_words and 
                cleaned not in cleaned_terms):
                cleaned_terms.append(cleaned)
        
        return cleaned_terms
    
    @lru_cache(maxsize=64)
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
    
    def fast_search(self, query: str, max_results: int = 5) -> List[Dict[str, Any]]:
        """Optimized search: metadata-first with content fallback"""
        try:
            query_lower = query.lower()
            results = []
            
            # Phase 1: Metadata search (if any files have metadata)
            if self.metadata_index:
                for filename, metadata in self.metadata_index.items():
                    if metadata:
                        # Check COMMON_SEARCHES field (highest priority)
                        common_searches = str(metadata.get('COMMON_SEARCHES', '')).lower()
                        if query_lower in common_searches:
                            category = self._find_category_for_file(filename)
                            if category:
                                results.append({
                                    'category': category,
                                    'file': filename,
                                    'match_type': 'metadata_exact',
                                    'score': 1.5
                                })
                        
                        # Check PURPOSE and KEY_SECTIONS (medium priority)
                        elif len(results) < max_results:
                            purpose = str(metadata.get('PURPOSE', '')).lower()
                            key_sections = str(metadata.get('KEY_SECTIONS', '')).lower()
                            if query_lower in purpose or query_lower in key_sections:
                                category = self._find_category_for_file(filename)
                                if category:
                                    results.append({
                                        'category': category,
                                        'file': filename,
                                        'match_type': 'metadata_partial',
                                        'score': 1.2
                                    })
                
                # If we found good metadata matches, return them (skip content search)
                if results and results[0]['score'] >= 1.2:
                    return sorted(results, key=lambda x: x['score'], reverse=True)[:max_results]
            
            # Phase 2: Content search (only if no metadata matches OR need more results)
            # Direct term matches
            if query_lower in self.search_index:
                for category, filename in self.search_index[query_lower]:
                    if not any(r['file'] == filename for r in results):
                        results.append({
                            'category': category,
                            'file': filename,
                            'match_type': 'content_exact',
                            'score': 1.0
                        })
            
            # Partial content matches (only if still need more results)
            if len(results) < max_results:
                for term, files in self.search_index.items():
                    if query_lower in term and query_lower != term:
                        for category, filename in files:
                            if not any(r['file'] == filename for r in results):
                                results.append({
                                    'category': category,
                                    'file': filename,
                                    'match_type': 'content_partial',
                                    'score': 0.7
                                })
                                if len(results) >= max_results:
                                    break
                    if len(results) >= max_results:
                        break
            
            return sorted(results, key=lambda x: x['score'], reverse=True)[:max_results]
            
        except Exception as e:
            logger.error(f"Error in optimized search: {e}")
            return []
    
    def _find_category_for_file(self, filename: str) -> Optional[str]:
        """Find which category a file belongs to"""
        for category, files in self.CONTEXT_CATEGORIES.items():
            if filename in files:
                return category
        return None
    
    def get_targeted_context(self, query: str, tool_name: str = None) -> str:
        """Get context targeted to specific query"""
        try:
            # Use fast search to find relevant files
            search_results = self.fast_search(query, max_results=3)
            
            if not search_results:
                # Fallback to tool-specific context
                return self.get_context_for_tool(tool_name) if tool_name else ""
            
            context_parts = []
            
            # Add index first for navigation
            if 'index' in self.contexts:
                for filename, content in self.contexts['index'].items():
                    context_parts.append(f"=== {filename} ===\n{content}\n")
            
            # Add targeted files
            for result in search_results:
                category = result['category']
                filename = result['file']
                if category in self.contexts and filename in self.contexts[category]:
                    content = self.contexts[category][filename]
                    context_parts.append(f"=== {filename} (Match: {result['match_type']}) ===\n{content}\n")
            
            return "\n".join(context_parts)
            
        except Exception as e:
            logger.error(f"Error getting targeted context: {e}")
            return self.get_context_for_tool(tool_name) if tool_name else ""
    
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
        self.file_hashes = {}
        self.search_index.clear()
        self.context_cache.clear()
        # Clear LRU cache
        self.get_context_for_tool.cache_clear()
        self._load_all_contexts()
        self._build_search_index()
    
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
