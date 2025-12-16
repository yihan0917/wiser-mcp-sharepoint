"""
AI Insights Validator for PowerPoint Generation
Validates AI insights structure against AI_INSIGHTS_TEMPLATE.md to prevent empty slides
"""
import os
from typing import Dict, List, Any, Tuple, Optional
from .common import logger

class InsightsValidator:
    """Validates AI insights structure against template requirements"""
    
    # Define required fields for each slide type based on AI_INSIGHTS_TEMPLATE.md
    SLIDE_TYPE_REQUIREMENTS = {
        'metric': {
            'required': ['type', 'value', 'label', 'context'],
            'optional': [],
            'description': 'Display a large key metric with supporting context'
        },
        'comparison': {
            'required': ['type', 'left', 'right'],
            'nested_required': {
                'left': ['title', 'points'],
                'right': ['title', 'points']
            },
            'optional': [],
            'description': 'Side-by-side comparison of two items'
        },
        'recommendation': {
            'required': ['type', 'recommendation', 'rationale'],
            'optional': [],
            'description': 'Actionable recommendation with supporting rationale'
        },
        'analysis': {
            'required': ['type', 'content'],
            'optional': ['footer'],
            'description': 'Detailed analysis with bullet points and optional footer'
        },
        'chart_with_analysis': {
            'required': ['type', 'chart_data', 'chart_type', 'analysis'],
            'nested_required': {
                'chart_data': ['labels', 'values']
            },
            'optional': [],
            'description': 'Custom chart with detailed insights'
        },
        'table': {
            'required': ['type', 'table_data', 'insights'],
            'optional': ['header_row'],
            'description': 'Data table with insights'
        },
        'two_column': {
            'required': ['type', 'sections'],
            'nested_required': {
                'sections': ['title', 'content']  # Each section must have these
            },
            'optional': [],
            'description': 'Flexible two-column layout for comparisons or action plans'
        }
    }
    
    def __init__(self, template_path: Optional[str] = None):
        """Initialize validator with optional template path"""
        if template_path is None:
            # Default to docs/AI_INSIGHTS_TEMPLATE.md relative to this file
            current_dir = os.path.dirname(os.path.abspath(__file__))
            project_root = os.path.dirname(os.path.dirname(current_dir))
            template_path = os.path.join(project_root, 'docs', 'AI_INSIGHTS_TEMPLATE.md')
        
        self.template_path = template_path
        self.template_exists = os.path.exists(template_path)
        
        if not self.template_exists:
            logger.warning(f"AI_INSIGHTS_TEMPLATE.md not found at {template_path}")
    
    def validate_insights(self, insights_data: List[Dict]) -> Tuple[bool, List[str], List[Dict]]:
        """
        Validate AI insights structure against template requirements
        
        Args:
            insights_data: List of insight dictionaries
            
        Returns:
            Tuple of (is_valid, error_messages, valid_insights)
        """
        if not isinstance(insights_data, list):
            return False, ["AI insights must be a list/array"], []
        
        errors = []
        valid_insights = []
        valid_types = list(self.SLIDE_TYPE_REQUIREMENTS.keys())
        
        for idx, insight in enumerate(insights_data):
            # Validate basic structure
            if not isinstance(insight, dict):
                errors.append(f"Insight at index {idx} must be a dictionary, got {type(insight).__name__}")
                continue
            
            # Check for required top-level fields
            if 'title' not in insight:
                errors.append(f"Insight at index {idx} missing required 'title' field")
                continue
            
            if 'data' not in insight:
                errors.append(f"Insight '{insight.get('title')}' missing required 'data' field")
                continue
            
            insight_title = insight.get('title', f'Insight {idx}')
            insight_data = insight.get('data', {})
            
            # Validate data is a dictionary
            if not isinstance(insight_data, dict):
                errors.append(f"Insight '{insight_title}': 'data' must be a dictionary, got {type(insight_data).__name__}")
                continue
            
            # Validate slide type
            slide_type = insight_data.get('type')
            if not slide_type:
                errors.append(f"Insight '{insight_title}': missing 'type' in data. Valid types: {', '.join(valid_types)}")
                continue
            
            if slide_type not in valid_types:
                errors.append(f"Insight '{insight_title}': invalid type '{slide_type}'. Valid types: {', '.join(valid_types)}")
                continue
            
            # Validate type-specific requirements
            type_errors = self._validate_slide_type(insight_title, slide_type, insight_data)
            if type_errors:
                errors.extend(type_errors)
                continue
            
            # If we got here, insight is valid
            valid_insights.append(insight)
        
        is_valid = len(errors) == 0
        return is_valid, errors, valid_insights
    
    def _validate_slide_type(self, title: str, slide_type: str, data: Dict) -> List[str]:
        """Validate data structure for specific slide type"""
        errors = []
        requirements = self.SLIDE_TYPE_REQUIREMENTS[slide_type]
        
        # Check required fields
        for field in requirements['required']:
            if field not in data:
                errors.append(
                    f"Insight '{title}' (type: {slide_type}): missing required field '{field}'. "
                    f"See AI_INSIGHTS_TEMPLATE.md for {slide_type} structure"
                )
            elif field != 'type':  # Don't check if 'type' is empty
                # Check if field has content
                value = data.get(field)
                if value is None or (isinstance(value, (list, dict, str)) and not value):
                    errors.append(
                        f"Insight '{title}' (type: {slide_type}): field '{field}' is empty. "
                        f"This will result in an empty slide"
                    )
        
        # Check nested requirements (for comparison, chart_with_analysis, two_column)
        if 'nested_required' in requirements:
            for parent_field, nested_fields in requirements['nested_required'].items():
                if parent_field in data:
                    parent_value = data[parent_field]
                    
                    # Handle sections array (two_column type)
                    if parent_field == 'sections' and isinstance(parent_value, list):
                        for section_idx, section in enumerate(parent_value):
                            if not isinstance(section, dict):
                                errors.append(
                                    f"Insight '{title}': sections[{section_idx}] must be a dictionary"
                                )
                                continue
                            for nested_field in nested_fields:
                                if nested_field not in section:
                                    errors.append(
                                        f"Insight '{title}': sections[{section_idx}] missing '{nested_field}'"
                                    )
                                elif not section[nested_field]:
                                    errors.append(
                                        f"Insight '{title}': sections[{section_idx}]['{nested_field}'] is empty"
                                    )
                    
                    # Handle left/right objects (comparison type)
                    elif isinstance(parent_value, dict):
                        for nested_field in nested_fields:
                            if nested_field not in parent_value:
                                errors.append(
                                    f"Insight '{title}': {parent_field}.{nested_field} is required. "
                                    f"See AI_INSIGHTS_TEMPLATE.md for {slide_type} structure"
                                )
                            elif not parent_value[nested_field]:
                                errors.append(
                                    f"Insight '{title}': {parent_field}.{nested_field} is empty"
                                )
        
        # Type-specific validation
        if slide_type == 'metric':
            # Ensure context is a list
            if 'context' in data and not isinstance(data['context'], list):
                errors.append(f"Insight '{title}': 'context' must be a list of strings")
        
        elif slide_type == 'analysis':
            # Ensure content is a list
            if 'content' in data and not isinstance(data['content'], list):
                errors.append(f"Insight '{title}': 'content' must be a list of strings")
        
        elif slide_type == 'recommendation':
            # Ensure rationale is a list
            if 'rationale' in data and not isinstance(data['rationale'], list):
                errors.append(f"Insight '{title}': 'rationale' must be a list of strings")
        
        elif slide_type == 'table':
            # Ensure table_data is a 2D list
            if 'table_data' in data:
                if not isinstance(data['table_data'], list):
                    errors.append(f"Insight '{title}': 'table_data' must be a 2D array")
                elif data['table_data'] and not isinstance(data['table_data'][0], list):
                    errors.append(f"Insight '{title}': 'table_data' must be a 2D array (array of arrays)")
            
            # Ensure insights is a list
            if 'insights' in data and not isinstance(data['insights'], list):
                errors.append(f"Insight '{title}': 'insights' must be a list of strings")
        
        elif slide_type == 'chart_with_analysis':
            # Validate chart_data structure
            if 'chart_data' in data:
                chart_data = data['chart_data']
                if not isinstance(chart_data, dict):
                    errors.append(f"Insight '{title}': 'chart_data' must be a dictionary")
                else:
                    if 'labels' in chart_data and 'values' in chart_data:
                        if len(chart_data['labels']) != len(chart_data['values']):
                            errors.append(
                                f"Insight '{title}': chart_data labels and values must have same length"
                            )
            
            # Ensure analysis is a list
            if 'analysis' in data and not isinstance(data['analysis'], list):
                errors.append(f"Insight '{title}': 'analysis' must be a list of strings")
        
        return errors
    
    def get_template_reference(self, slide_type: str) -> str:
        """Get template reference message for a specific slide type"""
        if slide_type not in self.SLIDE_TYPE_REQUIREMENTS:
            return f"Invalid slide type. Valid types: {', '.join(self.SLIDE_TYPE_REQUIREMENTS.keys())}"
        
        req = self.SLIDE_TYPE_REQUIREMENTS[slide_type]
        msg = f"\n{req['description']}\n"
        msg += f"Required fields: {', '.join(req['required'])}\n"
        if req.get('optional'):
            msg += f"Optional fields: {', '.join(req['optional'])}\n"
        if self.template_exists:
            msg += f"\nSee {self.template_path} for detailed examples"
        
        return msg
    
    def format_validation_report(self, errors: List[str], valid_count: int, total_count: int) -> str:
        """Format validation errors into a readable report"""
        if not errors:
            return f"✓ All {total_count} AI insights validated successfully"
        
        report = f"\n{'='*80}\n"
        report += f"AI INSIGHTS VALIDATION ERRORS ({len(errors)} errors found)\n"
        report += f"{'='*80}\n\n"
        
        for idx, error in enumerate(errors, 1):
            report += f"{idx}. {error}\n"
        
        report += f"\n{'='*80}\n"
        report += f"Valid insights: {valid_count}/{total_count}\n"
        
        if self.template_exists:
            report += f"\nPlease review {self.template_path} for correct structure\n"
        
        report += f"{'='*80}\n"
        
        return report

# Global instance
insights_validator = InsightsValidator()
