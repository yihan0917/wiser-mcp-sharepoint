"""
PowerPoint helper for creating professional presentation reports
"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.shapes import MSO_SHAPE
import pandas as pd
from datetime import datetime
from typing import Dict, List, Optional
from .common import logger
import io

class PowerPointHelper:
    """Helper class for creating professional PowerPoint presentations"""
    
    # Professional color scheme
    COLORS = {
        'primary': RGBColor(0, 112, 192),      # Professional blue
        'secondary': RGBColor(68, 114, 196),   # Light blue
        'accent': RGBColor(237, 125, 49),      # Orange
        'success': RGBColor(112, 173, 71),     # Green
        'header_bg': RGBColor(68, 114, 196),   # Header background
        'text': RGBColor(64, 64, 64),          # Dark gray text
        'light_gray': RGBColor(217, 217, 217), # Light gray
    }
    
    def __init__(self):
        self.prs = Presentation()
        self.prs.slide_width = Inches(10)
        self.prs.slide_height = Inches(7.5)
        
    def create_title_slide(self, title: str, subtitle: str = None):
        """Create professional title slide with centered title and date"""
        slide_layout = self.prs.slide_layouts[6]  # Blank layout
        slide = self.prs.slides.add_slide(slide_layout)
        
        # Add title - centered
        title_box = slide.shapes.add_textbox(
            Inches(0.5), Inches(2.5), Inches(9), Inches(2)
        )
        title_frame = title_box.text_frame
        title_frame.text = title
        title_frame.word_wrap = True
        title_para = title_frame.paragraphs[0]
        title_para.alignment = PP_ALIGN.CENTER
        title_para.font.size = Pt(36)
        title_para.font.bold = True
        title_para.font.color.rgb = self.COLORS['primary']
        
        # Add subtitle/date - centered
        if subtitle is None:
            subtitle = datetime.now().strftime("%B %d, %Y")
        
        subtitle_box = slide.shapes.add_textbox(
            Inches(1), Inches(4.5), Inches(8), Inches(0.8)
        )
        subtitle_frame = subtitle_box.text_frame
        subtitle_frame.text = subtitle
        subtitle_para = subtitle_frame.paragraphs[0]
        subtitle_para.alignment = PP_ALIGN.CENTER
        subtitle_para.font.size = Pt(24)
        subtitle_para.font.color.rgb = self.COLORS['text']
        
        return slide
    
    def create_content_slide(self, title: str):
        """Create slide with professional header"""
        slide_layout = self.prs.slide_layouts[6]  # Blank layout
        slide = self.prs.slides.add_slide(slide_layout)
        
        # Add header background with color
        header_shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(0), Inches(0), Inches(10), Inches(0.8)
        )
        header_shape.fill.solid()
        header_shape.fill.fore_color.rgb = self.COLORS['header_bg']
        header_shape.line.fill.background()
        
        # Add title text in header
        title_box = slide.shapes.add_textbox(
            Inches(0.5), Inches(0.15), Inches(9), Inches(0.5)
        )
        title_frame = title_box.text_frame
        title_frame.text = title
        title_para = title_frame.paragraphs[0]
        title_para.font.size = Pt(28)
        title_para.font.bold = True
        title_para.font.color.rgb = RGBColor(255, 255, 255)  # White text
        
        return slide
    
    def add_chart_to_slide(self, slide, chart_data: Dict, chart_type: str, 
                          left: float, top: float, width: float, height: float):
        """Add chart to slide with proper formatting"""
        try:
            labels = chart_data.get('labels', [])
            values = chart_data.get('values', [])
            
            if not labels or not values:
                return None
            
            # Create chart data
            chart_data_obj = CategoryChartData()
            chart_data_obj.categories = labels
            chart_data_obj.add_series('', values)
            
            # Determine chart type
            if chart_type in ['source_effectiveness', 'department_hiring']:
                xl_chart_type = XL_CHART_TYPE.PIE
            elif chart_type == 'hiring_trends':
                xl_chart_type = XL_CHART_TYPE.LINE
            else:
                xl_chart_type = XL_CHART_TYPE.COLUMN_CLUSTERED
            
            # Add chart to slide
            chart = slide.shapes.add_chart(
                xl_chart_type, Inches(left), Inches(top),
                Inches(width), Inches(height), chart_data_obj
            ).chart
            
            # Format chart
            chart.has_legend = True
            chart.legend.position = 4  # Right side - prevents overlap with chart area
            chart.legend.font.size = Pt(8)
            chart.legend.include_in_layout = False
            
            # Format category axis (bar labels) - reduce font size significantly
            if hasattr(chart, 'category_axis'):
                chart.category_axis.tick_labels.font.size = Pt(7)
                # Rotate labels if too many categories
                if len(labels) > 8:
                    chart.category_axis.tick_labels.orientation = -45
            
            # Format value axis
            if hasattr(chart, 'value_axis'):
                chart.value_axis.tick_labels.font.size = Pt(8)
            
            return chart
            
        except Exception as e:
            logger.error(f"Error adding chart to slide: {e}")
            return None
    
    def add_bullet_points(self, slide, insights: List[str], 
                         left: float, top: float, width: float, height: float):
        """Add bullet points to slide"""
        text_box = slide.shapes.add_textbox(
            Inches(left), Inches(top), Inches(width), Inches(height)
        )
        text_frame = text_box.text_frame
        text_frame.word_wrap = True
        
        for i, insight in enumerate(insights):
            if i == 0:
                p = text_frame.paragraphs[0]
            else:
                p = text_frame.add_paragraph()
            
            p.text = insight
            p.level = 0
            p.font.size = Pt(14)
            p.font.color.rgb = self.COLORS['text']
            p.space_before = Pt(6)
            p.space_after = Pt(6)
        
        return text_box
    
    def add_table(self, slide, data: List[List[str]], 
                 left: float, top: float, width: float, height: float,
                 header_row: bool = True):
        """Add formatted table to slide"""
        rows = len(data)
        cols = len(data[0]) if data else 0
        
        if rows == 0 or cols == 0:
            return None
        
        table = slide.shapes.add_table(
            rows, cols, Inches(left), Inches(top),
            Inches(width), Inches(height)
        ).table
        
        # Populate table
        for i, row_data in enumerate(data):
            for j, cell_value in enumerate(row_data):
                cell = table.cell(i, j)
                cell.text = str(cell_value)
                cell.text_frame.paragraphs[0].font.size = Pt(11)
                
                # Format header row
                if i == 0 and header_row:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = self.COLORS['header_bg']
                    cell.text_frame.paragraphs[0].font.bold = True
                    cell.text_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
        
        return table
    
    def create_chart_slide(self, title: str, chart_data: Dict, chart_type: str, 
                          insights: List[str]):
        """Create slide with chart and insights"""
        slide = self.create_content_slide(title)
        
        # Add chart on left side with more space for legend
        self.add_chart_to_slide(
            slide, chart_data, chart_type,
            left=0.5, top=1.2, width=5.0, height=4.8
        )
        
        # Add insights on right side (avoid overlap)
        if insights:
            # Add "Key Insights" label
            label_box = slide.shapes.add_textbox(
                Inches(6.5), Inches(1.2), Inches(3), Inches(0.4)
            )
            label_frame = label_box.text_frame
            label_frame.text = "Key Insights"
            label_para = label_frame.paragraphs[0]
            label_para.font.size = Pt(18)
            label_para.font.bold = True
            label_para.font.color.rgb = self.COLORS['primary']
            
            # Add bullet points
            self.add_bullet_points(
                slide, insights,
                left=6.5, top=1.8, width=3, height=4
            )
        
        return slide
    
    def create_data_definitions_slide(self, columns: Dict[str, str]):
        """Create slide explaining data column definitions"""
        slide = self.create_content_slide("Data Column Definitions")
        
        # Convert to table format
        table_data = [["Column Name", "Definition"]]
        for col_name, definition in columns.items():
            table_data.append([col_name, definition])
        
        # Add table
        self.add_table(
            slide, table_data,
            left=0.5, top=1.2, width=9, height=5.5,
            header_row=True
        )
        
        return slide
    
    def save_to_bytes(self) -> bytes:
        """Save presentation to bytes"""
        pptx_buffer = io.BytesIO()
        self.prs.save(pptx_buffer)
        pptx_buffer.seek(0)
        return pptx_buffer.getvalue()

# Global instance
powerpoint_helper = PowerPointHelper()
