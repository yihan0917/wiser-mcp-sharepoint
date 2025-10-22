"""
Visualization helper for creating Excel charts
"""
import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import PieChart, BarChart, LineChart, Reference
from openpyxl.utils.dataframe import dataframe_to_rows
import io
import tempfile
from typing import Dict, List
from .common import logger

class VisualizationHelper:
    """Helper class for creating Excel files with embedded charts"""
    
    def __init__(self):
        self.temp_dir = tempfile.gettempdir()
    
    def create_excel_with_charts(self, data: pd.DataFrame, chart_configs: List[Dict], file_name: str) -> bytes:
        """Create Excel file with embedded charts and data"""
        try:
            wb = Workbook()
            
            # Create data sheet
            ws_data = wb.active
            ws_data.title = "Raw Data"
            
            # Add data to worksheet
            for r in dataframe_to_rows(data, index=False, header=True):
                ws_data.append(r)
            
            # Create charts sheet
            ws_charts = wb.create_sheet("Charts & Analysis")
            
            chart_row = 1
            
            chart_index = 0
            for config in chart_configs:
                chart_data = config.get('data', {})
                chart_type = config.get('chart_type', 'bar')
                
                if not chart_data.get('labels') or not chart_data.get('values'):
                    continue
                
                # Add chart title
                ws_charts.cell(row=chart_row, column=1, value=chart_data.get('title', 'Chart'))
                chart_row += 2
                
                # Add data for this chart
                labels = chart_data['labels']
                values = chart_data['values']
                
                # Write labels and values
                for i, (label, value) in enumerate(zip(labels, values)):
                    ws_charts.cell(row=chart_row + i, column=1, value=label)
                    ws_charts.cell(row=chart_row + i, column=2, value=value)
                
                # Create completely new chart instance with unique properties
                try:
                    if chart_type in ['source_effectiveness', 'department_hiring']:
                        chart = PieChart()
                        chart.title = f"{chart_data.get('title', 'Chart')} ({chart_index})"
                        
                        # Define data range
                        data_range = Reference(ws_charts, 
                                             min_col=2, max_col=2,
                                             min_row=chart_row, max_row=chart_row + len(values) - 1)
                        labels_range = Reference(ws_charts,
                                               min_col=1, max_col=1, 
                                               min_row=chart_row, max_row=chart_row + len(labels) - 1)
                        
                        chart.add_data(data_range)
                        chart.set_categories(labels_range)
                        
                    elif chart_type in ['hiring_trends', 'time_to_hire_distribution']:
                        if chart_type == 'hiring_trends':
                            chart = LineChart()
                        else:
                            chart = BarChart()
                        
                        chart.title = f"{chart_data.get('title', 'Chart')} ({chart_index})"
                        chart.x_axis.title = "Categories"
                        chart.y_axis.title = "Count"
                        
                        data_range = Reference(ws_charts,
                                             min_col=2, max_col=2,
                                             min_row=chart_row, max_row=chart_row + len(values) - 1)
                        labels_range = Reference(ws_charts,
                                               min_col=1, max_col=1,
                                               min_row=chart_row, max_row=chart_row + len(labels) - 1)
                        
                        chart.add_data(data_range, titles_from_data=False)
                        chart.set_categories(labels_range)
                    else:
                        chart = None
                    
                    # Position chart only if chart was created
                    if chart:
                        chart_cell = f"D{chart_row}"
                        ws_charts.add_chart(chart, chart_cell)
                        chart_index += 1
                        
                except Exception as e:
                    logger.error(f"Error creating chart {chart_index}: {e}")
                
                chart_row += len(values) + 10  # Space between charts
            
            # Save to bytes
            excel_buffer = io.BytesIO()
            wb.save(excel_buffer)
            excel_buffer.seek(0)
            
            return excel_buffer.getvalue()
            
        except Exception as e:
            logger.error(f"Error creating Excel with charts: {e}")
            raise

# Global instance
visualization_helper = VisualizationHelper()
