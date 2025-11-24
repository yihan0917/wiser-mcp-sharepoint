# PowerPoint AI Generation - JSON Template Guide

This document provides the complete JSON structure templates for creating PowerPoint presentations using the `Generate_AI_PowerPoint` tool.

## Base Structure

```json
{
  "presentation_title": "Your Presentation Title",
  "output_filename": "optional_filename.pptx",
  "output_folder": "optional_folder_name",
  "slides": [
    // Array of slide objects (see templates below)
  ]
}
```

## Slide Object Base Structure

Each slide must have:
```json
{
  "title": "Slide Title",
  "data": {
    "type": "slide_type",
    // Type-specific data fields
  }
}
```

## Slide Type Templates

### 1. Metric Slide (`type: "metric"`)
Displays a large metric value with supporting context.

```json
{
  "title": "Key Performance Metric",
  "data": {
    "type": "metric",
    "value": "76 Days",           // Main metric value (string or number)
    "label": "Average Time to Hire",  // Label for the metric
    "context": [                 // Optional: Array of supporting bullet points
      "Total records analyzed: 97",
      "Fill rate: 100%",
      "Date range: Q3 2023 - Q1 2024"
    ]
  }
}
```

### 2. Analysis Slide (`type: "analysis"`)
Displays bullet points for detailed analysis.

```json
{
  "title": "Detailed Analysis",
  "data": {
    "type": "analysis",
    "content": [                 // Array of bullet points
      "First key insight or finding",
      "Second important analysis point",
      "Third conclusion or recommendation"
    ],
    "footer": "Optional footer note"  // Optional: Small footer text
  }
}
```

### 3. Comparison Slide (`type: "comparison"`)
Side-by-side comparison with two columns.

```json
{
  "title": "Before vs After Comparison",
  "data": {
    "type": "comparison",
    "left": {
      "title": "Current State",   // Left column title
      "points": [                // Left column bullet points
        "Current metric: 76 days",
        "Process bottlenecks: 3",
        "Efficiency: Moderate"
      ]
    },
    "right": {
      "title": "Target State",    // Right column title
      "points": [                // Right column bullet points
        "Target metric: 55 days",
        "Process bottlenecks: 0-1",
        "Efficiency: High"
      ]
    }
  }
}
```

### 4. Two-Column Slide (`type: "two_column"`)
General two-column layout with sections.

```json
{
  "title": "Two-Column Information",
  "data": {
    "type": "two_column",
    "sections": [               // Array of exactly 2 sections
      {
        "title": "Left Section Title",
        "content": [            // Left section bullet points
          "First point in left column",
          "Second point in left column"
        ]
      },
      {
        "title": "Right Section Title", 
        "content": [            // Right section bullet points
          "First point in right column",
          "Second point in right column"
        ]
      }
    ]
  }
}
```

### 5. Recommendation Slide (`type: "recommendation"`)
Actionable recommendation with rationale.

```json
{
  "title": "Strategic Recommendation",
  "data": {
    "type": "recommendation",
    "recommendation": "Implement automated scheduling system to reduce hiring time by 28%",
    "rationale": [              // Array of rationale bullet points
      "Current scheduling takes 12.7 days on average",
      "Automated system could reduce to 8 days",
      "ROI expected within Q1 2024"
    ]
  }
}
```

### 6. Table Slide (`type: "table"`)
Data table with optional insights.

```json
{
  "title": "Data Summary Table",
  "data": {
    "type": "table",
    "table_data": [             // Array of arrays (rows and columns)
      ["Header 1", "Header 2", "Header 3"],  // First row = headers
      ["Row 1 Col 1", "Row 1 Col 2", "Row 1 Col 3"],
      ["Row 2 Col 1", "Row 2 Col 2", "Row 2 Col 3"]
    ],
    "header_row": true,         // Optional: Whether first row is header (default: true)
    "insights": [               // Optional: Bullet points below table
      "Key insight from the data",
      "Important pattern or trend"
    ]
  }
}
```

### 7. Chart with Analysis (`type: "chart_with_analysis"`)
Chart on left, analysis bullet points on right.

```json
{
  "title": "Performance Chart Analysis",
  "data": {
    "type": "chart_with_analysis",
    "chart_data": {
      "labels": ["Q1", "Q2", "Q3", "Q4"],    // X-axis labels
      "values": [65, 70, 76, 72],            // Y-axis values
      "chart_title": "Hiring Time Trends"    // Optional chart title
    },
    "chart_type": "bar",        // Options: "bar", "line", "pie"
    "analysis": [               // Right-side bullet points
      "Q3 shows peak hiring time of 76 days",
      "Improvement needed in Q4",
      "Target: Reduce to 55 days average"
    ]
  }
}
```

## Complete Example

```json
{
  "presentation_title": "Q3 Hiring Analysis Report",
  "output_filename": "Q3_Analysis.pptx",
  "output_folder": "Reports",
  "slides": [
    {
      "title": "Executive Summary",
      "data": {
        "type": "metric",
        "value": "76 Days",
        "label": "Average Time to Hire",
        "context": [
          "97 positions analyzed",
          "100% fill rate achieved",
          "Q3 2023 - Q1 2024 data"
        ]
      }
    },
    {
      "title": "Key Findings",
      "data": {
        "type": "analysis",
        "content": [
          "Hiring Manager Interview stage takes longest (12.7 days)",
          "Final Interview coordination needs improvement (10.6 days)",
          "Manager review process has delays (9.1 days)",
          "Combined bottlenecks account for 43% of total time"
        ]
      }
    },
    {
      "title": "Current vs Target Performance",
      "data": {
        "type": "comparison",
        "left": {
          "title": "Current State",
          "points": [
            "Average: 76 days",
            "Bottlenecks: 3 major",
            "Consistency: Moderate"
          ]
        },
        "right": {
          "title": "Target State", 
          "points": [
            "Target: 55 days",
            "Bottlenecks: 0-1 minor",
            "Consistency: High"
          ]
        }
      }
    }
  ]
}
```

## Important Notes

1. **Required Fields**: Every slide must have `title` and `data` with `type`
2. **Data Structure**: The `data` object structure varies by slide type
3. **Arrays**: Use arrays for bullet points, table rows, chart data
4. **Strings**: All text content should be strings
5. **Optional Fields**: Many fields are optional and will be skipped if not provided
6. **Validation**: Invalid JSON or missing required fields will cause errors

## Field Reference Quick Guide

| Slide Type | Required Fields | Optional Fields |
|------------|----------------|-----------------|
| `metric` | `type`, `value`, `label` | `context` |
| `analysis` | `type`, `content` | `footer` |
| `comparison` | `type`, `left`, `right` | - |
| `two_column` | `type`, `sections` | - |
| `recommendation` | `type`, `recommendation` | `rationale` |
| `table` | `type`, `table_data` | `header_row`, `insights` |
| `chart_with_analysis` | `type`, `chart_data`, `analysis` | `chart_type` |
