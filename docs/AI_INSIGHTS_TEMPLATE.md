# AI Insights JSON Template for Create_PowerPoint_Report

## Overview

When using the `Create_PowerPoint_Report` tool with the `ai_insights` parameter, you must provide a JSON array (or JSON string) following this template structure.

**IMPORTANT:** The MCP framework automatically parses JSON arrays, so you can pass either:
- A JSON array directly: `[{...}, {...}]`
- A JSON string: `"[{...}, {...}]"`

Both formats are accepted and will work correctly.

---

## Template Structure

```json
[
  {
    "title": "Slide Title Here",
    "data": {
      "type": "metric|comparison|recommendation|analysis|chart_with_analysis|table|two_column",
      ... type-specific fields ...
    }
  }
]
```

---

## Slide Type Templates

### 1. Metric Slide
Display a large key metric with supporting context.

```json
{
  "title": "Key Metric Title",
  "data": {
    "type": "metric",
    "value": "76 days",
    "label": "Metric description or subtitle",
    "context": [
      "Supporting point 1",
      "Supporting point 2",
      "Supporting point 3"
    ]
  }
}
```

**Example:**
```json
{
  "title": "Average Time-to-Hire",
  "data": {
    "type": "metric",
    "value": "76 days",
    "label": "Q3 2023 Engineering Roles",
    "context": [
      "13% increase from Q2 2023",
      "Primary driver: Interview coordination delays",
      "Target: Reduce to 60 days by Q4"
    ]
  }
}
```

---

### 2. Comparison Slide
Side-by-side comparison of two items.

```json
{
  "title": "Comparison Title",
  "data": {
    "type": "comparison",
    "left": {
      "title": "Option A Title",
      "points": [
        "Point 1",
        "Point 2",
        "Point 3"
      ]
    },
    "right": {
      "title": "Option B Title",
      "points": [
        "Point 1",
        "Point 2",
        "Point 3"
      ]
    }
  }
}
```

**Example:**
```json
{
  "title": "US vs India Hiring Performance",
  "data": {
    "type": "comparison",
    "left": {
      "title": "United States",
      "points": [
        "Average time: 72 days",
        "4 positions filled",
        "Strong technical talent pool",
        "Challenge: Higher compensation"
      ]
    },
    "right": {
      "title": "India",
      "points": [
        "Average time: 68 days",
        "5 positions filled",
        "Excellent technical skills",
        "Advantage: Cost-effective"
      ]
    }
  }
}
```

---

### 3. Recommendation Slide
Actionable recommendation with supporting rationale.

```json
{
  "title": "Recommendation Title",
  "data": {
    "type": "recommendation",
    "recommendation": "Clear, actionable recommendation statement",
    "rationale": [
      "Reason or supporting point 1",
      "Reason or supporting point 2",
      "Expected impact or outcome"
    ]
  }
}
```

**Example:**
```json
{
  "title": "Recommendation: Streamline Interview Scheduling",
  "data": {
    "type": "recommendation",
    "recommendation": "Implement dedicated interview blocks for hiring managers to reduce stage from 12.7 to 7 days",
    "rationale": [
      "**Context:** Multi-location hiring requires coordinated scheduling",
      "**Current Impact:** 12.7 days average (longest bottleneck)",
      "**Solution:** Create shared calendar blocks",
      "**Expected Outcome:** 21% reduction in time-to-hire"
    ]
  }
}
```

---

### 4. Analysis Slide
Detailed analysis with bullet points and optional footer.

```json
{
  "title": "Analysis Title",
  "data": {
    "type": "analysis",
    "content": [
      "Analysis point 1",
      "Analysis point 2",
      "",
      "**Bold point for emphasis**",
      "Analysis point 3"
    ],
    "footer": "Optional footer note or context"
  }
}
```

**Example:**
```json
{
  "title": "Bottleneck Analysis",
  "data": {
    "type": "analysis",
    "content": [
      "**Hiring Manager Interview: 12.7 days** (Longest stage)",
      "**Technical Interview: 9.9 days** (Second longest)",
      "**Final Interview: 10.6 days** (Panel coordination)",
      "",
      "**Total interview time: 33 days** (43% of hiring cycle)",
      "",
      "**Root Cause:** Multi-location coordination gaps",
      "**Impact:** Risk of losing top candidates"
    ],
    "footer": "Based on analysis of 97 positions in Q3 2023"
  }
}
```

---

### 5. Chart with Analysis Slide
Custom chart with detailed insights.

```json
{
  "title": "Chart Title",
  "data": {
    "type": "chart_with_analysis",
    "chart_data": {
      "labels": ["Label 1", "Label 2", "Label 3"],
      "values": [10, 20, 15]
    },
    "chart_type": "bar",
    "analysis": [
      "Insight 1 about the data",
      "Insight 2 about the data",
      "Insight 3 about the data"
    ]
  }
}
```

**Example:**
```json
{
  "title": "Location Efficiency Analysis",
  "data": {
    "type": "chart_with_analysis",
    "chart_data": {
      "labels": ["US", "Canada", "India"],
      "values": [72, 70, 68]
    },
    "chart_type": "bar",
    "analysis": [
      "**India:** Most efficient (68 days average)",
      "**Canada:** Strong performance (70 days)",
      "**US:** Slight delay (72 days)",
      "",
      "**Opportunity:** Share India best practices with US team"
    ]
  }
}
```

---

### 6. Table Slide
Data table with insights.

```json
{
  "title": "Table Title",
  "data": {
    "type": "table",
    "table_data": [
      ["Header 1", "Header 2", "Header 3"],
      ["Row 1 Col 1", "Row 1 Col 2", "Row 1 Col 3"],
      ["Row 2 Col 1", "Row 2 Col 2", "Row 2 Col 3"]
    ],
    "header_row": true,
    "insights": [
      "Key insight 1",
      "Key insight 2"
    ]
  }
}
```

**Example:**
```json
{
  "title": "Role Performance Summary",
  "data": {
    "type": "table",
    "table_data": [
      ["Role", "Positions", "Avg Days", "Status"],
      ["Software Engineer", "8", "76", "On Track"],
      ["Data Analyst", "5", "74", "Excellent"],
      ["DevOps Engineer", "3", "82", "Needs Attention"]
    ],
    "header_row": true,
    "insights": [
      "Software Engineer and Data Analyst performing well",
      "DevOps Engineer role requires process optimization"
    ]
  }
}
```

---

### 7. Two-Column Slide
Flexible two-column layout for comparisons or action plans.

```json
{
  "title": "Two-Column Title",
  "data": {
    "type": "two_column",
    "sections": [
      {
        "title": "Left Column Title",
        "content": [
          "Point 1",
          "Point 2",
          "Point 3"
        ]
      },
      {
        "title": "Right Column Title",
        "content": [
          "Point 1",
          "Point 2",
          "Point 3"
        ]
      }
    ]
  }
}
```

**Example:**
```json
{
  "title": "Q4 Action Plan",
  "data": {
    "type": "two_column",
    "sections": [
      {
        "title": "Immediate Actions (30 Days)",
        "content": [
          "Establish shared interview calendar blocks",
          "Deploy async technical assessments",
          "Standardize interview templates by role",
          "Assign regional recruiting coordinators"
        ]
      },
      {
        "title": "Long-Term Initiatives (90 Days)",
        "content": [
          "Build multi-location hiring playbook",
          "Align assessments with career path",
          "Create interview training program",
          "Implement predictive analytics"
        ]
      }
    ]
  }
}
```

---

## Complete Example

Here's a complete example with multiple slide types:

```json
[
  {
    "title": "Executive Summary",
    "data": {
      "type": "metric",
      "value": "76 days",
      "label": "Average Time-to-Hire Q3 2023",
      "context": [
        "97 positions filled",
        "100% fill rate achieved",
        "13% increase from Q2"
      ]
    }
  },
  {
    "title": "Key Findings",
    "data": {
      "type": "analysis",
      "content": [
        "**Hiring Manager Interview: 12.7 days** (Longest stage)",
        "**Technical Interview: 9.9 days** (Second longest)",
        "",
        "**Total interview time: 33 days** (43% of cycle)",
        "**Root Cause:** Multi-location coordination gaps"
      ],
      "footer": "Based on Q3 2023 data analysis"
    }
  },
  {
    "title": "Strategic Recommendation",
    "data": {
      "type": "recommendation",
      "recommendation": "Implement dedicated interview scheduling blocks",
      "rationale": [
        "Current bottleneck: 12.7 days for hiring manager interviews",
        "Solution: Shared calendar blocks across timezones",
        "Expected impact: 21% reduction in time-to-hire"
      ]
    }
  },
  {
    "title": "Action Plan",
    "data": {
      "type": "two_column",
      "sections": [
        {
          "title": "Immediate (30 Days)",
          "content": [
            "Create shared interview calendars",
            "Deploy async assessments",
            "Standardize templates"
          ]
        },
        {
          "title": "Long-Term (90 Days)",
          "content": [
            "Build hiring playbook",
            "Implement analytics dashboard",
            "Create training program"
          ]
        }
      ]
    }
  }
]
```

---

## Best Practices

### 1. Keep It Concise
- Limit bullet points to 3-5 per slide
- Use clear, actionable language
- Focus on insights, not data repetition

### 2. Use Formatting
- Use `**bold**` for emphasis
- Use empty strings `""` for spacing between sections
- Keep labels and titles short and descriptive

### 3. Structure Your Story
- Start with key metrics or executive summary
- Follow with analysis
- End with recommendations and action plans

### 4. Mix Slide Types
- Don't use all the same type
- Combine metrics, analysis, and recommendations
- Use two-column for action plans

### 5. Context Matters
- Add footer notes for data sources or context
- Include relevant business context in rationale
- Reference company-specific factors when applicable

---

## Common Mistakes to Avoid

❌ **Wrong:** Missing required fields
```json
{
  "title": "My Slide",
  "data": {
    "type": "metric"
    // Missing: value, label, context
  }
}
```

✅ **Correct:** All required fields included
```json
{
  "title": "My Slide",
  "data": {
    "type": "metric",
    "value": "76 days",
    "label": "Average Time",
    "context": ["Point 1", "Point 2"]
  }
}
```

---

❌ **Wrong:** Invalid type
```json
{
  "title": "My Slide",
  "data": {
    "type": "summary"  // Not a valid type
  }
}
```

✅ **Correct:** Valid type from the 7 supported types
```json
{
  "title": "My Slide",
  "data": {
    "type": "analysis",  // Valid type
    "content": ["Point 1", "Point 2"]
  }
}
```

---

❌ **Wrong:** Incorrect data structure
```json
{
  "title": "Comparison",
  "data": {
    "type": "comparison",
    "left": "Option A",  // Should be an object
    "right": "Option B"
  }
}
```

✅ **Correct:** Proper nested structure
```json
{
  "title": "Comparison",
  "data": {
    "type": "comparison",
    "left": {
      "title": "Option A",
      "points": ["Point 1", "Point 2"]
    },
    "right": {
      "title": "Option B",
      "points": ["Point 1", "Point 2"]
    }
  }
}
```

---

## Quick Reference

| Slide Type | Required Fields | Optional Fields |
|------------|----------------|-----------------|
| `metric` | type, value, label, context | - |
| `comparison` | type, left (title, points), right (title, points) | - |
| `recommendation` | type, recommendation, rationale | - |
| `analysis` | type, content | footer |
| `chart_with_analysis` | type, chart_data, chart_type, analysis | - |
| `table` | type, table_data, insights | header_row |
| `two_column` | type, sections (title, content) | - |

---

## Need Help?

- See usage examples in `SETUP_GUIDE.md` (Create_PowerPoint_Report Enhancement section)
- Review this template before creating AI insights
- Check the Quick Reference table above for field requirements
