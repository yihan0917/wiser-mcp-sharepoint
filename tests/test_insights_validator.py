"""
Test suite for AI Insights Validator
Tests validation logic against AI_INSIGHTS_TEMPLATE.md requirements
"""
import sys
import os

# Add parent directory to path for imports
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from src.mcp_sharepoint.insights_validator import insights_validator

def test_valid_metric_slide():
    """Test valid metric slide structure"""
    insights = [{
        "title": "Average Time-to-Hire",
        "data": {
            "type": "metric",
            "value": "76 days",
            "label": "Q3 2023 Engineering Roles",
            "context": [
                "13% increase from Q2 2023",
                "Primary driver: Interview coordination delays"
            ]
        }
    }]
    
    is_valid, errors, valid_insights = insights_validator.validate_insights(insights)
    print(f"\n✓ Valid Metric Slide Test:")
    print(f"  Valid: {is_valid}")
    print(f"  Errors: {errors}")
    print(f"  Valid insights: {len(valid_insights)}")
    assert is_valid, "Valid metric slide should pass validation"
    assert len(valid_insights) == 1

def test_invalid_metric_slide_missing_fields():
    """Test metric slide with missing required fields"""
    insights = [{
        "title": "Average Time-to-Hire",
        "data": {
            "type": "metric",
            "value": "76 days"
            # Missing: label, context
        }
    }]
    
    is_valid, errors, valid_insights = insights_validator.validate_insights(insights)
    print(f"\n✗ Invalid Metric Slide Test (missing fields):")
    print(f"  Valid: {is_valid}")
    print(f"  Errors: {errors}")
    print(f"  Valid insights: {len(valid_insights)}")
    assert not is_valid, "Metric slide missing required fields should fail"
    assert len(errors) > 0
    assert len(valid_insights) == 0

def test_valid_analysis_slide():
    """Test valid analysis slide structure"""
    insights = [{
        "title": "Bottleneck Analysis",
        "data": {
            "type": "analysis",
            "content": [
                "**Hiring Manager Interview: 12.7 days** (Longest stage)",
                "**Technical Interview: 9.9 days** (Second longest)",
                "",
                "**Root Cause:** Multi-location coordination gaps"
            ],
            "footer": "Based on Q3 2023 data"
        }
    }]
    
    is_valid, errors, valid_insights = insights_validator.validate_insights(insights)
    print(f"\n✓ Valid Analysis Slide Test:")
    print(f"  Valid: {is_valid}")
    print(f"  Errors: {errors}")
    print(f"  Valid insights: {len(valid_insights)}")
    assert is_valid, "Valid analysis slide should pass validation"
    assert len(valid_insights) == 1

def test_invalid_analysis_slide_empty_content():
    """Test analysis slide with empty content"""
    insights = [{
        "title": "Bottleneck Analysis",
        "data": {
            "type": "analysis",
            "content": []  # Empty content will result in empty slide
        }
    }]
    
    is_valid, errors, valid_insights = insights_validator.validate_insights(insights)
    print(f"\n✗ Invalid Analysis Slide Test (empty content):")
    print(f"  Valid: {is_valid}")
    print(f"  Errors: {errors}")
    print(f"  Valid insights: {len(valid_insights)}")
    assert not is_valid, "Analysis slide with empty content should fail"
    assert len(errors) > 0

def test_valid_recommendation_slide():
    """Test valid recommendation slide structure"""
    insights = [{
        "title": "Recommendation: Streamline Interviews",
        "data": {
            "type": "recommendation",
            "recommendation": "Implement dedicated interview blocks for hiring managers",
            "rationale": [
                "Current bottleneck: 12.7 days average",
                "Solution: Shared calendar blocks",
                "Expected impact: 21% reduction"
            ]
        }
    }]
    
    is_valid, errors, valid_insights = insights_validator.validate_insights(insights)
    print(f"\n✓ Valid Recommendation Slide Test:")
    print(f"  Valid: {is_valid}")
    print(f"  Errors: {errors}")
    print(f"  Valid insights: {len(valid_insights)}")
    assert is_valid, "Valid recommendation slide should pass validation"
    assert len(valid_insights) == 1

def test_valid_comparison_slide():
    """Test valid comparison slide structure"""
    insights = [{
        "title": "US vs India Hiring Performance",
        "data": {
            "type": "comparison",
            "left": {
                "title": "United States",
                "points": [
                    "Average time: 72 days",
                    "4 positions filled"
                ]
            },
            "right": {
                "title": "India",
                "points": [
                    "Average time: 68 days",
                    "5 positions filled"
                ]
            }
        }
    }]
    
    is_valid, errors, valid_insights = insights_validator.validate_insights(insights)
    print(f"\n✓ Valid Comparison Slide Test:")
    print(f"  Valid: {is_valid}")
    print(f"  Errors: {errors}")
    print(f"  Valid insights: {len(valid_insights)}")
    assert is_valid, "Valid comparison slide should pass validation"
    assert len(valid_insights) == 1

def test_invalid_comparison_slide_missing_nested():
    """Test comparison slide with missing nested fields"""
    insights = [{
        "title": "US vs India Hiring Performance",
        "data": {
            "type": "comparison",
            "left": {
                "title": "United States"
                # Missing: points
            },
            "right": {
                "title": "India",
                "points": []  # Empty points
            }
        }
    }]
    
    is_valid, errors, valid_insights = insights_validator.validate_insights(insights)
    print(f"\n✗ Invalid Comparison Slide Test (missing nested fields):")
    print(f"  Valid: {is_valid}")
    print(f"  Errors: {errors}")
    print(f"  Valid insights: {len(valid_insights)}")
    assert not is_valid, "Comparison slide with missing nested fields should fail"
    assert len(errors) > 0

def test_valid_two_column_slide():
    """Test valid two-column slide structure"""
    insights = [{
        "title": "Q4 Action Plan",
        "data": {
            "type": "two_column",
            "sections": [
                {
                    "title": "Immediate Actions (30 Days)",
                    "content": [
                        "Establish shared interview calendar blocks",
                        "Deploy async technical assessments"
                    ]
                },
                {
                    "title": "Long-Term Initiatives (90 Days)",
                    "content": [
                        "Build multi-location hiring playbook",
                        "Implement predictive analytics"
                    ]
                }
            ]
        }
    }]
    
    is_valid, errors, valid_insights = insights_validator.validate_insights(insights)
    print(f"\n✓ Valid Two-Column Slide Test:")
    print(f"  Valid: {is_valid}")
    print(f"  Errors: {errors}")
    print(f"  Valid insights: {len(valid_insights)}")
    assert is_valid, "Valid two-column slide should pass validation"
    assert len(valid_insights) == 1

def test_invalid_slide_type():
    """Test slide with invalid type"""
    insights = [{
        "title": "Invalid Slide",
        "data": {
            "type": "summary",  # Invalid type
            "content": ["Some content"]
        }
    }]
    
    is_valid, errors, valid_insights = insights_validator.validate_insights(insights)
    print(f"\n✗ Invalid Slide Type Test:")
    print(f"  Valid: {is_valid}")
    print(f"  Errors: {errors}")
    print(f"  Valid insights: {len(valid_insights)}")
    assert not is_valid, "Slide with invalid type should fail"
    assert len(errors) > 0

def test_mixed_valid_invalid():
    """Test mix of valid and invalid insights"""
    insights = [
        {
            "title": "Valid Metric",
            "data": {
                "type": "metric",
                "value": "76 days",
                "label": "Average Time",
                "context": ["Point 1"]
            }
        },
        {
            "title": "Invalid - Missing Data",
            # Missing: data field
        },
        {
            "title": "Valid Analysis",
            "data": {
                "type": "analysis",
                "content": ["Point 1", "Point 2"]
            }
        },
        {
            "title": "Invalid - Empty Content",
            "data": {
                "type": "analysis",
                "content": []
            }
        }
    ]
    
    is_valid, errors, valid_insights = insights_validator.validate_insights(insights)
    print(f"\n⚠ Mixed Valid/Invalid Test:")
    print(f"  Valid: {is_valid}")
    print(f"  Errors: {errors}")
    print(f"  Valid insights: {len(valid_insights)}/4")
    assert not is_valid, "Mixed insights should fail overall validation"
    assert len(valid_insights) == 2, "Should have 2 valid insights"
    assert len(errors) > 0

def test_validation_report_format():
    """Test validation report formatting"""
    insights = [
        {
            "title": "Invalid Metric",
            "data": {
                "type": "metric",
                "value": "76 days"
                # Missing: label, context
            }
        }
    ]
    
    is_valid, errors, valid_insights = insights_validator.validate_insights(insights)
    report = insights_validator.format_validation_report(errors, len(valid_insights), len(insights))
    
    print(f"\n📋 Validation Report Format Test:")
    print(report)
    assert "VALIDATION ERRORS" in report
    assert "Valid insights: 0/1" in report

if __name__ == "__main__":
    print("="*80)
    print("AI INSIGHTS VALIDATOR TEST SUITE")
    print("="*80)
    
    try:
        # Valid structure tests
        test_valid_metric_slide()
        test_valid_analysis_slide()
        test_valid_recommendation_slide()
        test_valid_comparison_slide()
        test_valid_two_column_slide()
        
        # Invalid structure tests
        test_invalid_metric_slide_missing_fields()
        test_invalid_analysis_slide_empty_content()
        test_invalid_comparison_slide_missing_nested()
        test_invalid_slide_type()
        
        # Mixed scenarios
        test_mixed_valid_invalid()
        test_validation_report_format()
        
        print("\n" + "="*80)
        print("✓ ALL TESTS PASSED")
        print("="*80)
        
    except AssertionError as e:
        print(f"\n✗ TEST FAILED: {e}")
        sys.exit(1)
    except Exception as e:
        print(f"\n✗ UNEXPECTED ERROR: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
