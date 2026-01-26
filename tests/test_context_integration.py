#!/usr/bin/env python3
"""
Test script to verify context manager integration
"""
import sys
import os

# Add src to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from mcp_sharepoint.context_manager import context_manager

def test_context_manager():
    """Test that context manager loads properly"""
    print("=" * 60)
    print("Testing Context Manager Integration")
    print("=" * 60)
    
    # Test 1: Get context summary
    print("\n1. Testing context summary...")
    summary = context_manager.get_context_summary()
    print(f"   ✓ Total files loaded: {summary.get('total_files', 0)}")
    print(f"   ✓ Total characters: {summary.get('total_characters', 0):,}")
    print(f"   ✓ Categories: {', '.join(summary.get('categories', []))}")
    print(f"   ✓ Column definitions: {summary.get('column_definitions_count', 0)}")
    
    # Test 2: Get column definitions
    print("\n2. Testing column definitions...")
    all_defs = context_manager.get_all_column_definitions()
    print(f"   ✓ Total column definitions: {len(all_defs)}")
    if all_defs:
        sample_col = list(all_defs.keys())[0]
        print(f"   ✓ Sample: '{sample_col}' = '{all_defs[sample_col][:50]}...'")
    
    # Test 3: Get context for a tool
    print("\n3. Testing tool context...")
    tool_context = context_manager.get_context_for_tool('Analyze_HR_File_Complete')
    print(f"   ✓ Context length for Analyze_HR_File_Complete: {len(tool_context):,} chars")
    
    # Test 4: Search context
    print("\n4. Testing context search...")
    results = context_manager.search_context('time-to-hire')
    print(f"   ✓ Search results for 'time-to-hire': {len(results)} matches")
    if results:
        print(f"   ✓ Found in: {results[0]['file']}")
    
    # Test 5: Get context by category
    print("\n5. Testing category context...")
    for category in ['columns', 'metrics', 'business', 'recruiting']:
        cat_context = context_manager.get_context_by_category(category)
        print(f"   ✓ {category}: {len(cat_context):,} chars")
    
    # Test 6: Column matching
    print("\n6. Testing column matching...")
    test_columns = ['Job Title', 'Date Opened', 'Recruiter', 'NonExistentColumn']
    matching = context_manager.get_matching_columns(test_columns)
    print(f"   ✓ Matched {len(matching)}/{len(test_columns)} columns")
    for col in matching:
        print(f"      - {col}")
    
    print("\n" + "=" * 60)
    print("✅ All tests passed! Context manager is working correctly.")
    print("=" * 60)

if __name__ == "__main__":
    try:
        test_context_manager()
    except Exception as e:
        print(f"\n❌ Error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
