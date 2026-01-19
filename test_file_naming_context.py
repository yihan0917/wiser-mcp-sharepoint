"""
Test script to verify file naming convention is loaded in context manager
"""
import os
import sys
import importlib.util

# Directly load context_manager.py without going through __init__.py
context_manager_path = os.path.join(
    os.path.dirname(__file__), 
    'src', 
    'mcp_sharepoint', 
    'context_manager.py'
)

spec = importlib.util.spec_from_file_location("context_manager", context_manager_path)
context_manager_module = importlib.util.module_from_spec(spec)
sys.modules['context_manager'] = context_manager_module

# Mock the logger before loading
import logging
logger = logging.getLogger('mcp_sharepoint')
logging.basicConfig(level=logging.INFO)

# Create a mock common module with just the logger
class MockCommon:
    logger = logger

sys.modules['mcp_sharepoint.common'] = MockCommon()

# Now load the module
spec.loader.exec_module(context_manager_module)
ContextManager = context_manager_module.ContextManager

def test_file_naming_context():
    """Test that file naming convention is properly loaded"""
    
    print("Initializing ContextManager...")
    cm = ContextManager()
    
    # Test 1: Check if file_naming category exists
    print('\n=== Test 1: Check file_naming category ===')
    summary = cm.get_context_summary()
    print(f'Categories: {summary["categories"]}')
    print(f'File naming files: {summary["files_by_category"].get("file_naming", [])}')
    
    # Test 2: Get file_naming context
    print('\n=== Test 2: Get file_naming context ===')
    file_naming_context = cm.get_context_by_category('file_naming')
    if file_naming_context:
        print(f'✓ File naming context loaded: {len(file_naming_context)} characters')
        print(f'\nFirst 500 chars:\n{file_naming_context[:500]}...')
    else:
        print('✗ ERROR: File naming context not loaded!')
        return False
    
    # Test 3: Check tool context mapping
    print('\n=== Test 3: Check tool context for Analyze_HR_File_Complete ===')
    tool_context = cm.get_context_for_tool('Analyze_HR_File_Complete')
    if 'file_naming_convention.md' in tool_context:
        print('✓ File naming convention is included in tool context')
        print(f'Total tool context size: {len(tool_context)} characters')
    else:
        print('✗ ERROR: File naming convention not in tool context!')
        return False
    
    # Test 4: Search for specific terms
    print('\n=== Test 4: Search for file naming patterns ===')
    results = cm.search_context('SmartRecruiter', categories=['file_naming'])
    print(f'Found {len(results)} matches for "SmartRecruiter"')
    if results:
        print(f'✓ Match in: {results[0]["file"]} at line {results[0]["line"]}')
        print(f'Excerpt: {results[0]["excerpt"][:200]}...')
    
    # Test 5: Verify key patterns are in the context
    print('\n=== Test 5: Verify key patterns ===')
    patterns = ['content_source', 'WK##', 'Q1', 'Q2', 'Q3', 'Q4', '.xlsx']
    for pattern in patterns:
        if pattern in file_naming_context:
            print(f'✓ Pattern "{pattern}" found')
        else:
            print(f'✗ Pattern "{pattern}" NOT found')
    
    print('\n=== All Tests Passed! ===')
    return True

if __name__ == '__main__':
    test_file_naming_context()
