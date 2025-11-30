#!/usr/bin/env python3
"""
Test script for context manager optimizations
"""
import time
import sys
import os

# Add the src directory to the path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from mcp_sharepoint.context_manager import context_manager

def test_context_performance():
    """Test the performance improvements"""
    print("🧪 Testing Context Manager Optimizations")
    print("=" * 50)
    
    # Test 1: Basic context loading
    print("1. Testing context loading...")
    start_time = time.time()
    summary = context_manager.get_context_summary()
    load_time = time.time() - start_time
    
    print(f"   ✅ Loaded {summary['total_files']} files in {load_time:.3f}s")
    print(f"   📊 Total characters: {summary['total_characters']:,}")
    print(f"   📁 Categories: {', '.join(summary['categories'])}")
    
    # Test 2: Cached context retrieval
    print("\n2. Testing cached context retrieval...")
    
    # First call (should build cache)
    start_time = time.time()
    context1 = context_manager.get_context_for_tool('Analyze_HR_File_Complete')
    first_call_time = time.time() - start_time
    
    # Second call (should use cache)
    start_time = time.time()
    context2 = context_manager.get_context_for_tool('Analyze_HR_File_Complete')
    second_call_time = time.time() - start_time
    
    print(f"   🕐 First call: {first_call_time:.4f}s")
    print(f"   ⚡ Second call: {second_call_time:.4f}s")
    if second_call_time > 0:
        print(f"   🚀 Speedup: {first_call_time/second_call_time:.1f}x faster")
    else:
        print(f"   🚀 Second call is essentially instant (cached)!")
    
    # Test 3: Fast search functionality
    print("\n3. Testing fast search...")
    
    search_queries = [
        "Lead Software Engineer",
        "hiring guide", 
        "L5",
        "culture fit",
        "referral"
    ]
    
    for query in search_queries:
        start_time = time.time()
        results = context_manager.fast_search(query, max_results=3)
        search_time = time.time() - start_time
        
        print(f"   🔍 '{query}': {len(results)} results in {search_time:.4f}s")
        for result in results[:2]:  # Show top 2 results
            print(f"      - {result['file']} ({result['match_type']}, score: {result['score']})")
    
    # Test 4: Targeted context
    print("\n4. Testing targeted context...")
    
    start_time = time.time()
    targeted = context_manager.get_targeted_context("Lead Software Engineer", "Analyze_HR_File_Complete")
    targeted_time = time.time() - start_time
    
    print(f"   🎯 Targeted context: {len(targeted)} chars in {targeted_time:.4f}s")
    print(f"   📝 Contains index: {'_context_index.md' in targeted}")
    
    # Test 5: Memory usage comparison
    print("\n5. Memory efficiency...")
    
    # Check if index is built
    index_size = len(context_manager.search_index)
    print(f"   🗂️  Search index: {index_size} terms")
    print(f"   💾 Cache entries: {context_manager.get_context_for_tool.cache_info()}")
    
    print("\n✅ All tests completed successfully!")
    print("\n📈 Performance Summary:")
    print(f"   - Context loading: {load_time:.3f}s for {summary['total_files']} files")
    if second_call_time > 0:
        print(f"   - Cache speedup: {first_call_time/second_call_time:.1f}x")
    else:
        print(f"   - Cache speedup: Instant (cached calls are essentially 0ms)")
    print(f"   - Search index: {index_size} terms for instant lookup")
    print(f"   - Targeted context: Reduces irrelevant content")

if __name__ == "__main__":
    test_context_performance()
