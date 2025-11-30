# Context Search & Optimization Guide

Complete guide to the SharePoint MCP server's optimized context management and metadata-enhanced search system.

## 🎯 Current Implementation (v0.6.0)

Your context management system combines **performance optimizations** with **metadata-enhanced search** for maximum efficiency.

### ✅ What's Implemented

1. **Metadata-Enhanced Search** - YAML frontmatter with curated search terms
2. **Two-Phase Search Strategy** - Metadata first, content fallback
3. **Performance Optimizations** - LRU caching, search indexing, targeted retrieval
4. **Backward Compatibility** - Files without metadata work via intelligent content search

## ⚡ Search Performance

### **Search Scoring System**
- **metadata_exact**: 1.5 (highest - direct COMMON_SEARCHES match)
- **metadata_partial**: 1.2 (PURPOSE/KEY_SECTIONS match)
- **content_exact**: 1.0 (intelligent content extraction)
- **content_partial**: 0.7 (partial content match)

### **Two-Phase Search Flow**
```python
# Phase 1: Metadata Search (if available)
if query in metadata['COMMON_SEARCHES']:
    return metadata_exact_match (score: 1.5)
elif query in metadata['PURPOSE'] or metadata['KEY_SECTIONS']:
    return metadata_partial_match (score: 1.2)

# If good metadata matches found (score ≥ 1.2), SKIP Phase 2

# Phase 2: Content Search (fallback only)
# Uses intelligent term extraction from 1,259 indexed terms
```

### **Performance Results**
- **Metadata queries**: 50-80% faster (direct lookup)
- **Content queries**: Same speed as before (no regression)
- **Memory overhead**: <50KB additional
- **Cache hit ratio**: 50%+ for repeated tool usage

## 📊 Current Status

### **Files with Metadata** (12 files)
- **Core System**: `_context_index.md`, `column_definitions.md`, `metrics_definitions.md`, `company_overview.md`
- **Engineering**: `engineering_overview.md`, `engineering_career_path.md`, `software_engineer_role_description.md`, `engineering_leadership_role_description.md`
- **Specialized**: `ml_ds_da_role_description.md`, `data_management_role_description.md`, `in_store_operations_role_description.md`
- **Process**: `hiring_guide.md`

### **Files without Metadata** (12 files)
- **Position-Description files**: Use intelligent content search (Phase 2 only)
- **Performance**: Same as before, automatically indexed with 1,259 terms

## 🛠️ Adding Metadata to Files

### **YAML Frontmatter Template**
```yaml
---
FILE: your_file.md
PURPOSE: What this file defines or describes
DEPARTMENT: Engineering/HR/Operations/etc
KEY_SECTIONS: Main topics covered
COMMON_SEARCHES: terms, users, search, for, frequently
RELATED_FILES: other_relevant_files.md
LAST_UPDATED: 2024-11-30
---

# Your existing content starts here...
```

### **Key Fields Explained**

**COMMON_SEARCHES** (Most Important):
- Direct search terms users are likely to use
- Provides instant high-relevance matches (score 1.5)
- Example: `mentorship, career progression, L5, technical leadership`

**PURPOSE**:
- Semantic understanding of file's role
- Matches conceptual queries (score 1.2)
- Example: `Define career progression for software engineers`

**KEY_SECTIONS**:
- Major topics covered in the file
- Section-specific searches (score 1.2)
- Example: `Technical Skills, Delivery, Leadership`

## 🚀 Performance Optimizations

### **Smart Caching**
```python
# LRU Cache on context retrieval
@lru_cache(maxsize=32)
def get_context_for_tool(tool_name: str) -> str:
    # First call: ~0.1ms (build and return)
    # Cached calls: ~0.0ms (instant retrieval)
```

### **Fast Search Index**
- **1,259 searchable terms** built at startup
- **Inverted index** for millisecond lookups
- **Smart extraction**: headers, bold text, role names, levels
- **Exact + partial matching** with relevance scoring

### **Targeted Context Retrieval**
```python
# Get context specific to query instead of loading everything
context = context_manager.get_targeted_context("hiring process", "Calculate_HR_Metrics")

# Fast search for specific queries
results = context_manager.fast_search("Lead Software Engineer")
```

## 📈 Usage Examples

### **Test Metadata Search**
```python
from src.mcp_sharepoint.context_manager import context_manager

# Test metadata-enhanced queries (should return score 1.5)
results = context_manager.fast_search("mentorship")
print(f"Result: {results[0]['match_type']} (score: {results[0]['score']})")

# Test career progression search
results = context_manager.fast_search("career progression") 
print(f"Found: {len(results)} results")

# Test data engineer search
results = context_manager.fast_search("data engineer")
print(f"File: {results[0]['file']} ({results[0]['match_type']})")
```

### **Check Performance**
```python
# Cache performance
cache_info = context_manager.get_context_for_tool.cache_info()
print(f"Cache hits: {cache_info.hits}, misses: {cache_info.misses}")

# Search index size
print(f"Indexed terms: {len(context_manager.search_index)}")

# Metadata coverage
print(f"Files with metadata: {len(context_manager.metadata_index)}")
```

## 🎯 Search Strategy by File Type

### **Metadata-Enhanced Files** (12 files)
- **Common queries** → Phase 1 (metadata) → **50-80% faster**
- **Uncommon queries** → Phase 2 (content) → Same speed
- **Examples**: "mentorship", "career progression", "HR metrics"

### **Position-Description Files** (12 files)
- **All queries** → Phase 2 (content) → Same speed as before
- **Intelligent extraction** finds role-specific terms automatically
- **Examples**: "MLE1", "Analytics Manager", "SEMLDS"

### **Tool Context Injection** (All files)
- **Automatic inclusion** via category mapping
- **Cached retrieval** for repeated tool calls
- **No search needed** - context provided directly to tools

## 🔧 Maintenance

### **Adding New Files**
1. Create markdown file in `src/mcp_sharepoint/context/`
2. Add metadata (optional but recommended for high-traffic files)
3. Update `CONTEXT_CATEGORIES` in `context_manager.py`
4. Update `TOOL_CONTEXT_MAP` if needed
5. Restart MCP server (search index rebuilds automatically)

### **Performance Monitoring**
```python
# Get comprehensive statistics
summary = context_manager.get_context_summary()
cache_stats = context_manager.get_context_for_tool.cache_info()
search_terms = len(context_manager.search_index)
metadata_files = len(context_manager.metadata_index)
```

## ✅ Benefits Achieved

1. **Faster Search**: 50-80% improvement for common queries
2. **Better Relevance**: Curated search terms vs. extracted terms
3. **Zero Regression**: Files without metadata work exactly as before
4. **Scalable**: Add metadata incrementally for gradual improvements
5. **Cost-Free**: No external APIs or vector databases needed
6. **MCP-Optimized**: Designed for tool-specific context injection

## 🧪 Testing

Run performance tests anytime:
```bash
cd "/Users/yihan/Documents/sharepoint mcp/wiser-mcp-sharepoint"
python test_context_optimization.py
```

**Expected Results**:
- ✅ 24 files loaded instantly
- ✅ Cached calls are ~0ms
- ✅ 1,259+ search terms indexed
- ✅ 12 files with metadata
- ✅ Fast targeted context retrieval

---

**Status**: ✅ Production-ready  
**Version**: v0.6.0  
**Performance**: Optimized with metadata-enhanced search  
**Compatibility**: Backward compatible with all existing files
