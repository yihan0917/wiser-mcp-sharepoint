# 🚀 Quick Start: Generate Training Data

## One-Command Generation

```bash
# Activate environment and run
source venv/bin/activate && python generate_training_data.py
```

That's it! The script will:
1. ✅ Generate 4 Excel files with 75-100 records each
2. ✅ Upload them to SharePoint Data folder
3. ✅ Name them with record counts

---

## What Gets Generated

| File | Records | Columns | Purpose |
|------|---------|---------|---------|
| **Filled Positions** | 75-100 | 17 | Candidate hiring data |
| **Time In Step Q3** | 75-100 | 27 | Pipeline time tracking |
| **Recruiting Report** | 75-100 | 19 | Funnel metrics |
| **All Departments** | 75-100 | 24 | Department analytics |

---

## Expected Output

```
================================================================================
TRAINING DATA GENERATOR FOR HR ANALYTICS
================================================================================

🔄 Generating training datasets...

📊 Generating File 1: Filled Positions (90 records)...
   ✅ Generated 90 records

📊 Generating File 2: Time In Step Q3 (97 records)...
   ✅ Generated 97 records

📊 Generating File 3: Recruiting Report (86 records)...
   ✅ Generated 86 records

📊 Generating File 4: All Departments Report (91 records)...
   ✅ Generated 91 records

================================================================================
✅ ALL FILES GENERATED SUCCESSFULLY!
================================================================================

📤 Uploading to SharePoint...
================================================================================
✅ Training_Filled_Positions_90_records.xlsx
   Size: 13,439 bytes
✅ Training_Time_In_Step_Q3_97_records.xlsx
   Size: 18,835 bytes
✅ Training_Recruiting_Report_86_records.xlsx
   Size: 14,780 bytes
✅ Training_All_Depts_Report_91_records.xlsx
   Size: 16,015 bytes

================================================================================
✅ UPLOAD COMPLETE!
================================================================================

📁 Files uploaded to SharePoint Data folder
🎯 Ready for HR analytics and training!
```

---

## Customization Examples

### Fixed Number of Records

Edit `generate_training_data.py`:

```python
# In main() function, change:
local_path, upload_name = generate_file1_filled_positions(num_records=100)  # Fixed 100
```

### Different Date Range

```python
# In generate_file1_filled_positions(), change:
creation_date = random_date(datetime(2024, 1, 1), datetime(2024, 12, 31))  # 2024 only
```

### More Applicants Per Position

```python
# In generate_file3_recruiting_report(), change:
applicants = random.randint(100, 2000)  # Instead of 50-1000
```

---

## Troubleshooting

### ❌ "ModuleNotFoundError"
```bash
pip install pandas openpyxl msal requests python-dotenv
```

### ❌ "Authentication failed"
Check your `.env` file has:
```env
SHP_ID_APP=...
SHP_ID_APP_SECRET=...
SHP_TENANT_ID=...
```

### ❌ "Upload failed"
- Verify internet connection
- Check SharePoint Data folder exists
- Confirm write permissions

---

## Files Created

**Script**: `generate_training_data.py`  
**Documentation**: `TRAINING_DATA_GENERATOR_README.md`  
**Quick Start**: This file

---

**Need more help?** See `TRAINING_DATA_GENERATOR_README.md` for full documentation.
