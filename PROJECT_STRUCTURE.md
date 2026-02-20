# Project Structure

```
win32way/
├── frontend/
│   ├── dashboard.html
│   └── micsoftlistui.html
│
├── backend/
│   ├── main.py
│   ├── graph.py
│   ├── server.py
│   ├── filterApi.py
│   ├── create_tickets.py
│   ├── create_filtered_tickets.py
│   ├── add_filter.py
│   ├── edit_filters.py
│   ├── modify_filters.py
│   └── custom_date_filter.py
│
├── database/
│   ├── ticket.json
│   ├── january_2026_tickets.json
│   ├── target_ticket.json
│   ├── outlook_emails.json
│   ├── email_filters.json
│   └── extracted_info.xlsx
│
├── tests/
│   ├── test.py
│   ├── test_all.py
│   ├── test_quick.py
│   ├── test_with_autoclick.py
│   ├── debug_filters.py
│   ├── debug_cdc_filter.py
│   ├── diagnose.py
│   └── testfile/
│       ├── test_filters.py
│       ├── test_filter_ids.py
│       ├── test_fixed.py
│       ├── test_mx_english.py
│       └── test_old_backup.py
│
├── utils/
│   ├── apiread.py
│   ├── fix_outlook_warning.py
│   └── astscan.py
│
├── backup/
│   ├── email_filters.json.backup
│   ├── test.py.backup
│   ├── oldfile/
│   └── oldtest.py
│
├── config.cfg
├── requirements.txt
└── PROJECT_STRUCTURE.md
```

## Import Issues to Fix

After moving files, these imports need to be updated:

1. **backend/main.py** - Change: `from graph import Graph` → `from .graph import Graph`

2. **backend/edit_filters.py, add_filter.py, modify_filters.py, filterApi.py** - Change: `from test import EmailFilterManager` → `from tests.test import EmailFilterManager`

3. **backend/server.py** - Script paths need update (lines 15, 17): 
   - Change: `test_quick.py` → `tests/test_quick.py`
   - Change: `test_all.py` → `tests/test_all.py`

4. Add `__init__.py` to backend/ folder to make it a Python package

