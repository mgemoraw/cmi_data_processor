CMI Data Platform
│
├── Dashboard
├── Project Explorer
├── Data Processing
│   ├── File Splitter
│   ├── File Renamer
│   ├── Data Cleaner
│   └── Data Aggregator
│
├── Analysis Tools
│   ├── Workbook Inspector
│   ├── Sheet Statistics
│   ├── Missing Data Analysis
│   ├── Duplicate Detection
│   ├── Template Validation
│   └── Productivity Summary
│
├── File Management
│   ├── Browse Files
│   ├── Open Workbook
│   ├── Bulk Rename
│   └── Archive Projects
│
├── Reports
│   ├── Export Summary
│   ├── Generate Reports
│   └── Project Statistics
│
└── Settings
    ├── Working Directory
    ├── Templates
    ├── Output Locations
    └── Preferences


cmi_data_platform/
│
├── main.py
│
├── core/
│   ├── app_settings.py
│   ├── worker.py
│   ├── logger.py
│   └── constants.py
│
├── ui/
│   ├── main_window.py
│   │
│   ├── pages/
│   │   ├── dashboard_page.py
│   │   ├── explorer_page.py
│   │   ├── processing_page.py
│   │   ├── analysis_page.py
│   │   ├── reports_page.py
│   │   └── settings_page.py
│   │
│   └── widgets/
│       ├── file_browser.py
│       ├── log_console.py
│       ├── progress_panel.py
│       └── metadata_card.py
│
├── engines/
│   ├── splitter_engine.py
│   ├── cleaning_engine.py
│   ├── naming_engine.py
│   └── processing_engine.py
│
├── analysis/
│   ├── workbook_inspector.py
│   ├── duplicate_detector.py
│   ├── missing_data.py
│   ├── validator.py
│   └── statistics.py
│
├── reports/
│   ├── excel_report.py
│   ├── pdf_report.py
│   └── summary_report.py
│
├── templates/
│
├── settings/
│   └── settings.json
│
└── output/