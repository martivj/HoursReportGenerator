# HoursReportGenerator

Tool for generating formatted Excel reports from time tracking CSV files.

Note that this tool is designed for specific project structures and may require customization for other projects. The projects supported are the IT2810 Web Development course (2024) and the IT2901 Informatics Project II course (2025) at NTNU.

## CSV Format Requirements

The script expects CSV files with the following format:

### Required Columns
The CSV must include these columns:
- `startTime`: ISO 8601 datetime string (e.g. "2024-09-05T09:58:00.000000000Z")
- `duration`: Integer number of minutes (can be wrapped in quotes, e.g. "75")
- `description`: String description of the session wrapped in triple quotes 

Note: description can contain single quotes but not double quotes

### Example Row
```csv
"2024-09-05T09:58:00.000000000Z",75,"""working on issue#45: added 'quote' thing"""
```

## Setup

1. Create a virtual environment (optional but recommended, you can also use conda):

```bash
python -m venv venv
source venv/bin/activate  # Linux/Mac
venv\Scripts\activate     # Windows
```

2. Install dependencies:

```bash
pip install -r requirements.txt
```

## Usage

1. Place your CSV files in the `data/` directory

2. Configure project tracking in `main.py`. Choose or create a config class:
   - [`WebDevConfig`](src/project_configs/web_dev.py) - For the IT2810 Web Development course (2024)
   - [`ITP2Config`](src/project_configs/itp2.py) - For the IT2901 Informatics Project II course (2025)
   
   Example `main.py`:
   ```python
   from src.project_configs.web_dev import WebDevConfig
   from src.report_generator import ReportGenerator

   def main():
       # Use WebDevConfig or ITP2Config, or create your own
       config = WebDevConfig()
       generator = ReportGenerator(config)
       generator.generate("HoursReport.xlsx")

   if __name__ == "__main__":
       main()
    ```

3. Run the script:

```bash
python main.py
```
4. Find generated reports in the `reports` directory

### Creating Custom Configs

To create a new project structure, create a class that extends [`ProjectConfig`](src/project_configs/project_config.py) and implement the following methods:

1. Define project parts with date ranges
2. Create groupings of related parts
3. Implement custom logic for labeling time entries

See existing configs for examples:

- [`web_dev.py`](src\project_configs\web_dev.py) - WebDevConfig
- [`itp2.py`](src\project_configs\itp2.py) - ITP2Config






### Generator Options

The `ReportGenerator` accepts two optional configuration parameters:

- `total_sheet_first=True` - Controls worksheet order in Excel:
  - `True`: Places the summary sheet as the first tab (default)
  - `False`: Places individual project sheets first, summary at end
- `close_open_excel=True` - Handles Excel application conflicts:
  - `True`: Closes any open Excel instances, reopens after generating report (default)
  - `False`: Does not open or close Excel instances (may cause file access issues)

Example with options:
```python
generator = ReportGenerator(
    config,
    total_sheet_first=False, # Put summary as the last sheet
    close_open_excel=True    # Close Excel before generating, reopen after
)
```

## Future Plans

- Create a Flask web app for uploading CSV files and downloading reports
