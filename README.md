# HoursReportGenerator

Tool for generating formatted Excel reports from time tracking CSV files.

Note that this tool was initially designed for very specific project structures and will require customization to adapt to other projects. The projects currently supported are the IT2810 Web Development course (2024) and the IT2901 Informatics Project II course (2025) at NTNU.

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

3. Create a `.env` file in the root directory:

```env
FLASK_APP=run.py
SECRET_KEY=your-secret-key-here
MAX_FILES=10
UPLOAD_FOLDER=temp/uploads
REDIS_URL=redis://redis:6379/0
```

## Running the Application

1. Start the Flask server:

```bash
python run.py
```

2. Open your browser to `http://localhost:5000`

3. Select your project type, choose CSV files to upload, and generate your report

## CSV Format Requirements

The application expects CSV files with the following format:

### Required Columns
- `startTime`: ISO 8601 datetime string (e.g. "2024-09-05T09:58:00.000000000Z")
- `duration`: Integer number of minutes (can be wrapped in quotes, e.g. "75")
- `description`: String description of the session wrapped in triple quotes 

Note: description can contain single quotes but not double quotes

### Example Row
```csv
"2024-09-05T09:58:00.000000000Z",75,"""working on issue#45: added 'quote' thing"""
```

## Creating Custom Project Configs

1. Create a new file in `src/project_configs/` (e.g. `my_project.py`)

2. Create a class that extends [`ProjectConfig`](src/project_configs/project_config.py):

```python
from datetime import date
from typing import Dict, List
from src.types.project_config import ProjectConfig, ProjectPart

class MyProjectConfig(ProjectConfig):
    @property
    def display_name(self) -> str:
        return "My Project Name"

    def get_project_parts(self) -> List[ProjectPart]:
        return [
            ProjectPart("Phase 1", date(2024, 1, 1), date(2024, 3, 1)),
            ProjectPart("Phase 2", date(2024, 3, 2), date(2024, 6, 1))
        ]

    def get_groupings(self) -> Dict[str, List[str]]:
        return {
            "Development": ["Phase 1", "Phase 2"]
        }

    def label_session(self, session_date: date, description: str) -> str:
        # Custom logic to label time entries, description can be used
        return "Phase 1" if session_date < date(2024, 3, 2) else "Phase 2"
```

The display name is used in the web interface to select the project type. The project parts are used to separate time entries in the report. The groupings are used to group project parts together in the report summaries. The label session method is used to determine which project part a time entry belongs to, custom logic can be added here based on the session date and description.

See existing configs for examples:
- [`web_dev.py`](src/project_configs/web_dev.py) - WebDevConfig 
- [`itp2.py`](src/project_configs/itp2.py) - ITP2Config



## Command Line Usage

You can still use the tool via command line:

1. Place your CSV files in the `data/` directory

2. Create and run a script (e.g. `main.py`):
```python
from src.project_configs.web_dev import WebDevConfig
from src.report_generator import ReportGenerator

def main():
    config = WebDevConfig()
    generator = ReportGenerator(config)
    generator.generate("HoursReport.xlsx")

if __name__ == "__main__":
    main()
```

### Report Generator Options

The `ReportGenerator` accepts two optional configuration parameters:

- `total_sheet_first=True` - Controls worksheet order in Excel:
  - `True`: Places the summary sheet as the first tab (default)
  - `False`: Places individual project sheets first, summary at end
- `close_open_excel=True` - Handles Excel instance management:
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

The output file name can be changed in the `generate` method:
```python
output_name = "ExampleReport.xlsx"
output_path = Path("reports") / output_name
generator.generate(str(output_path))
```