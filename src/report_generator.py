from pathlib import Path
from openpyxl import Workbook
import os

from src.data_processing.processor import DataProcessor
from src.formatters.excel_formatter import ExcelFormatter
from src.types.dataclasses import ProjectInfo
from src.types.project_config import ProjectConfig


class ReportGenerator:
    def __init__(
        self,
        config: ProjectConfig,
        total_sheet_first: bool = True,
        close_open_excel: bool = True,
        data_dir: Path = None,
    ):
        self.config = config
        self.total_sheet_first = total_sheet_first
        self.close_open_excel = close_open_excel

        # Use provided paths or defaults
        self._script_dir = Path(__file__).parent.parent
        self._data_dir = data_dir or (self._script_dir / "data")

        # Pre-defined color schemes (primary, secondary)
        self.COLOR_SCHEMES = [
            ("0072BC", "D9EAF7"),  # Blue theme
            ("FF5733", "FFD9CC"),  # Orange theme
            ("FFC300", "FFF2CC"),  # Yellow theme
            ("4CAF50", "C8E6C9"),  # Green theme
            ("9C27B0", "E1BEE7"),  # Purple theme
            ("F44336", "FFCDD2"),  # Red theme
            ("607D8B", "CFD8DC"),  # Blue Grey theme
            ("795548", "D7CCC8"),  # Brown theme
        ]

    def _get_title_from_filename(self, filepath: str, existing_titles: list) -> str:
        """Extract title from filename and handle duplicates with (x) suffix"""
        filename = Path(filepath).stem
        base_title = filename.split("_")[0].lower().capitalize()

        if base_title not in existing_titles:
            return base_title

        counter = 2
        while f"{base_title} ({counter})" in existing_titles:
            counter += 1
        return f"{base_title} ({counter})"

    def _process_csv_files(self) -> list:
        """Process all CSV files in the data directory"""
        processor = DataProcessor(self.config)
        csv_files = [f for f in os.listdir(self._data_dir) if f.endswith(".csv")]

        datasets_info = []
        existing_titles = []

        for i, csv_file in enumerate(csv_files):
            color_scheme = self.COLOR_SCHEMES[i % len(self.COLOR_SCHEMES)]
            title = self._get_title_from_filename(csv_file, existing_titles)
            existing_titles.append(title)

            datasets_info.append(
                {
                    "csv_path": os.path.join(self._data_dir, csv_file),
                    "info": ProjectInfo(
                        title=title,
                        primary_color=color_scheme[0],
                        secondary_color=color_scheme[1],
                    ),
                }
            )

        processed_data = []
        for dataset in datasets_info:
            group_dfs, group_summaries = processor.get_processed_data(
                dataset["csv_path"]
            )
            processed_data.append((dataset["info"], group_dfs, group_summaries))

        return processed_data

    def generate(self, output_path: str) -> None:
        """Generate the Excel report at the specified path"""
        processed_data = self._process_csv_files()

        wb = None
        try:
            wb = Workbook()
            formatter = ExcelFormatter(
                wb=wb,
                data=processed_data,
                output_path=str(output_path),
                total_sheet_first=self.total_sheet_first,
                close_open_excel=self.close_open_excel,
            )
            formatter.format()
        finally:
            if wb:
                wb.close()
