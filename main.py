from src.project_configs.web_dev import WebDevConfig
from src.report_generator import ReportGenerator


def main() -> None:

    config = WebDevConfig()
    total_sheet_first = True
    close_open_excel = True

    generator = ReportGenerator(
        config,
        total_sheet_first=total_sheet_first,
        close_open_excel=close_open_excel,
    )

    generator.generate("HoursReport.xlsx")


if __name__ == "__main__":
    main()
