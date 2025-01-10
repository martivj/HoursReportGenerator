from datetime import date
from typing import Dict, List
from src.types.project_config import ProjectConfig, ProjectPart


class ITP2Config(ProjectConfig):

    @property
    def display_name(self) -> str:
        return "IT2901 ITP2"

    def get_project_parts(self) -> List[ProjectPart]:
        return [
            # Preliminary Report
            ProjectPart(
                "Report start",  # Include "Report" keyword in description
                date(2025, 1, 13),
                date(2025, 2, 17),
            ),
            ProjectPart(
                "Code start",
                date(2025, 1, 13),
                date(2025, 2, 17),
            ),
            ProjectPart(
                "Report Midterm",  # Include "Report" keyword in description
                date(2025, 2, 18),
                date(2025, 3, 14),
            ),
            ProjectPart(
                "Self Accessment",  # Include "Self Accessment" keyword in description
                date(2025, 3, 15),
                date(2025, 3, 21),
            ),
            ProjectPart(
                "Code Midterm",
                date(2025, 2, 18),
                date(2025, 3, 21),
            ),
            ProjectPart(
                "Report Final",  # Include "Report" keyword in description
                date(2025, 3, 22),
                date(2025, 5, 9),
            ),
            ProjectPart(
                "Video",  # Include "Video" keyword in description
                date(2025, 5, 10),
                date(2025, 5, 16),
            ),
            ProjectPart(
                "Code Final",
                date(2025, 3, 22),
                date(2025, 5, 9),
            ),
        ]

    def get_groupings(self) -> Dict[str, List[str]]:
        return {
            "Code": ["Code start", "Code Midterm", "Code Final"],
            "Report": ["Report start", "Report Midterm", "Report Final"],
            "Self Accessment": ["Self Accessment"],
            "Video": ["Video"],
        }

    def label_session(self, session_date: date, description: str) -> str:
        parts = self.get_project_parts()
        is_report = "report" in description.lower()
        is_video = "video" in description.lower()
        is_self_accessment = "self accessment" in description.lower()

        matching_parts = []
        for part in parts:
            # Skip empty parts or parts without dates
            if (
                not part.name
                or not hasattr(part, "start_date")
                or not hasattr(part, "end_date")
            ):
                continue

            # Check if date falls within range
            if part.start_date <= session_date <= part.end_date:
                matching_parts.append(part)

        # If no matching parts found
        if not matching_parts:
            return "Unknown"

        # If there are report parts and this is a report
        if is_report:
            for part in matching_parts:
                if "report" in part.name.lower():
                    return part.name

        # If there are self accessment parts and this is a self accessment
        if is_self_accessment:
            for part in matching_parts:
                if "self accessment" in part.name.lower():
                    return part.name

        # If there are video parts and this is a video
        if is_video:
            for part in matching_parts:
                if "video" in part.name.lower():
                    return part.name

        # Fallback to first matching part if none of the above conditions met
        return matching_parts[0].name
