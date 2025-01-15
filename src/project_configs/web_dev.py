from datetime import date
from typing import Dict, List
from src.types.project_config import ProjectConfig, ProjectPart


class WebDevConfig(ProjectConfig):
    @property
    def display_name(self) -> str:
        return "IT2810 WebDev"

    @property
    def description(self) -> str:
        return """
        <p class="text-base text-black font-bold mb-2 not-italic">Configuration for the IT2810 Web Development course</p>
        <p>Project parts are labeled based on dates and description keywords:</p>
        <ul class="list-disc pl-5 space-y-1 my-2">
            <li>Start the description with "Peer Review" for peer review work</li>
            <li>No keywords given -> assumes code work</li>
        </ul>
        <p>Deadlines follow the 2024 course schedule.</p>
        """

    def get_project_parts(self) -> List[ProjectPart]:
        return [
            # Project 1
            ProjectPart("P1 Part 1", date(2024, 8, 1), date(2024, 9, 20)),
            ProjectPart("Peer Review P1 Part 1", date(2024, 9, 21), date(2024, 9, 27)),
            ProjectPart("P1 Part 1", date(2024, 9, 21), date(2024, 9, 27)),
            # Project 2 Part 1
            ProjectPart("P2 Part 1", date(2024, 9, 28), date(2024, 10, 11)),
            ProjectPart(
                "Peer Review P2 Part 1", date(2024, 10, 12), date(2024, 10, 18)
            ),
            ProjectPart("P2 Part 1", date(2024, 10, 12), date(2024, 10, 18)),
            # Project 2 Part 2
            ProjectPart("P2 Part 2", date(2024, 10, 19), date(2024, 11, 1)),
            ProjectPart("Peer Review P2 Part 2", date(2024, 11, 2), date(2024, 11, 8)),
            ProjectPart("P2 Part 2", date(2024, 11, 2), date(2024, 11, 8)),
            # Project 2 Part 3
            ProjectPart("P2 Part 3", date(2024, 11, 9), date(2024, 11, 20)),
            ProjectPart(
                "Peer Review P2 Part 3", date(2024, 11, 21), date(2024, 11, 29)
            ),
            # Final Delivery
            ProjectPart("Final Delivery", date(2024, 11, 21), date(2024, 12, 6)),
        ]

    def get_groupings(self) -> Dict[str, List[str]]:
        return {
            "Project 1": ["P1 Part 1"],
            "Project 2": ["P2 Part 1", "P2 Part 2", "P2 Part 3", "Final Delivery"],
            "Peer Reviews": [
                "Peer Review P1 Part 1",
                "Peer Review P2 Part 1",
                "Peer Review P2 Part 2",
                "Peer Review P2 Part 3",
            ],
        }

    def label_session(self, session_date: date, description: str) -> str:
        parts = self.get_project_parts()
        is_peer_review = description.lower().startswith("peer review")

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

        # If there are peer review parts and this is a peer review
        if is_peer_review:
            peer_review_parts = [p for p in matching_parts if "Peer Review" in p.name]
            if peer_review_parts:
                return peer_review_parts[0].name

        # Otherwise use the first non-peer review part
        regular_parts = [p for p in matching_parts if "Peer Review" not in p.name]
        if regular_parts:
            return regular_parts[0].name

        # Fallback to first matching part if none of the above conditions met
        return matching_parts[0].name
