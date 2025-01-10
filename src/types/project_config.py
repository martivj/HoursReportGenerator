from abc import ABC, abstractmethod
from datetime import date
from typing import Dict, List

from src.types.dataclasses import ProjectPart


class ProjectConfig(ABC):
    @abstractmethod
    def get_project_parts(self) -> List[ProjectPart]:
        """Return list of project parts with their date ranges"""
        pass

    @abstractmethod
    def get_groupings(self) -> Dict[str, List[str]]:
        """Return the groupings for the project parts"""
        pass

    @abstractmethod
    def label_session(self, session_date: date, description: str) -> str:
        """Label each session based on date and description"""
        pass
