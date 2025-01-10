from dataclasses import dataclass
from datetime import date
from pandas import DataFrame


@dataclass
class ProjectInfo:
    title: str
    primary_color: str
    secondary_color: str


@dataclass
class ProjectPart:
    name: str
    start_date: date
    end_date: date


@dataclass
class GroupSummary:
    total_hours: float
    hours_per_week: DataFrame
