from pandas import DataFrame
import pandas as pd
from typing import Dict, Tuple
from src.types.dataclasses import GroupSummary
from src.types.project_config import ProjectConfig


class DataProcessor:
    def __init__(self, project_config: ProjectConfig):
        self.project_config = project_config

    def _read_csv(self, file_path: str) -> DataFrame:
        """Read the CSV file into a DataFrame."""
        return pd.read_csv(file_path)

    def _preprocess_data(self, df: DataFrame) -> DataFrame:
        """Process the raw data into required format."""
        # Sort by full timestamp first
        df = df.sort_values(by="startTime", ascending=True)
        df["date"] = pd.to_datetime(df["startTime"], errors="coerce").dt.date
        df["duration_minutes"] = df["duration"]
        df["description"] = df["description"].str.replace('"', "")
        df = df[["date", "duration_minutes", "description"]]
        return df

    def _split_data(self, df: DataFrame) -> Dict[str, DataFrame]:
        """Split data by project parts based on config."""
        df["Part"] = df.apply(
            lambda row: self.project_config.label_session(
                row["date"], row["description"]
            ),
            axis=1,
        )

        groupings = self.project_config.get_groupings()

        group_dfs = {}

        # For each group, collect all matching parts at once
        for group_name, group_parts in groupings.items():
            matching_rows = df["Part"].isin(group_parts)
            if matching_rows.any():
                group_df = df[matching_rows].copy()
                group_df = group_df.reset_index(drop=True)
                group_dfs[group_name] = group_df
        return group_dfs

    def _calculate_summary(self, df: DataFrame) -> GroupSummary:
        """Calculate hours summary per dataset."""
        df.loc[:, "week"] = pd.to_datetime(df["date"]).dt.isocalendar().week
        df.loc[:, "duration_hours"] = df["duration_minutes"] / 60
        total_hours = df["duration_hours"].sum()
        hours_per_week = df.groupby("week")["duration_hours"].sum().reset_index()
        return total_hours, hours_per_week

    def _prepare_data_for_output(self, df: DataFrame) -> DataFrame:
        """Format data for Excel output."""
        df["week"] = pd.to_datetime(df["date"]).dt.isocalendar().week
        return df.rename(
            columns={
                "Part": "Part",
                "week": "Week",
                "date": "Date",
                "duration_minutes": "Minutes",
                "description": "Description",
            }
        )[["Part", "Week", "Date", "Minutes", "Description"]]

    def get_processed_data(
        self, file_path: str
    ) -> Tuple[Dict[str, DataFrame], Dict[str, GroupSummary]]:
        """Process the data and return dictionaries mapping group names to their data."""
        # Read and preprocess
        df: DataFrame = self._read_csv(file_path)
        df: DataFrame = self._preprocess_data(df)

        # Split into groups
        group_dfs: Dict[str, DataFrame] = self._split_data(df)

        # Calculate summaries and prepare data for each group
        group_summaries: Dict[str, GroupSummary] = {}
        prepared_group_dfs: Dict[str, DataFrame] = {}

        for group_name, group_df in group_dfs.items():
            # Calculate summary
            total_hours, hours_per_week = self._calculate_summary(group_df)
            group_summaries[group_name] = GroupSummary(
                total_hours=total_hours, hours_per_week=hours_per_week
            )

            # Prepare data for output
            prepared_group_dfs[group_name] = self._prepare_data_for_output(group_df)

        return prepared_group_dfs, group_summaries
