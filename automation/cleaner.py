"""Data cleaning module — reusable, composable cleaning functions."""
import pandas as pd
import numpy as np
import re
from utils.logger import setup_logger

logger = setup_logger("cleaner")


class DataCleaner:
    """Applies a configurable pipeline of cleaning steps to a DataFrame."""

    def __init__(self) -> None:
        self._log: list[str] = []

    # ------------------------------------------------------------------
    # Individual cleaning steps
    # ------------------------------------------------------------------

    def standardize_column_names(self, df: pd.DataFrame) -> pd.DataFrame:
        """Lowercase, strip, and snake_case all column names."""
        original = list(df.columns)
        df.columns = (
            df.columns.str.strip()
            .str.lower()
            .str.replace(r"[\s\-/\\]+", "_", regex=True)
            .str.replace(r"[^\w]", "", regex=True)
            .str.replace(r"_+", "_", regex=True)
            .str.strip("_")
        )
        renamed = {o: n for o, n in zip(original, df.columns) if o != n}
        if renamed:
            self._log.append(f"Renamed {len(renamed)} column(s): {renamed}")
            logger.info("Column names standardized: %s", renamed)
        return df

    def trim_spaces(self, df: pd.DataFrame) -> pd.DataFrame:
        """Strip leading/trailing whitespace from all string columns."""
        str_cols = df.select_dtypes(include="object").columns
        for col in str_cols:
            df[col] = df[col].str.strip()
        self._log.append(f"Trimmed whitespace in {len(str_cols)} string column(s)")
        logger.info("Whitespace trimmed from %d columns", len(str_cols))
        return df

    def remove_duplicates(self, df: pd.DataFrame) -> pd.DataFrame:
        """Drop fully-duplicate rows."""
        before = len(df)
        df = df.drop_duplicates().reset_index(drop=True)
        removed = before - len(df)
        self._log.append(f"Removed {removed} duplicate row(s)")
        logger.info("Duplicates removed: %d", removed)
        return df

    def auto_convert_types(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Attempt numeric and datetime conversion for object columns.
        Only converts when >80% of non-null values parse successfully.
        """
        for col in df.select_dtypes(include="object").columns:
            non_null = df[col].dropna()
            if len(non_null) == 0:
                continue
            threshold = 0.80

            # Try numeric
            num = pd.to_numeric(non_null, errors="coerce")
            if num.notna().sum() / len(non_null) >= threshold:
                df[col] = pd.to_numeric(df[col], errors="coerce")
                continue

            # Try datetime (suppress noisy warnings)
            try:
                dt = pd.to_datetime(non_null, errors="coerce", infer_datetime_format=True)
                if dt.notna().sum() / len(non_null) >= threshold:
                    df[col] = pd.to_datetime(df[col], errors="coerce", infer_datetime_format=True)
            except Exception:
                pass

        self._log.append("Auto type conversion applied to eligible columns")
        logger.info("Type conversion complete")
        return df

    def handle_missing_values(self, df: pd.DataFrame, strategy: str = "flag") -> pd.DataFrame:
        """
        Handle missing values according to the chosen strategy.

        Strategies:
          flag      – leave NaN as-is (default, no rows dropped)
          fill_mean – numeric cols → column mean; text cols → 'Unknown'
          fill_zero – all NaN → 0 / empty string
          drop      – drop any row with at least one NaN
        """
        missing_before = int(df.isnull().sum().sum())

        if strategy == "drop":
            df = df.dropna().reset_index(drop=True)
        elif strategy == "fill_mean":
            for col in df.select_dtypes(include="number").columns:
                df[col] = df[col].fillna(df[col].mean())
            for col in df.select_dtypes(include="object").columns:
                df[col] = df[col].fillna("Unknown")
        elif strategy == "fill_zero":
            df = df.fillna(0)
        # "flag" → no action

        missing_after = int(df.isnull().sum().sum())
        handled = missing_before - missing_after
        self._log.append(
            f"Missing values handled ({strategy}): {missing_before} → {missing_after} "
            f"({handled} resolved)"
        )
        logger.info("Missing values: %d → %d using strategy '%s'", missing_before, missing_after, strategy)
        return df

    def filter_sparse_rows(
        self, df: pd.DataFrame, min_fill_ratio: float = 0.40
    ) -> tuple[pd.DataFrame, pd.DataFrame]:
        """
        Separate rows that have fewer than min_fill_ratio non-null values.
        Returns (valid_df, removed_df).
        """
        threshold = max(1, int(len(df.columns) * min_fill_ratio))
        valid_mask = df.notna().sum(axis=1) >= threshold
        valid = df[valid_mask].reset_index(drop=True)
        removed = df[~valid_mask].reset_index(drop=True)
        self._log.append(
            f"Sparse-row filter (fill≥{min_fill_ratio:.0%}): "
            f"{len(valid)} kept, {len(removed)} removed"
        )
        logger.info("Sparse rows removed: %d", len(removed))
        return valid, removed

    # ------------------------------------------------------------------
    # Pipeline entry point
    # ------------------------------------------------------------------

    def clean(
        self,
        df: pd.DataFrame,
        options: dict,
    ) -> tuple[pd.DataFrame, pd.DataFrame]:
        """
        Run the full cleaning pipeline.

        Returns (cleaned_df, removed_df) where removed_df contains rows
        that were filtered out during the sparse-row step.
        """
        self._log = []

        if options.get("standardize_columns", True):
            df = self.standardize_column_names(df)
        if options.get("trim_spaces", True):
            df = self.trim_spaces(df)
        if options.get("remove_duplicates", True):
            df = self.remove_duplicates(df)
        if options.get("type_conversion", True):
            df = self.auto_convert_types(df)

        strategy = options.get("missing_strategy", "flag")
        df = self.handle_missing_values(df, strategy)

        cleaned_df, removed_df = self.filter_sparse_rows(df)
        return cleaned_df, removed_df

    def get_log(self) -> list[str]:
        return list(self._log)
