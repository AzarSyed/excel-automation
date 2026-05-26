"""Shared utility functions used across the application."""
import pandas as pd
from typing import Any


def compute_analytics(df: pd.DataFrame) -> dict[str, Any]:
    """Build a comprehensive analytics dictionary from a cleaned DataFrame."""
    analytics: dict[str, Any] = {
        "total_rows": len(df),
        "total_columns": len(df.columns),
        "duplicate_rows": int(df.duplicated().sum()),
        "total_missing_values": int(df.isnull().sum().sum()),
        "missing_by_column": df.isnull().sum().to_dict(),
        "column_types": {col: str(dt) for col, dt in df.dtypes.items()},
    }

    numeric_cols = df.select_dtypes(include="number").columns.tolist()
    if numeric_cols:
        desc = df[numeric_cols].describe().round(2)
        analytics["numeric_summary"] = desc.to_dict()

    cat_cols = df.select_dtypes(include="object").columns.tolist()
    if cat_cols:
        analytics["top_categories"] = {
            col: df[col].value_counts().head(10).to_dict()
            for col in cat_cols[:6]
        }

    return analytics


def dataframe_info(df: pd.DataFrame) -> dict[str, Any]:
    """Return lightweight metadata about a DataFrame."""
    return {
        "rows": len(df),
        "columns": len(df.columns),
        "memory_mb": round(df.memory_usage(deep=True).sum() / 1024 ** 2, 3),
        "column_names": list(df.columns),
        "dtypes": {col: str(dt) for col, dt in df.dtypes.items()},
    }


def standardize_col_name(name: str) -> str:
    """Apply the same transformation DataCleaner uses for column names."""
    import re
    s = name.strip().lower()
    s = re.sub(r"[\s\-/\\]+", "_", s)
    s = re.sub(r"[^\w]", "", s)
    s = re.sub(r"_+", "_", s)
    return s.strip("_")


def format_number(n: float | int) -> str:
    """Format large numbers with comma separators."""
    if isinstance(n, float):
        return f"{n:,.2f}"
    return f"{int(n):,}"
