"""Data validation module — rule-based validation with structured reporting."""
import re
import pandas as pd
from utils.logger import setup_logger

logger = setup_logger("validator")

EMAIL_RE = re.compile(r"^[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}$")


class DataValidator:
    """Validates a DataFrame against a configurable set of rules."""

    def __init__(self) -> None:
        self._results: dict = {}

    # ------------------------------------------------------------------
    # Individual checks (each returns a boolean mask — True = invalid)
    # ------------------------------------------------------------------

    def check_missing_columns(
        self, df: pd.DataFrame, required: list[str]
    ) -> list[str]:
        """Return a list of required column names absent from df."""
        missing = [c for c in required if c not in df.columns]
        if missing:
            logger.warning("Missing required columns: %s", missing)
        return missing

    def invalid_email_mask(
        self, df: pd.DataFrame, email_cols: list[str]
    ) -> pd.Series:
        """Return a boolean mask — True where an email value is malformed."""
        mask = pd.Series(False, index=df.index)
        for col in email_cols:
            if col not in df.columns:
                continue
            col_mask = (
                df[col].notna()
                & ~df[col].astype(str).str.match(EMAIL_RE.pattern)
            )
            mask = mask | col_mask
        return mask

    def invalid_numeric_mask(
        self, df: pd.DataFrame, numeric_cols: list[str]
    ) -> pd.Series:
        """Return a boolean mask — True where a numeric column contains non-numeric text."""
        mask = pd.Series(False, index=df.index)
        for col in numeric_cols:
            if col not in df.columns:
                continue
            if pd.api.types.is_numeric_dtype(df[col]):
                continue
            converted = pd.to_numeric(df[col], errors="coerce")
            col_mask = df[col].notna() & converted.isna()
            mask = mask | col_mask
        return mask

    # ------------------------------------------------------------------
    # Summary + invalid-row extraction
    # ------------------------------------------------------------------

    def generate_summary(self, df: pd.DataFrame, config: dict) -> dict:
        """
        Produce a validation summary dict.

        Config keys:
          required_columns – list[str]
          email_columns    – list[str]
          numeric_columns  – list[str]
        """
        required = config.get("required_columns", [])
        email_cols = config.get("email_columns", [])
        numeric_cols = config.get("numeric_columns", [])

        missing_cols = self.check_missing_columns(df, required)
        email_mask = self.invalid_email_mask(df, email_cols)
        numeric_mask = self.invalid_numeric_mask(df, numeric_cols)
        combined_invalid = email_mask | numeric_mask

        summary = {
            "total_rows": len(df),
            "missing_required_columns": missing_cols,
            "invalid_email_count": int(email_mask.sum()),
            "invalid_numeric_count": int(numeric_mask.sum()),
            "total_invalid_rows": int(combined_invalid.sum()),
            "duplicate_count": int(df.duplicated().sum()),
            "missing_value_counts": df.isnull().sum().to_dict(),
            "completeness_pct": round(
                (1 - df.isnull().sum().sum() / max(df.size, 1)) * 100, 1
            ),
        }
        self._results = summary
        logger.info(
            "Validation done — invalid rows: %d, invalid emails: %d, invalid numeric: %d",
            summary["total_invalid_rows"],
            summary["invalid_email_count"],
            summary["invalid_numeric_count"],
        )
        return summary

    def get_invalid_rows(self, df: pd.DataFrame, config: dict) -> pd.DataFrame:
        """Return a DataFrame containing only rows that failed any validation rule."""
        email_cols = config.get("email_columns", [])
        numeric_cols = config.get("numeric_columns", [])
        mask = self.invalid_email_mask(df, email_cols) | self.invalid_numeric_mask(
            df, numeric_cols
        )
        invalid = df[mask].copy().reset_index(drop=True)
        invalid.insert(0, "_validation_issue", "")

        # Annotate why each row failed
        email_mask = self.invalid_email_mask(df, email_cols)[mask].values
        numeric_mask = self.invalid_numeric_mask(df, numeric_cols)[mask].values
        reasons = []
        for em, nm in zip(email_mask, numeric_mask):
            parts = []
            if em:
                parts.append("invalid email")
            if nm:
                parts.append("invalid numeric")
            reasons.append(", ".join(parts))
        invalid["_validation_issue"] = reasons
        return invalid

    def get_results(self) -> dict:
        return dict(self._results)
