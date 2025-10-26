import re
import pandas as pd
import numpy as np
from typing import Optional


class SemanticTypeInference:
    """
    Infer column data types based on semantic meaning from column names.
    Applied before pandas auto-detection to prevent type contamination.
    """

    NUMERIC_PATTERNS = [
        r'.*amount.*', r'.*price.*', r'.*cost.*', r'.*revenue.*',
        r'.*total.*', r'.*sum.*', r'.*value.*', r'.*balance.*',
        r'.*qty.*', r'.*quantity.*', r'.*count.*', r'.*sales.*',
        r'.*profit.*', r'.*expense.*', r'.*fee.*', r'.*charge.*',
        r'.*rate.*', r'.*percent.*', r'.*margin.*'
    ]

    TEXT_ID_PATTERNS = [
        r'^id$', r'.*_id$', r'.*id$',
        r'.*number$', r'.*code$', r'.*ref$', r'.*key$',
        r'^sku$', r'.*sku$',
        r'.*zip.*', r'.*postal.*',
        r'account.*code.*', r'.*batch.*',
        r'tracking.*', r'invoice.*number.*',
        r'order.*number.*', r'customer.*id.*',
        r'employee.*id.*'
    ]

    DATE_PATTERNS = [
        r'.*date.*', r'.*time.*', r'.*timestamp.*',
        r'.*when.*', r'.*created.*', r'.*modified.*',
        r'.*due.*', r'.*start.*', r'.*end.*',
        r'.*period.*', r'.*year.*', r'.*month.*'
    ]

    def infer_type_from_name(self, col_name: str) -> Optional[str]:
        """
        Infer appropriate SQL type from column name.

        Returns:
            'DECIMAL' for numeric amounts/prices
            'VARCHAR' for IDs/codes that might have leading zeros
            'TIMESTAMP' for date/time columns
            None if no pattern matches
        """
        col_lower = str(col_name).lower().strip()

        if any(re.match(pattern, col_lower) for pattern in self.DATE_PATTERNS):
            return 'TIMESTAMP'

        if any(re.match(pattern, col_lower) for pattern in self.TEXT_ID_PATTERNS):
            return 'VARCHAR'

        if any(re.match(pattern, col_lower) for pattern in self.NUMERIC_PATTERNS):
            return 'DECIMAL'

        return None

    def generate_type_hints(self, df: pd.DataFrame) -> dict[str, str]:
        """
        Generate type hints dictionary for all columns based on names.

        Args:
            df: DataFrame with column names to analyze

        Returns:
            Dictionary mapping column names to SQL types
        """
        type_hints = {}
        for col in df.columns:
            suggested_type = self.infer_type_from_name(col)
            if suggested_type:
                type_hints[col] = suggested_type
        return type_hints

    def apply_semantic_types(self, df: pd.DataFrame, override_hints: dict[str, str] = None) -> pd.DataFrame:
        """
        Apply semantic type conversions to DataFrame.
        Prevents pandas type inference contamination.

        Args:
            df: DataFrame to process
            override_hints: Explicit type hints from user config (takes precedence)

        Returns:
            DataFrame with corrected types
        """
        auto_hints = self.generate_type_hints(df)

        if override_hints:
            auto_hints.update(override_hints)

        for col, sql_type in auto_hints.items():
            if col not in df.columns:
                continue

            try:
                if sql_type in ['DECIMAL', 'DOUBLE', 'FLOAT', 'NUMERIC']:
                    df[col] = pd.to_numeric(df[col], errors='coerce')

                elif sql_type in ['VARCHAR', 'TEXT', 'STRING']:
                    df[col] = df[col].astype(str)
                    df[col] = df[col].replace('nan', np.nan)
                    df[col] = df[col].replace('None', np.nan)

                elif sql_type in ['INTEGER', 'BIGINT', 'INT']:
                    df[col] = pd.to_numeric(df[col], errors='coerce')
                    df[col] = df[col].astype('Int64')

                elif sql_type in ['TIMESTAMP', 'DATE', 'DATETIME']:
                    if df[col].dtype != 'datetime64[ns]':
                        df[col] = pd.to_datetime(df[col], errors='coerce')

            except Exception:
                pass

        return df

    def detect_type_contamination(self, df: pd.DataFrame) -> list[dict]:
        """
        Detect suspicious type inferences where column name doesn't match type.

        Returns:
            List of warnings with column, expected type, actual type
        """
        warnings = []

        for col in df.columns:
            col_lower = str(col).lower()
            dtype = str(df[col].dtype)

            if any(kw in col_lower for kw in ['amount', 'price', 'total', 'revenue', 'cost']):
                if 'datetime' in dtype or 'timestamp' in dtype:
                    warnings.append({
                        'column': col,
                        'expected': 'numeric',
                        'actual': dtype,
                        'suggestion': f'type_hints: {{"{col}": "DECIMAL"}}'
                    })

            if col_lower.endswith('id') or 'code' in col_lower:
                if 'datetime' in dtype:
                    warnings.append({
                        'column': col,
                        'expected': 'text',
                        'actual': dtype,
                        'suggestion': f'type_hints: {{"{col}": "VARCHAR"}}'
                    })

        return warnings
