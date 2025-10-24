import pandas as pd
import numpy as np
from typing import Dict, Any
import re

class DataNormalizer:
    def normalize(self, df: pd.DataFrame, options: Dict[str, Any] = None) -> pd.DataFrame:
        if options is None:
            options = {}

        # Apply normalization steps
        df = self.clean_whitespace(df, options)
        df = self.normalize_numbers(df, options)
        df = self.normalize_dates(df, options)
        df = self.handle_missing_values(df, options)
        df = self.fix_data_types(df, options)

        return df

    def clean_whitespace(self, df: pd.DataFrame, options: Dict) -> pd.DataFrame:
        for col in df.select_dtypes(include=[object]).columns:
            # Convert to string and strip whitespace
            df[col] = df[col].astype(str).str.strip()

            # Replace multiple spaces with single space
            df[col] = df[col].str.replace(r'\s+', ' ', regex=True)

            # Remove non-breaking spaces
            df[col] = df[col].str.replace('\xa0', ' ')

            # Handle line breaks
            if not options.get('preserve_linebreaks', False):
                df[col] = df[col].str.replace(r'[\r\n]+', ' ', regex=True)

        return df

    def normalize_numbers(self, df: pd.DataFrame, options: Dict) -> pd.DataFrame:
        for col in df.columns:
            if df[col].dtype == object:
                # Check if column contains numeric strings
                sample = df[col].dropna().head(100).astype(str)

                if self._looks_like_numbers(sample):
                    # Clean numeric strings
                    series = df[col].astype(str)

                    # Remove currency symbols
                    series = series.str.replace(r'[$€£¥₹]', '', regex=True)

                    # Handle thousand separators
                    # Assume US format by default (1,234.56)
                    series = series.str.replace(',', '')

                    # Handle parentheses for negative numbers
                    series = series.str.replace(r'^\((.*)\)$', r'-\1', regex=True)

                    # Remove spaces
                    series = series.str.replace(' ', '')

                    # Convert to numeric
                    df[col] = pd.to_numeric(series, errors='coerce')

        return df

    def _looks_like_numbers(self, sample: pd.Series) -> bool:
        if len(sample) == 0:
            return False

        # Pattern for numbers with optional formatting
        pattern = r'^[+-]?[\d,. ]+$|^\([0-9,. ]+\)$|^[$€£¥₹][0-9,. ]+$'
        matches = sample.str.match(pattern).sum()
        return matches > len(sample) * 0.5

    def normalize_dates(self, df: pd.DataFrame, options: Dict) -> pd.DataFrame:
        for col in df.columns:
            # Handle Excel date serial numbers
            if pd.api.types.is_numeric_dtype(df[col]):
                sample = df[col].dropna()
                if len(sample) > 0:
                    # Check if values are in Excel date range (1 to ~60000)
                    if (sample >= 1).all() and (sample <= 60000).all() and (sample % 1 == 0).mean() > 0.9:
                        # Likely Excel dates
                        df[col] = pd.to_datetime(df[col], unit='D', origin='1899-12-30', errors='coerce')

            # Try to parse text dates
            elif df[col].dtype == object:
                try:
                    # First try pandas auto-detection
                    parsed = pd.to_datetime(df[col], errors='coerce')
                    if parsed.notna().sum() > len(df[col]) * 0.5:  # If >50% parsed successfully
                        df[col] = parsed
                except:
                    pass

        return df

    def handle_missing_values(self, df: pd.DataFrame, options: Dict) -> pd.DataFrame:
        # Common missing value representations
        missing_values = [
            'NA', 'N/A', 'n/a', '#N/A', 'null', 'NULL', 'None', 'NONE',
            '-', '--', '---', '.', '..', '...', '?', '??', '???',
            'nan', 'NaN', 'NAN'
        ]

        # Add custom missing values from options
        missing_values.extend(options.get('custom_na_values', []))

        # Replace with NaN
        df = df.replace(missing_values, np.nan)

        # Handle empty strings
        if options.get('empty_string_as_na', True):
            df = df.replace('', np.nan)
            df = df.replace('nan', np.nan)  # String 'nan' from str conversion

        # Remove completely empty rows/columns
        if options.get('drop_empty_rows', True):
            df = df.dropna(how='all')

        if options.get('drop_empty_cols', True):
            df = df.dropna(axis=1, how='all')

        return df

    def fix_data_types(self, df: pd.DataFrame, options: Dict) -> pd.DataFrame:
        for col in df.columns:
            if df[col].dtype == object:
                # Try to infer better type
                non_null = df[col].dropna()

                if len(non_null) == 0:
                    continue

                # Check for boolean
                unique_lower = non_null.astype(str).str.lower().unique()
                if len(unique_lower) <= 4 and set(unique_lower).issubset(
                    {'true', 'false', 'yes', 'no', '1', '0', 't', 'f', 'y', 'n'}
                ):
                    bool_map = {
                        'true': True, 'false': False,
                        'yes': True, 'no': False,
                        '1': True, '0': False,
                        't': True, 'f': False,
                        'y': True, 'n': False
                    }
                    df[col] = df[col].astype(str).str.lower().map(bool_map)
                    continue

                # Try numeric conversion if not already done
                if not pd.api.types.is_numeric_dtype(df[col]):
                    try:
                        numeric = pd.to_numeric(non_null, errors='coerce')
                        if numeric.notna().sum() > len(non_null) * 0.9:  # 90% successful
                            df[col] = pd.to_numeric(df[col], errors='coerce')
                    except:
                        pass

        return df