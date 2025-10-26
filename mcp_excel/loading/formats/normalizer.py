import pandas as pd
import numpy as np
from typing import Dict, Any, Optional
import re

class DataNormalizer:
    def __init__(self):
        self.locale_config: Optional[Dict[str, Any]] = None

    def set_locale(self, locale_config: Dict[str, Any]):
        self.locale_config = locale_config

    def normalize(self, df: pd.DataFrame, options: Dict[str, Any] = None) -> pd.DataFrame:
        if options is None:
            options = {}

        df = self.clean_whitespace(df, options)
        df = self.normalize_numbers(df, options)
        df = self.normalize_dates(df, options)
        df = self.handle_missing_values(df, options)
        df = self.fix_data_types(df, options)

        return df

    def clean_whitespace(self, df: pd.DataFrame, options: Dict) -> pd.DataFrame:
        for col in df.select_dtypes(include=[object]).columns:
            col_data = df[col]
            if isinstance(col_data, pd.DataFrame):
                continue

            df[col] = col_data.astype(str).str.strip()
            df[col] = df[col].str.replace(r'\s+', ' ', regex=True)
            df[col] = df[col].str.replace('\xa0', ' ')

            if not options.get('preserve_linebreaks', False):
                df[col] = df[col].str.replace(r'[\r\n]+', ' ', regex=True)

        return df

    def normalize_numbers(self, df: pd.DataFrame, options: Dict) -> pd.DataFrame:
        decimal_sep, thousands_sep = self._get_locale_separators(df, options)

        for col in df.columns:
            col_data = df[col]
            if isinstance(col_data, pd.DataFrame):
                continue

            if col_data.dtype == object:
                sample = col_data.dropna().head(100).astype(str)

                if self._looks_like_numbers(sample):
                    series = col_data.astype(str)

                    currency_symbols = self.locale_config.get('currency_symbols', []) if self.locale_config else []
                    for symbol in currency_symbols:
                        series = series.str.replace(symbol, '', regex=False)
                    series = series.str.replace(r'[$€£¥₹]', '', regex=True)

                    if thousands_sep:
                        series = series.str.replace(thousands_sep, '', regex=False)
                    if decimal_sep and decimal_sep != '.':
                        series = series.str.replace(decimal_sep, '.', regex=False)

                    series = series.str.replace(r'^\((.*)\)$', r'-\1', regex=True)
                    series = series.str.replace(' ', '')

                    df[col] = pd.to_numeric(series, errors='coerce')

        return df

    def _get_locale_separators(self, df: pd.DataFrame, options: Dict) -> tuple[str, str]:
        if self.locale_config and not self.locale_config.get('auto_detect', True):
            decimal_sep = self.locale_config.get('decimal_separator', '.')
            thousands_sep = self.locale_config.get('thousands_separator', ',')
        else:
            decimal_sep, thousands_sep = self._detect_number_format(df)

        return decimal_sep, thousands_sep

    def _detect_number_format(self, df: pd.DataFrame) -> tuple[str, str]:
        samples = []
        for col in df.select_dtypes(include=[object]).columns:
            col_data = df[col]
            if isinstance(col_data, pd.DataFrame):
                continue

            sample = col_data.dropna().head(50).astype(str)
            numeric_pattern = r'[\d.,]+\d'
            matches = sample.str.extract(f'({numeric_pattern})', expand=False).dropna()
            samples.extend(matches.tolist())

        if not samples:
            return '.', ','

        has_comma_decimal = sum(1 for s in samples if re.search(r'\d,\d{2}$', s))
        has_dot_decimal = sum(1 for s in samples if re.search(r'\d\.\d{2}$', s))
        has_dot_thousands = sum(1 for s in samples if re.search(r'\d\.\d{3}', s))
        has_comma_thousands = sum(1 for s in samples if re.search(r'\d,\d{3}', s))

        if has_comma_decimal > has_dot_decimal and has_dot_thousands > has_comma_thousands:
            return ',', '.'
        else:
            return '.', ','

    def _looks_like_numbers(self, sample: pd.Series) -> bool:
        if len(sample) == 0:
            return False

        pattern = r'^[+-]?[\d,. ]+$|^\([0-9,. ]+\)$|^[$€£¥₹][0-9,. ]+$'
        matches = sample.str.match(pattern).sum()
        return matches > len(sample) * 0.5

    def normalize_dates(self, df: pd.DataFrame, options: Dict) -> pd.DataFrame:
        for col in df.columns:
            col_data = df[col]
            if isinstance(col_data, pd.DataFrame):
                continue

            if pd.api.types.is_numeric_dtype(col_data):
                sample = col_data.dropna()
                if len(sample) > 0:
                    if (sample >= 1).all() and (sample <= 60000).all() and (sample % 1 == 0).mean() > 0.9:
                        df[col] = pd.to_datetime(col_data, unit='D', origin='1899-12-30', errors='coerce')

            elif col_data.dtype == object:
                try:
                    parsed = pd.to_datetime(col_data, errors='coerce')
                    if parsed.notna().sum() > len(col_data) * 0.5:
                        df[col] = parsed
                except:
                    pass

        return df

    def handle_missing_values(self, df: pd.DataFrame, options: Dict) -> pd.DataFrame:
        missing_values = [
            'NA', 'N/A', 'n/a', '#N/A', 'null', 'NULL', 'None', 'NONE',
            '-', '--', '---', '.', '..', '...', '?', '??', '???',
            'nan', 'NaN', 'NAN'
        ]

        missing_values.extend(options.get('custom_na_values', []))
        df = df.replace(missing_values, np.nan)

        if options.get('empty_string_as_na', True):
            df = df.replace('', np.nan)
            df = df.replace('nan', np.nan)

        if options.get('drop_empty_rows', True):
            df = df.dropna(how='all')

        if options.get('drop_empty_cols', True):
            df = df.dropna(axis=1, how='all')

        return df

    def fix_data_types(self, df: pd.DataFrame, options: Dict) -> pd.DataFrame:
        for col in df.columns:
            col_data = df[col]
            if isinstance(col_data, pd.DataFrame):
                continue

            if col_data.dtype == object:
                non_null = col_data.dropna()

                if len(non_null) == 0:
                    continue

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
                    df[col] = col_data.astype(str).str.lower().map(bool_map)
                    continue

                if not pd.api.types.is_numeric_dtype(col_data):
                    try:
                        numeric = pd.to_numeric(non_null, errors='coerce')
                        if numeric.notna().sum() > len(non_null) * 0.9:
                            df[col] = pd.to_numeric(col_data, errors='coerce')
                    except:
                        pass

        return df