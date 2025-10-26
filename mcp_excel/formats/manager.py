from pathlib import Path
import pandas as pd
from typing import Optional, Dict, Any, List
import tempfile

from .detector import FormatDetector, FormatInfo
from .handlers import FormatHandler, XLSXHandler, XLSHandler, CSVHandler, ParseOptions
from .normalizer import DataNormalizer

class FormatManager:
    def __init__(self, cache_dir: Optional[Path] = None):
        self.detector = FormatDetector()
        self.handlers: List[FormatHandler] = [
            XLSXHandler(),
            XLSHandler(),
            CSVHandler(),
        ]
        self.normalizer = DataNormalizer()
        self.cache_dir = cache_dir

    def load_file(
        self,
        file_path: Path,
        sheet: Optional[str] = None,
        options: Optional[Dict[str, Any]] = None
    ) -> pd.DataFrame:
        # Detect format
        format_info = self.detector.detect(file_path)

        if format_info.format_type == 'unknown':
            # Try to load as Excel anyway
            format_info.format_type = 'xlsx'

        # Find appropriate handler
        handler = self._get_handler(format_info.format_type)
        if not handler:
            raise ValueError(f"No handler available for format {format_info.format_type}")

        # Validate file
        is_valid, error = handler.validate(file_path)
        if not is_valid:
            # Try as different format
            if format_info.format_type == 'xlsx':
                # Try as XLS
                xls_handler = XLSHandler()
                is_valid, error = xls_handler.validate(file_path)
                if is_valid:
                    handler = xls_handler
                    format_info.format_type = 'xls'

            if not is_valid:
                raise ValueError(f"File validation failed: {error}")

        # Create parse options
        parse_options = self._create_parse_options(format_info, options)

        # Parse file
        try:
            df = handler.parse(file_path, sheet, parse_options)
        except Exception as e:
            # Try alternative handler
            if format_info.format_type == 'xlsx':
                # Try pandas directly
                df = pd.read_excel(
                    file_path,
                    sheet_name=sheet,
                    skiprows=parse_options.skip_rows,
                    skipfooter=parse_options.skip_footer,
                    na_values=parse_options.na_values
                )
            else:
                raise e

        # Normalize data
        if options and options.get('normalize', True):
            df = self.normalizer.normalize(df, options)

        return df

    def get_sheets(self, file_path: Path) -> List[str]:
        format_info = self.detector.detect(file_path)

        if format_info.format_type == 'unknown':
            format_info.format_type = 'xlsx'

        handler = self._get_handler(format_info.format_type)
        if handler:
            try:
                return handler.get_sheets(file_path)
            except:
                # Try alternative handlers
                for h in self.handlers:
                    if h != handler:
                        try:
                            return h.get_sheets(file_path)
                        except:
                            continue

        return ['Sheet1']

    def _get_handler(self, format_type: str) -> Optional[FormatHandler]:
        for handler in self.handlers:
            if handler.can_handle(format_type):
                return handler
        return None

    def _create_parse_options(
        self,
        format_info: FormatInfo,
        user_options: Optional[Dict[str, Any]]
    ) -> ParseOptions:
        options = ParseOptions()

        if format_info.encoding:
            options.encoding = format_info.encoding

        if user_options:
            for key, value in user_options.items():
                if hasattr(options, key):
                    setattr(options, key, value)

        return options