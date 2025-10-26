from pathlib import Path
import pandas as pd
from typing import Optional, Dict, Any, List

from .detector import FormatDetector, FormatInfo
from .handlers import FormatHandler, XLSXHandler, XLSHandler, CSVHandler, ParseOptions
from .normalizer import DataNormalizer
from ...exceptions import FormatDetectionError, FileError
from ...utils import log

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
        options: Optional[Dict[str, Any]] = None,
        semantic_hints: Optional[Dict[str, str]] = None
    ) -> pd.DataFrame:
        format_info = self.detector.detect(file_path)

        if format_info.format_type == 'unknown':
            format_info.format_type = 'xlsx'

        handler = self._get_handler(format_info.format_type)
        if not handler:
            raise ValueError(f"No handler available for format {format_info.format_type}")

        is_valid, error = handler.validate(file_path)
        if not is_valid:
            if format_info.format_type == 'xlsx':
                xls_handler = XLSHandler()
                is_valid, error = xls_handler.validate(file_path)
                if is_valid:
                    handler = xls_handler
                    format_info.format_type = 'xls'

            if not is_valid:
                raise ValueError(f"File validation failed: {error}")

        parse_options = self._create_parse_options(format_info, options)

        try:
            df = handler.parse(file_path, sheet, parse_options)
        except Exception as e:
            if format_info.format_type == 'xlsx':
                df = pd.read_excel(
                    file_path,
                    sheet_name=sheet,
                    skiprows=parse_options.skip_rows,
                    skipfooter=parse_options.skip_footer,
                    na_values=parse_options.na_values
                )
            else:
                raise e

        if options and options.get('normalize', True):
            df = self.normalizer.normalize(df, options, semantic_hints)

        return df

    def get_sheets(self, file_path: Path) -> List[str]:
        format_info = self.detector.detect(file_path)

        if format_info.format_type == 'unknown':
            format_info.format_type = 'xlsx'

        handler = self._get_handler(format_info.format_type)
        if not handler:
            log.warn("no_handler_for_format", file=str(file_path), format=format_info.format_type)
            return ['Sheet1']

        errors_encountered = []
        handlers_tried = []

        try:
            sheets = handler.get_sheets(file_path)
            log.info("sheets_read_success", file=str(file_path), handler=handler.__class__.__name__, sheets=len(sheets))
            return sheets
        except PermissionError as e:
            log.error("permission_denied", file=str(file_path), handler=handler.__class__.__name__, error=str(e))
            raise FileError(
                f"Permission denied reading {file_path}",
                file_path=str(file_path),
                operation="get_sheets",
                data={"error": str(e), "handler": handler.__class__.__name__}
            )
        except FileNotFoundError as e:
            log.error("file_not_found", file=str(file_path), error=str(e))
            raise FileError(
                f"File not found: {file_path}",
                file_path=str(file_path),
                operation="get_sheets",
                data={"error": str(e)}
            )
        except MemoryError as e:
            log.error("memory_error", file=str(file_path), handler=handler.__class__.__name__, error=str(e))
            raise FileError(
                f"Insufficient memory to read {file_path}",
                file_path=str(file_path),
                operation="get_sheets",
                data={"error": "MemoryError", "suggestion": "File might be too large for available memory"}
            )
        except Exception as e:
            error_detail = f"{handler.__class__.__name__}: {type(e).__name__}: {str(e)}"
            errors_encountered.append(error_detail)
            handlers_tried.append(handler.__class__.__name__)
            log.debug("handler_failed_trying_fallback", file=str(file_path), handler=handler.__class__.__name__, error=str(e), error_type=type(e).__name__)

        for h in self.handlers:
            if h != handler:
                try:
                    sheets = h.get_sheets(file_path)
                    log.info("fallback_handler_success", file=str(file_path), failed_handler=handler.__class__.__name__, success_handler=h.__class__.__name__, sheets=len(sheets))
                    return sheets
                except (PermissionError, FileNotFoundError, MemoryError):
                    raise
                except Exception as handler_error:
                    error_detail = f"{h.__class__.__name__}: {type(handler_error).__name__}: {str(handler_error)}"
                    errors_encountered.append(error_detail)
                    handlers_tried.append(h.__class__.__name__)
                    log.debug("fallback_handler_failed", file=str(file_path), handler=h.__class__.__name__, error=str(handler_error), error_type=type(handler_error).__name__)
                    continue

        log.error("all_handlers_failed", file=str(file_path), handlers_tried=handlers_tried, error_count=len(errors_encountered))
        raise FormatDetectionError(
            f"Failed to read sheets from {file_path.name}. Tried {len(handlers_tried)} handlers but all failed. File might be corrupted, encrypted, or in an unsupported format.",
            file_path=str(file_path),
            attempted_formats=handlers_tried,
            data={"errors": errors_encountered, "suggestion": "Check if the file opens correctly in Excel/LibreOffice. If it's password-protected, remove the password first."}
        )

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