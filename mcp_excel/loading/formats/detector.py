import zipfile
from pathlib import Path
from typing import Optional, Dict, Any
from dataclasses import dataclass

from ...exceptions import FileError
from ...utils import log

@dataclass
class FormatInfo:
    format_type: str
    version: Optional[str] = None
    encoding: str = 'utf-8'
    has_macros: bool = False
    is_encrypted: bool = False
    is_compressed: bool = False
    confidence: float = 0.0
    metadata: Dict[str, Any] = None

    def __post_init__(self):
        if self.metadata is None:
            self.metadata = {}

class FormatDetector:
    SIGNATURES = {
        b'PK\x03\x04': 'zip_based',
        b'PK\x05\x06': 'zip_based',
        b'PK\x07\x08': 'zip_based',
        b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1': 'ole2',
        b'\x09\x08\x10\x00\x00\x06\x05\x00': 'xls_biff5',
        b'\x09\x08\x05\x00\x00\x06\x03\x00': 'xls_biff3',
        b'\x09\x08\x04\x00\x00\x06\x02\x00': 'xls_biff2',
    }

    def detect(self, file_path: Path) -> FormatInfo:
        if not file_path.exists():
            return FormatInfo(format_type='unknown', confidence=0.0)

        with open(file_path, 'rb') as f:
            header = f.read(8192)

        for signature, format_family in self.SIGNATURES.items():
            if header.startswith(signature):
                return self._analyze_format_family(file_path, format_family, header)

        if self._is_text_based(header):
            return self._analyze_text_format(file_path, header)

        extension = file_path.suffix.lower()
        if extension == '.xlsx':
            return FormatInfo(format_type='xlsx', confidence=0.5)
        elif extension == '.xls':
            return FormatInfo(format_type='xls', confidence=0.5)
        elif extension == '.csv':
            return FormatInfo(format_type='csv', encoding='utf-8', confidence=0.5)
        elif extension == '.tsv':
            return FormatInfo(format_type='tsv', encoding='utf-8', confidence=0.5)

        return FormatInfo(format_type='unknown', confidence=0.0)

    def _analyze_format_family(self, file_path: Path, family: str, header: bytes) -> FormatInfo:
        if family == 'zip_based':
            return self._analyze_zip_format(file_path)
        elif family == 'ole2' or family.startswith('xls'):
            return FormatInfo(
                format_type='xls',
                version='Excel 97-2003',
                encoding='cp1252',
                confidence=0.95
            )

        return FormatInfo(format_type='unknown', confidence=0.0)

    def _analyze_zip_format(self, file_path: Path) -> FormatInfo:
        try:
            with zipfile.ZipFile(file_path, 'r') as zf:
                filelist = zf.namelist()

                if 'xl/workbook.xml' in filelist:
                    has_macros = 'xl/vbaProject.bin' in filelist
                    format_type = 'xlsm' if has_macros else 'xlsx'
                    is_encrypted = 'EncryptionInfo' in filelist

                    return FormatInfo(
                        format_type=format_type,
                        version='Office Open XML',
                        encoding='utf-8',
                        has_macros=has_macros,
                        is_encrypted=is_encrypted,
                        is_compressed=True,
                        confidence=0.95,
                        metadata={'files': len(filelist)}
                    )

                elif 'xl/workbook.bin' in filelist:
                    return FormatInfo(
                        format_type='xlsb',
                        version='Excel Binary',
                        encoding='binary',
                        has_macros='xl/vbaProject.bin' in filelist,
                        is_encrypted='EncryptionInfo' in filelist,
                        is_compressed=True,
                        confidence=0.95,
                        metadata={'files': len(filelist)}
                    )

        except FileNotFoundError as e:
            log.error("format_detection_file_not_found", file=str(file_path))
            raise FileError(
                f"File not found: {file_path}",
                file_path=str(file_path),
                operation="detect_format"
            )
        except PermissionError as e:
            log.error("format_detection_permission_denied", file=str(file_path))
            raise FileError(
                f"Permission denied: {file_path}",
                file_path=str(file_path),
                operation="detect_format"
            )
        except Exception as e:
            log.warn("format_detection_unexpected", file=str(file_path), error=str(e))

        return FormatInfo(format_type='unknown', confidence=0.0)

    def _is_text_based(self, header: bytes) -> bool:
        try:
            text_chars = bytes(range(32, 127)) + b'\n\r\t'
            text_ratio = sum(1 for byte in header[:1000] if byte in text_chars) / min(1000, len(header))
            return text_ratio > 0.8
        except Exception as e:
            log.debug("text_detection_failed", error=str(e))
            return False

    def _analyze_text_format(self, file_path: Path, header: bytes) -> FormatInfo:
        try:
            text_sample = header.decode('utf-8', errors='ignore')

            lines = text_sample.split('\n')[:10]
            if len(lines) >= 2:
                tab_counts = [line.count('\t') for line in lines if line]
                if tab_counts and all(c == tab_counts[0] for c in tab_counts) and tab_counts[0] > 0:
                    return FormatInfo(format_type='tsv', encoding='utf-8', confidence=0.8)

                comma_counts = [line.count(',') for line in lines if line]
                if comma_counts and all(c == comma_counts[0] for c in comma_counts) and comma_counts[0] > 0:
                    return FormatInfo(format_type='csv', encoding='utf-8', confidence=0.8)

                semicolon_counts = [line.count(';') for line in lines if line]
                if semicolon_counts and all(c == semicolon_counts[0] for c in semicolon_counts) and semicolon_counts[0] > 0:
                    return FormatInfo(format_type='csv', encoding='utf-8', confidence=0.8)
        except Exception as e:
            log.debug("text_format_analysis_failed", error=str(e))

        return FormatInfo(format_type='unknown', confidence=0.0)