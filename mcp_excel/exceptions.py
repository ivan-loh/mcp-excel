from typing import Optional, Any


class MCPExcelError(Exception):
    """
    Base exception for all mcp-excel errors.

    Follows MCP/JSON-RPC 2.0 error format:
    - code: Integer error code
    - message: Human-readable description
    - data: Optional diagnostic context
    """

    def __init__(
        self,
        message: str,
        code: int = -32603,
        data: Optional[dict[str, Any]] = None
    ):
        super().__init__(message)
        self.code = code
        self.message = message
        self.data = data or {}

    def to_dict(self) -> dict[str, Any]:
        """Convert to MCP error format"""
        return {
            "code": self.code,
            "message": self.message,
            "data": self.data
        }


class FileError(MCPExcelError):
    """
    File operation errors (-32001).

    Raised for:
    - File not found
    - Permission denied
    - I/O errors
    - Path validation failures
    """

    def __init__(
        self,
        message: str,
        file_path: str,
        operation: str,
        data: Optional[dict] = None
    ):
        error_data = {
            "file_path": str(file_path),
            "operation": operation
        }
        if data:
            error_data.update(data)
        super().__init__(message, code=-32001, data=error_data)


class FormatDetectionError(MCPExcelError):
    """
    Format detection and parsing errors (-32002).

    Raised when:
    - File format cannot be detected
    - All format handlers fail
    - File is corrupted
    - Unsupported file format

    Includes list of attempted formats and their errors.
    """

    def __init__(
        self,
        message: str,
        file_path: str,
        attempted_formats: list[str],
        data: Optional[dict] = None
    ):
        error_data = {
            "file_path": str(file_path),
            "attempted_formats": attempted_formats
        }
        if data:
            error_data.update(data)
        super().__init__(message, code=-32002, data=error_data)


class DataTransformError(MCPExcelError):
    """
    Data transformation and normalization errors (-32003).

    Raised when:
    - Date parsing fails unexpectedly
    - Number conversion fails
    - Type inference encounters critical errors
    - Data normalization produces invalid results
    """

    def __init__(
        self,
        message: str,
        column: Optional[str] = None,
        transformation: Optional[str] = None,
        data: Optional[dict] = None
    ):
        error_data = {}
        if column:
            error_data["column"] = column
        if transformation:
            error_data["transformation"] = transformation
        if data:
            error_data.update(data)
        super().__init__(message, code=-32003, data=error_data)


class QueryError(MCPExcelError):
    """
    Query execution errors (-32004).

    Raised when:
    - SQL syntax is invalid
    - Query execution fails
    - Query times out
    - Table/view doesn't exist during query
    """

    def __init__(
        self,
        message: str,
        sql: Optional[str] = None,
        data: Optional[dict] = None
    ):
        error_data = {}
        if sql:
            error_data["sql"] = sql[:200]
        if data:
            error_data.update(data)
        super().__init__(message, code=-32004, data=error_data)


class ResourceNotFoundError(MCPExcelError):
    """
    Resource not found errors (-32002 per MCP spec).

    Raised when:
    - Table doesn't exist
    - View doesn't exist
    - File not found in catalog
    - Alias not registered
    """

    def __init__(
        self,
        message: str,
        resource_type: str,
        resource_name: str,
        data: Optional[dict] = None
    ):
        error_data = {
            "resource_type": resource_type,
            "resource_name": resource_name
        }
        if data:
            error_data.update(data)
        super().__init__(message, code=-32002, data=error_data)


class ValidationError(MCPExcelError):
    """
    Parameter validation errors (-32602 per MCP spec).

    Raised when:
    - Invalid parameter values
    - Type mismatches
    - Value out of range
    - Missing required parameters
    """

    def __init__(
        self,
        message: str,
        parameter: str,
        expected: str,
        received: Any,
        data: Optional[dict] = None
    ):
        error_data = {
            "parameter": parameter,
            "expected": expected,
            "received": str(received)
        }
        if data:
            error_data.update(data)
        super().__init__(message, code=-32602, data=error_data)


class ExtensionError(MCPExcelError):
    """
    DuckDB extension errors (-32005).

    Raised when:
    - Extension installation fails
    - Extension loading fails
    - Extension not available
    """

    def __init__(
        self,
        message: str,
        extension_name: str,
        operation: str,
        data: Optional[dict] = None
    ):
        error_data = {
            "extension_name": extension_name,
            "operation": operation
        }
        if data:
            error_data.update(data)
        super().__init__(message, code=-32005, data=error_data)
