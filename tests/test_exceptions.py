import pytest
from mcp_excel.exceptions import (
    MCPExcelError,
    FileError,
    FormatDetectionError,
    DataTransformError,
    QueryError,
    ResourceNotFoundError,
    ValidationError,
    ExtensionError
)


class TestMCPExcelError:
    def test_base_exception_default_code(self):
        err = MCPExcelError("test error")
        assert err.code == -32603
        assert err.message == "test error"
        assert err.data == {}

    def test_base_exception_custom_code(self):
        err = MCPExcelError("test error", code=-32001)
        assert err.code == -32001

    def test_base_exception_with_data(self):
        err = MCPExcelError("test error", data={"key": "value"})
        assert err.data["key"] == "value"

    def test_to_dict(self):
        err = MCPExcelError("test error", code=-32603, data={"key": "value"})
        result = err.to_dict()
        assert result["code"] == -32603
        assert result["message"] == "test error"
        assert result["data"]["key"] == "value"


class TestFileError:
    def test_file_error_code(self):
        err = FileError("Permission denied", "/path/to/file", "read")
        assert err.code == -32001

    def test_file_error_data_structure(self):
        err = FileError("Permission denied", "/path/to/file", "read")
        assert err.data["file_path"] == "/path/to/file"
        assert err.data["operation"] == "read"

    def test_file_error_with_additional_data(self):
        err = FileError(
            "Permission denied",
            "/path/to/file",
            "read",
            data={"error": "EACCES"}
        )
        assert err.data["file_path"] == "/path/to/file"
        assert err.data["operation"] == "read"
        assert err.data["error"] == "EACCES"


class TestFormatDetectionError:
    def test_format_detection_error_code(self):
        err = FormatDetectionError(
            "Failed to detect format",
            "/path/to/file",
            ["xlsx", "csv"]
        )
        assert err.code == -32002

    def test_format_detection_error_data_structure(self):
        err = FormatDetectionError(
            "Failed to detect format",
            "/path/to/file",
            ["xlsx", "csv", "json"]
        )
        assert err.data["file_path"] == "/path/to/file"
        assert err.data["attempted_formats"] == ["xlsx", "csv", "json"]

    def test_format_detection_error_with_errors_list(self):
        errors = ["xlsx: BadZipFile", "csv: UnicodeError"]
        err = FormatDetectionError(
            "All handlers failed",
            "/path/to/file",
            ["xlsx", "csv"],
            data={"errors": errors}
        )
        assert err.data["errors"] == errors


class TestDataTransformError:
    def test_data_transform_error_code(self):
        err = DataTransformError("Transform failed")
        assert err.code == -32003

    def test_data_transform_error_with_column(self):
        err = DataTransformError("Date parsing failed", column="date_col")
        assert err.data["column"] == "date_col"

    def test_data_transform_error_with_transformation(self):
        err = DataTransformError(
            "Conversion failed",
            column="amount",
            transformation="to_numeric"
        )
        assert err.data["column"] == "amount"
        assert err.data["transformation"] == "to_numeric"


class TestQueryError:
    def test_query_error_code(self):
        err = QueryError("Query failed")
        assert err.code == -32004

    def test_query_error_with_sql(self):
        sql = "SELECT * FROM nonexistent_table"
        err = QueryError("Query failed", sql=sql)
        assert err.data["sql"] == sql

    def test_query_error_truncates_long_sql(self):
        long_sql = "SELECT * FROM table WHERE " + "x " * 100
        err = QueryError("Query failed", sql=long_sql)
        assert len(err.data["sql"]) <= 200


class TestResourceNotFoundError:
    def test_resource_not_found_error_code(self):
        err = ResourceNotFoundError("Not found", "table", "my_table")
        assert err.code == -32002

    def test_resource_not_found_error_data_structure(self):
        err = ResourceNotFoundError(
            "Table not found",
            "table",
            "my_table"
        )
        assert err.data["resource_type"] == "table"
        assert err.data["resource_name"] == "my_table"


class TestValidationError:
    def test_validation_error_code(self):
        err = ValidationError(
            "Invalid parameter",
            "max_rows",
            "integer",
            "not a number"
        )
        assert err.code == -32602

    def test_validation_error_data_structure(self):
        err = ValidationError(
            "Invalid parameter",
            "max_rows",
            "integer",
            "abc"
        )
        assert err.data["parameter"] == "max_rows"
        assert err.data["expected"] == "integer"
        assert err.data["received"] == "abc"


class TestExtensionError:
    def test_extension_error_code(self):
        err = ExtensionError(
            "Extension failed",
            "excel",
            "INSTALL"
        )
        assert err.code == -32005

    def test_extension_error_data_structure(self):
        err = ExtensionError(
            "Extension failed",
            "excel",
            "LOAD"
        )
        assert err.data["extension_name"] == "excel"
        assert err.data["operation"] == "LOAD"


class TestExceptionInheritance:
    def test_all_inherit_from_base(self):
        assert issubclass(FileError, MCPExcelError)
        assert issubclass(FormatDetectionError, MCPExcelError)
        assert issubclass(DataTransformError, MCPExcelError)
        assert issubclass(QueryError, MCPExcelError)
        assert issubclass(ResourceNotFoundError, MCPExcelError)
        assert issubclass(ValidationError, MCPExcelError)
        assert issubclass(ExtensionError, MCPExcelError)

    def test_all_inherit_from_exception(self):
        assert issubclass(MCPExcelError, Exception)
        assert issubclass(FileError, Exception)
        assert issubclass(FormatDetectionError, Exception)


class TestExceptionRaising:
    def test_can_raise_and_catch_file_error(self):
        with pytest.raises(FileError) as exc_info:
            raise FileError("Test error", "/path", "read")
        assert exc_info.value.code == -32001

    def test_can_catch_as_base_exception(self):
        with pytest.raises(MCPExcelError):
            raise FileError("Test error", "/path", "read")

    def test_exception_message(self):
        err = FileError("Permission denied", "/path/to/file", "read")
        assert str(err) == "Permission denied"
