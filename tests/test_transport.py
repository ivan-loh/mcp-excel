import pytest
import tempfile
from pathlib import Path
import pandas as pd
from fastmcp import Client
from fastmcp.client.transports import StreamableHttpTransport
from fastmcp.utilities.tests import run_server_in_process
import mcp_excel.server as server

pytestmark = [pytest.mark.integration, pytest.mark.asyncio]


def run_test_server(host: str, port: int) -> None:
    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir = Path(tmpdir)

        df = pd.DataFrame({
            "Product": ["A", "B", "C"],
            "Quantity": [10, 20, 30],
            "Price": [100.0, 200.0, 300.0]
        })
        file_path = tmpdir / "test.xlsx"
        df.to_excel(file_path, sheet_name="Data", index=False)

        server.init_server()
        server.load_dir(path=str(tmpdir), alias="test")
        server.mcp.run(transport="streamable-http", host=host, port=port)


@pytest.fixture
async def http_server():
    with run_server_in_process(run_test_server) as url:
        yield f"{url}/mcp"


@pytest.mark.asyncio
async def test_streamable_http_ping(http_server: str):
    async with Client(transport=StreamableHttpTransport(http_server)) as client:
        result = await client.ping()
        assert result is True


@pytest.mark.asyncio
async def test_streamable_http_list_tools(http_server: str):
    async with Client(transport=StreamableHttpTransport(http_server)) as client:
        tools = await client.list_tools()
        tool_names = [tool.name for tool in tools]

        assert "tool_load_dir" in tool_names
        assert "tool_query" in tool_names
        assert "tool_list_tables" in tool_names
        assert "tool_get_schema" in tool_names
        assert "tool_refresh" in tool_names


@pytest.mark.asyncio
async def test_streamable_http_call_list_tables(http_server: str):
    async with Client(transport=StreamableHttpTransport(http_server)) as client:
        result = await client.call_tool("tool_list_tables", {})

        assert result.data is not None
        tables_data = result.data
        assert "tables" in tables_data
        assert len(tables_data["tables"]) > 0


@pytest.mark.asyncio
async def test_streamable_http_call_query(http_server: str):
    async with Client(transport=StreamableHttpTransport(http_server)) as client:
        tables_result = await client.call_tool("tool_list_tables", {})
        table_name = tables_result.data["tables"][0]["table"]

        query_result = await client.call_tool("tool_query", {
            "sql": f'SELECT COUNT(*) as count FROM "{table_name}"'
        })

        assert query_result.data is not None
        assert "rows" in query_result.data
        assert query_result.data["row_count"] == 1


@pytest.mark.asyncio
async def test_streamable_http_call_get_schema(http_server: str):
    async with Client(transport=StreamableHttpTransport(http_server)) as client:
        tables_result = await client.call_tool("tool_list_tables", {})
        table_name = tables_result.data["tables"][0]["table"]

        schema_result = await client.call_tool("tool_get_schema", {
            "table": table_name
        })

        assert schema_result.data is not None
        assert "columns" in schema_result.data
        assert len(schema_result.data["columns"]) > 0
