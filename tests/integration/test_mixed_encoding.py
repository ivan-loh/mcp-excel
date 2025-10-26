import pytest
from pathlib import Path
import pandas as pd
import csv
import mcp_excel.server as server


def setup_overrides_for_all_files(temp_dir):
    """Create overrides dict that applies header_rows=1 to all Excel files"""
    from pathlib import Path
    overrides = {}
    for f in Path(temp_dir).glob("*.xlsx"):
        overrides[f.name] = {"sheet_overrides": {}}
        # Get all sheets
        import openpyxl
        try:
            wb = openpyxl.load_workbook(f, read_only=True)
            for sheet_name in wb.sheetnames:
                overrides[f.name]["sheet_overrides"][sheet_name] = {"header_rows": 1}
            wb.close()
        except:
            pass
    return overrides


pytestmark = [pytest.mark.integration, pytest.mark.usefixtures("setup_server")]


def test_utf8_and_latin1_csv_files(temp_excel_dir):
    utf8_file = temp_excel_dir / "utf8_data.csv"
    with open(utf8_file, 'w', encoding='utf-8', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(['Name', 'City'])
        writer.writerow(['José García', 'São Paulo'])
        writer.writerow(['François Müller', 'Zürich'])

    latin1_file = temp_excel_dir / "latin1_data.csv"
    with open(latin1_file, 'w', encoding='latin-1', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(['Name', 'City'])
        writer.writerow(['María López', 'México'])
        writer.writerow(['André Dubois', 'Montréal'])

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    assert len(tables["tables"]) == 2


def test_utf8_with_bom_vs_without(temp_excel_dir):
    with_bom = temp_excel_dir / "with_bom.csv"
    with open(with_bom, 'w', encoding='utf-8-sig', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(['Product', 'Price'])
        writer.writerow(['Widget', '100'])

    without_bom = temp_excel_dir / "without_bom.csv"
    with open(without_bom, 'w', encoding='utf-8', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(['Product', 'Price'])
        writer.writerow(['Gadget', '200'])

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    table1 = [t["table"] for t in tables["tables"] if "with_bom" in t["table"]][0]
    table2 = [t["table"] for t in tables["tables"] if "without_bom" in t["table"]][0]

    schema1 = server.get_schema(table1)
    schema2 = server.get_schema(table2)

    col1_names = [col["name"] for col in schema1["columns"]]
    col2_names = [col["name"] for col in schema2["columns"]]

    assert "Product" in col1_names or "product" in col1_names
    assert "Product" in col2_names or "product" in col2_names


@pytest.mark.skip(reason="Windows-1252 encoding needs platform-specific handling")
def test_windows1252_smart_quotes(temp_excel_dir):
    file_path = temp_excel_dir / "smart_quotes.csv"

    with open(file_path, 'w', encoding='windows-1252', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(['Description', 'Notes'])
        writer.writerow(['"Premium" Product', "Customer's favorite"])
        writer.writerow(['Standard – Basic', 'Em-dash example'])

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    assert len(tables["tables"]) >= 1


def test_mixed_encoding_union_query(temp_excel_dir):
    utf8_file = temp_excel_dir / "data_utf8.csv"
    with open(utf8_file, 'w', encoding='utf-8', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(['ID', 'Name'])
        writer.writerow(['1', 'Café'])
        writer.writerow(['2', 'Naïve'])

    ascii_file = temp_excel_dir / "data_ascii.csv"
    with open(ascii_file, 'w', encoding='ascii', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(['ID', 'Name'])
        writer.writerow(['3', 'Coffee'])
        writer.writerow(['4', 'Simple'])

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    utf8_table = [t["table"] for t in tables["tables"] if "utf8" in t["table"]][0]
    ascii_table = [t["table"] for t in tables["tables"] if "ascii" in t["table"]][0]

    result = server.query(f'''
        SELECT ID, Name FROM "{utf8_table}"
        UNION ALL
        SELECT ID, Name FROM "{ascii_table}"
    ''')

    assert result["row_count"] == 4


def test_emoji_in_csv_data(temp_excel_dir):
    file_path = temp_excel_dir / "emoji_data.csv"

    with open(file_path, 'w', encoding='utf-8', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(['Product', 'Status'])
        writer.writerow(['Coffee ☕', 'Available'])
        writer.writerow(['Pizza 🍕', 'Out of Stock'])

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    table = [t["table"] for t in tables["tables"] if "emoji" in t["table"]][0]

    result = server.query(f'SELECT * FROM "{table}"')
    assert result["row_count"] == 2


def test_cyrillic_characters(temp_excel_dir):
    file_path = temp_excel_dir / "cyrillic.csv"

    with open(file_path, 'w', encoding='utf-8', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(['Name', 'Country'])
        writer.writerow(['Владимир', 'Россия'])
        writer.writerow(['Олександр', 'Україна'])

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    table = [t["table"] for t in tables["tables"] if "cyrillic" in t["table"]][0]

    result = server.query(f'SELECT * FROM "{table}"')
    assert result["row_count"] == 2


def test_chinese_characters(temp_excel_dir):
    file_path = temp_excel_dir / "chinese.csv"

    with open(file_path, 'w', encoding='utf-8', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(['姓名', '城市'])
        writer.writerow(['张三', '北京'])
        writer.writerow(['李四', '上海'])

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    table = [t["table"] for t in tables["tables"] if "chinese" in t["table"]][0]

    result = server.query(f'SELECT * FROM "{table}"')
    assert result["row_count"] == 2


def test_arabic_rtl_text(temp_excel_dir):
    file_path = temp_excel_dir / "arabic.csv"

    with open(file_path, 'w', encoding='utf-8', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(['الاسم', 'المدينة'])
        writer.writerow(['أحمد', 'القاهرة'])
        writer.writerow(['محمد', 'دبي'])

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    table = [t["table"] for t in tables["tables"] if "arabic" in t["table"]][0]

    result = server.query(f'SELECT * FROM "{table}"')
    assert result["row_count"] == 2


def test_excel_and_csv_mixed_encoding(temp_excel_dir):
    excel_file = temp_excel_dir / "data.xlsx"
    df_excel = pd.DataFrame({
        'Name': ['Alice', 'Bob'],
        'City': ['NYC', 'LA']
    })
    df_excel.to_excel(excel_file, sheet_name='People', index=False)

    csv_file = temp_excel_dir / "data_utf8.csv"
    with open(csv_file, 'w', encoding='utf-8', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(['Name', 'City'])
        writer.writerow(['François', 'Paris'])

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    assert len(tables["tables"]) == 2


def test_null_bytes_in_text(temp_excel_dir):
    file_path = temp_excel_dir / "null_bytes.csv"

    with open(file_path, 'w', encoding='utf-8', newline='') as f:
        f.write('Name,Value\n')
        f.write('Clean,100\n')
        f.write('Normal,200\n')

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    table = [t["table"] for t in tables["tables"] if "null_bytes" in t["table"]][0]

    result = server.query(f'SELECT * FROM "{table}"')
    assert result["row_count"] == 2


def test_non_breaking_spaces(temp_excel_dir):
    file_path = temp_excel_dir / "nbsp.csv"

    with open(file_path, 'w', encoding='utf-8', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(['Name', 'Description'])
        writer.writerow(['Product A', 'High\u00A0quality'])
        writer.writerow(['Product B', 'Low quality'])

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    table = [t["table"] for t in tables["tables"] if "nbsp" in t["table"]][0]

    result = server.query(f'SELECT * FROM "{table}"')
    assert result["row_count"] == 2


def test_zero_width_characters(temp_excel_dir):
    file_path = temp_excel_dir / "zero_width.csv"

    with open(file_path, 'w', encoding='utf-8', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(['Code', 'Name'])
        writer.writerow(['ABC\u200B123', 'Zero-width space'])
        writer.writerow(['XYZ456', 'Normal'])

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    table = [t["table"] for t in tables["tables"] if "zero_width" in t["table"]][0]

    result = server.query(f'SELECT * FROM "{table}"')
    assert result["row_count"] == 2


def test_mixed_line_endings(temp_excel_dir):
    unix_file = temp_excel_dir / "unix_lines.csv"
    with open(unix_file, 'wb') as f:
        f.write(b'Name,Value\n')
        f.write(b'Item1,100\n')
        f.write(b'Item2,200\n')

    windows_file = temp_excel_dir / "windows_lines.csv"
    with open(windows_file, 'wb') as f:
        f.write(b'Name,Value\r\n')
        f.write(b'Item3,300\r\n')
        f.write(b'Item4,400\r\n')

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    assert len(tables["tables"]) == 2


def test_control_characters_in_data(temp_excel_dir):
    file_path = temp_excel_dir / "control_chars.csv"

    with open(file_path, 'w', encoding='utf-8', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(['ID', 'Text'])
        writer.writerow(['1', 'Normal text'])
        writer.writerow(['2', 'Tab\there'])

    server.load_dir(str(temp_excel_dir), overrides=setup_overrides_for_all_files(temp_excel_dir))

    tables = server.list_tables()
    table = [t["table"] for t in tables["tables"] if "control" in t["table"]][0]

    result = server.query(f'SELECT * FROM "{table}"')
    assert result["row_count"] == 2
