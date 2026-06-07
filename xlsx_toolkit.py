import csv
import io
import re
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, List, Sequence
from xml.etree import ElementTree as ET


XLSX_NS = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
XML_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
OFFICE_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


@dataclass
class SheetData:
    name: str
    rows: List[List[str]]


@dataclass
class WorkbookData:
    name: str
    sheets: List[SheetData]


def load_uploaded_workbook(file_storage) -> WorkbookData:
    filename = file_storage.filename or "uploaded-file"
    suffix = Path(filename).suffix.lower()
    raw_bytes = file_storage.read()
    file_storage.stream.seek(0)

    if suffix == ".csv":
        rows = read_csv_rows(raw_bytes)
        sheet_name = Path(filename).stem or "Sheet1"
        return WorkbookData(name=filename, sheets=[SheetData(name=sheet_name, rows=rows)])

    if suffix == ".xlsx":
        return read_xlsx_bytes(raw_bytes, filename)

    raise ValueError(f"暂不支持的文件类型: {suffix or 'unknown'}，请上传 .xlsx 或 .csv 文件。")


def read_csv_rows(raw_bytes: bytes) -> List[List[str]]:
    text = None
    for encoding in ("utf-8-sig", "utf-8", "gb18030", "gbk", "latin-1"):
        try:
            text = raw_bytes.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        raise ValueError("CSV 文件无法解码。")

    return [[clean_cell(cell) for cell in row] for row in csv.reader(io.StringIO(text))]


def read_xlsx_bytes(raw_bytes: bytes, filename: str = "workbook.xlsx") -> WorkbookData:
    with zipfile.ZipFile(io.BytesIO(raw_bytes)) as workbook:
        shared_strings = _read_shared_strings(workbook)
        sheets = []
        for sheet_name, target in _read_sheet_targets(workbook):
            try:
                rows = _read_sheet_rows(workbook, target, shared_strings)
            except KeyError:
                continue
            if rows:
                sheets.append(SheetData(name=sheet_name, rows=rows))
    if not sheets:
        raise ValueError("Excel 文件中没有读取到有效工作表。")
    return WorkbookData(name=filename, sheets=sheets)


def _read_shared_strings(workbook: zipfile.ZipFile) -> List[str]:
    if "xl/sharedStrings.xml" not in workbook.namelist():
        return []

    root = ET.fromstring(workbook.read("xl/sharedStrings.xml"))
    strings: List[str] = []
    for node in root.findall("main:si", XLSX_NS):
        parts = [text.text or "" for text in node.iterfind(".//main:t", XLSX_NS)]
        strings.append("".join(parts))
    return strings


def _read_sheet_targets(workbook: zipfile.ZipFile) -> List[tuple]:
    workbook_root = ET.fromstring(workbook.read("xl/workbook.xml"))
    rels_root = ET.fromstring(workbook.read("xl/_rels/workbook.xml.rels"))
    rel_tag = f"{{{XML_REL_NS}}}Relationship"
    rel_map = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels_root.findall(rel_tag)}

    sheets = []
    sheets_parent = workbook_root.find("main:sheets", XLSX_NS)
    if sheets_parent is None:
        return sheets

    relation_key = f"{{{OFFICE_REL_NS}}}id"
    for sheet in sheets_parent:
        relation_id = sheet.attrib.get(relation_key)
        target = rel_map.get(relation_id)
        if not target:
            continue
        if not target.startswith("xl/"):
            target = "xl/" + target.lstrip("/").replace("../", "")
        sheets.append((sheet.attrib.get("name", "Sheet"), target))
    return sheets


def _read_sheet_rows(
    workbook: zipfile.ZipFile, sheet_target: str, shared_strings: Sequence[str]
) -> List[List[str]]:
    root = ET.fromstring(workbook.read(sheet_target))
    result: List[List[str]] = []
    for row in root.findall(".//main:sheetData/main:row", XLSX_NS):
        row_values = {}
        for cell in row.findall("main:c", XLSX_NS):
            cell_ref = cell.attrib.get("r", "")
            col_letters = "".join(ch for ch in cell_ref if ch.isalpha())
            if not col_letters:
                continue
            col_index = column_letter_to_index(col_letters)
            row_values[col_index] = _read_cell_value(cell, shared_strings)
        if not row_values:
            continue
        width = max(row_values) + 1
        normalized = [clean_cell(row_values.get(index, "")) for index in range(width)]
        result.append(normalized)
    return trim_empty_rows(result)


def _read_cell_value(cell, shared_strings: Sequence[str]) -> str:
    cell_type = cell.attrib.get("t")
    value_node = cell.find("main:v", XLSX_NS)

    if cell_type == "s":
        if value_node is None or value_node.text is None:
            return ""
        try:
            return str(shared_strings[int(value_node.text)])
        except (ValueError, IndexError):
            return value_node.text

    if cell_type == "inlineStr":
        inline_text = cell.find(".//main:t", XLSX_NS)
        return inline_text.text if inline_text is not None and inline_text.text else ""

    if cell_type == "b":
        return "TRUE" if value_node is not None and value_node.text == "1" else "FALSE"

    if value_node is None or value_node.text is None:
        return ""

    return value_node.text


def trim_empty_rows(rows: Iterable[Sequence[str]]) -> List[List[str]]:
    cleaned = []
    for row in rows:
        values = list(row)
        while values and values[-1] == "":
            values.pop()
        if values:
            cleaned.append(values)
    return cleaned


def column_letter_to_index(column_letters: str) -> int:
    total = 0
    for char in column_letters.upper():
        total = total * 26 + (ord(char) - 64)
    return total - 1


def column_index_to_letter(index: int) -> str:
    if index < 0:
        raise ValueError("Column index must be non-negative.")
    letters = []
    number = index + 1
    while number:
        number, remainder = divmod(number - 1, 26)
        letters.append(chr(65 + remainder))
    return "".join(reversed(letters))


def clean_cell(value) -> str:
    if value is None:
        return ""
    text = str(value).replace("\r\n", "\n").replace("\r", "\n")
    return text.strip()


def workbook_to_xlsx_bytes(sheets: Sequence[SheetData]) -> bytes:
    safe_sheets = _dedupe_sheet_names(sheets)
    memory = io.BytesIO()
    with zipfile.ZipFile(memory, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", _build_content_types_xml(len(safe_sheets)))
        archive.writestr("_rels/.rels", _build_root_rels_xml())
        archive.writestr("xl/workbook.xml", _build_workbook_xml(safe_sheets))
        archive.writestr("xl/_rels/workbook.xml.rels", _build_workbook_rels_xml(safe_sheets))
        for sheet_index, sheet in enumerate(safe_sheets, start=1):
            archive.writestr(
                f"xl/worksheets/sheet{sheet_index}.xml",
                _build_sheet_xml(sheet.rows),
            )
    memory.seek(0)
    return memory.read()


def _dedupe_sheet_names(sheets: Sequence[SheetData]) -> List[SheetData]:
    seen = {}
    output = []
    for sheet in sheets:
        raw_name = re.sub(r"[:\\/?*\[\]]", "", (sheet.name or "Sheet").strip())[:31] or "Sheet"
        count = seen.get(raw_name, 0)
        seen[raw_name] = count + 1
        name = raw_name if count == 0 else f"{raw_name[:28]}-{count}"
        output.append(SheetData(name=name[:31], rows=sheet.rows))
    return output


def _build_content_types_xml(sheet_count: int) -> str:
    overrides = "\n".join(
        [
            f'<Override PartName="/xl/worksheets/sheet{index}.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            for index in range(1, sheet_count + 1)
        ]
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        f"{overrides}"
        "</Types>"
    )


def _build_root_rels_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{XML_REL_NS}">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="xl/workbook.xml"/>'
        "</Relationships>"
    )


def _build_workbook_xml(sheets: Sequence[SheetData]) -> str:
    sheet_nodes = []
    for index, sheet in enumerate(sheets, start=1):
        sheet_nodes.append(
            f'<sheet name="{xml_escape(sheet.name)}" sheetId="{index}" '
            f'r:id="rId{index}"/>'
        )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f"<sheets>{''.join(sheet_nodes)}</sheets>"
        "</workbook>"
    )


def _build_workbook_rels_xml(sheets: Sequence[SheetData]) -> str:
    rel_nodes = []
    for index, _sheet in enumerate(sheets, start=1):
        rel_nodes.append(
            f'<Relationship Id="rId{index}" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            f'Target="worksheets/sheet{index}.xml"/>'
        )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{XML_REL_NS}">'
        f"{''.join(rel_nodes)}"
        "</Relationships>"
    )


def _build_sheet_xml(rows: Sequence[Sequence[str]]) -> str:
    xml_rows = []
    for row_index, row in enumerate(rows, start=1):
        cells = []
        for col_index, value in enumerate(row):
            ref = f"{column_index_to_letter(col_index)}{row_index}"
            if value is None or value == "":
                continue
            numeric = _maybe_number(value)
            if numeric is not None:
                cells.append(f'<c r="{ref}"><v>{numeric}</v></c>')
            else:
                text = xml_escape(str(value))
                cells.append(
                    f'<c r="{ref}" t="inlineStr"><is><t xml:space="preserve">{text}</t></is></c>'
                )
        xml_rows.append(f'<row r="{row_index}">{"".join(cells)}</row>')
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f"<sheetData>{''.join(xml_rows)}</sheetData>"
        "</worksheet>"
    )


def _maybe_number(value):
    if isinstance(value, (int, float)):
        return value
    return None


def xml_escape(text: str) -> str:
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&apos;")
    )
