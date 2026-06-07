import re
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Sequence

from xlsx_toolkit import SheetData, WorkbookData


PARENT_SKU_HEADERS = ("*父SKU", "父SKU")
PROCESSED_SKU_HEADER = "处理后SKU"
REMOVABLE_PREFIX_PATTERN = re.compile(r"^(?:[A-Za-z]\d|[A-Za-z]{2,}\d*)-(.+)$")
CORE_SKU_PATTERN = re.compile(r"([A-Za-z]\d{6,}(?:-[A-Za-z0-9]+)*)$")


@dataclass
class ParentSkuPreviewRow:
    sheet_name: str
    row_number: int
    original_value: str
    processed_value: str
    changed: bool


def extract_parent_sku_core(raw_value: str) -> str:
    value = (raw_value or "").strip()
    if not value:
        return ""

    prefix_match = REMOVABLE_PREFIX_PATTERN.match(value)
    if prefix_match:
        return prefix_match.group(1).strip()

    match = CORE_SKU_PATTERN.search(value)
    if not match:
        return value
    return match.group(1)


def transform_parent_sku_workbook(
    workbook: WorkbookData, preview_limit: int = 120
) -> Dict[str, object]:
    output_sheets: List[SheetData] = []
    preview_rows: List[ParentSkuPreviewRow] = []
    matched_sheets = 0
    total_rows = 0
    changed_rows = 0

    for sheet in workbook.sheets:
        transformed_rows = [list(row) for row in sheet.rows]
        parent_column_index = find_parent_sku_column_index(transformed_rows)

        if parent_column_index is not None:
            matched_sheets += 1
            transformed_rows[0].insert(parent_column_index + 1, PROCESSED_SKU_HEADER)
            for row_offset, row in enumerate(transformed_rows[1:], start=2):
                original_value = row[parent_column_index] if len(row) > parent_column_index else ""
                processed_value = extract_parent_sku_core(original_value) if original_value else ""
                changed = bool(original_value) and processed_value != original_value

                if len(row) <= parent_column_index:
                    row.extend([""] * (parent_column_index + 1 - len(row)))
                row.insert(parent_column_index + 1, processed_value)

                if not original_value:
                    continue

                total_rows += 1
                if changed:
                    changed_rows += 1

                if len(preview_rows) < preview_limit:
                    preview_rows.append(
                        ParentSkuPreviewRow(
                            sheet_name=sheet.name,
                            row_number=row_offset,
                            original_value=original_value,
                            processed_value=processed_value,
                            changed=changed,
                        )
                    )

        output_sheets.append(SheetData(name=sheet.name, rows=transformed_rows))

    if matched_sheets == 0:
        raise ValueError("没有找到“*父SKU”或“父SKU”列，请确认上传的是正确的利润表。")

    original_name = Path(workbook.name).stem or "processed"
    output_filename = f"{original_name}-父SKU处理.xlsx"

    return {
        "sheets": output_sheets,
        "preview_rows": preview_rows,
        "summary": {
            "matched_sheets": matched_sheets,
            "total_rows": total_rows,
            "changed_rows": changed_rows,
            "unchanged_rows": total_rows - changed_rows,
            "preview_limit": preview_limit,
        },
        "output_filename": output_filename,
    }


def find_parent_sku_column_index(rows: Sequence[Sequence[str]]) -> Optional[int]:
    if not rows:
        return None

    header = rows[0]
    for index, value in enumerate(header):
        normalized = (value or "").strip()
        if normalized in PARENT_SKU_HEADERS:
            return index
    return None
