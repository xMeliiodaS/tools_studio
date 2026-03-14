from __future__ import annotations

from pathlib import Path
from typing import List, Optional
from xml.etree import ElementTree as ET
from zipfile import ZipFile

_NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pkg_rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}


def _column_index(cell_ref: str) -> int:
    col = "".join(ch for ch in cell_ref if ch.isalpha())
    idx = 0
    for ch in col:
        idx = idx * 26 + (ord(ch.upper()) - ord("A") + 1)
    return max(idx - 1, 0)


def _shared_strings(zf: ZipFile) -> List[str]:
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    values: List[str] = []
    for si in root.findall("main:si", _NS):
        parts = [t.text or "" for t in si.findall(".//main:t", _NS)]
        values.append("".join(parts))
    return values


def _resolve_sheet_path(zf: ZipFile, sheet_name: Optional[str]) -> str:
    workbook = ET.fromstring(zf.read("xl/workbook.xml"))
    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))

    sheets = workbook.findall("main:sheets/main:sheet", _NS)
    if not sheets:
        raise ValueError("Workbook has no sheets.")

    selected = None
    if sheet_name:
        for sheet in sheets:
            if sheet.attrib.get("name") == sheet_name:
                selected = sheet
                break
        if selected is None:
            raise ValueError(f"Sheet '{sheet_name}' not found in workbook.")
    else:
        selected = sheets[0]

    rel_id = selected.attrib.get(f"{{{_NS['rel']}}}id")
    for rel in rels.findall("pkg_rel:Relationship", _NS):
        if rel.attrib.get("Id") == rel_id:
            target = rel.attrib["Target"].lstrip("/")
            if not target.startswith("xl/"):
                target = f"xl/{target}"
            return target
    raise ValueError("Could not resolve sheet relationship.")


def read_xlsx_rows(path: str, sheet_name: Optional[str] = None) -> List[List[str]]:
    """Read worksheet rows from an .xlsx file into a matrix of strings."""
    xlsx_path = Path(path)
    with ZipFile(xlsx_path) as zf:
        shared = _shared_strings(zf)
        sheet_path = _resolve_sheet_path(zf, sheet_name)
        sheet = ET.fromstring(zf.read(sheet_path))

    matrix: List[List[str]] = []
    for row in sheet.findall("main:sheetData/main:row", _NS):
        row_values: List[str] = []
        for cell in row.findall("main:c", _NS):
            ref = cell.attrib.get("r", "A1")
            idx = _column_index(ref)
            while len(row_values) <= idx:
                row_values.append("")

            ctype = cell.attrib.get("t")
            value_node = cell.find("main:v", _NS)
            inline_node = cell.find("main:is/main:t", _NS)

            if inline_node is not None:
                value = inline_node.text or ""
            elif value_node is None:
                value = ""
            elif ctype == "s":
                shared_idx = int(value_node.text or 0)
                value = shared[shared_idx] if shared_idx < len(shared) else ""
            else:
                value = value_node.text or ""
            row_values[idx] = value
        matrix.append(row_values)
    return matrix
