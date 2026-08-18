# -*- coding: utf-8 -*-
from __future__ import annotations

import re
import subprocess
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from typing import List, Dict, Optional, Tuple


HEADER_KEYWORDS = {
    "product": ["제품", "물질"],
    "spec_no": ["규격번호", "규격코드"],
    "spec_name": ["규격명"],
    "component": ["구성요소", "특성", "시험범위"],
    "field_test": ["현장", "시험"],
    "remark": ["비고"],
}


@dataclass
class HwpTable:
    headers: List[str]
    rows: List[List[str]]

    def to_records(self) -> List[Dict[str, str]]:
        col_map = self._map_columns()
        records = []
        for row in self.rows:
            rec = {}
            for std_name, idx in col_map.items():
                rec[std_name] = row[idx].strip() if idx is not None and idx < len(row) else ""
            rec["_raw_row"] = row
            records.append(rec)
        return records

    def _map_columns(self) -> Dict[str, Optional[int]]:
        col_map: Dict[str, Optional[int]] = {k: None for k in HEADER_KEYWORDS}
        for idx, h in enumerate(self.headers):
            h_norm = h.replace(" ", "")
            for std_name, keywords in HEADER_KEYWORDS.items():
                if col_map[std_name] is not None:
                    continue
                if all(kw.replace(" ", "") in h_norm for kw in [keywords[0]]) and any(
                    kw.replace(" ", "") in h_norm for kw in keywords
                ):
                    col_map[std_name] = idx
        return col_map


def _run_hwp5proc_xml(hwp_path: str) -> str:
    try:
        result = subprocess.run(
            ["hwp5proc", "xml", hwp_path],
            capture_output=True,
            check=True,
        )
    except FileNotFoundError as e:
        raise RuntimeError(
            "hwp5proc 명령을 찾을 수 없습니다. `pip install olefile pyhwp` 로 설치하세요."
        ) from e
    except subprocess.CalledProcessError as e:
        raise RuntimeError(f"hwp5proc 실행 실패: {e.stderr.decode('utf-8', 'ignore')}") from e
    return result.stdout.decode("utf-8", errors="ignore")


def _paragraph_lines(para_elem: ET.Element) -> List[str]:
    lines: List[str] = []
    current: List[str] = []

    def rec(e: ET.Element) -> None:
        if e.tag == "TableControl":
            return
        if e.tag == "LineSeg":
            for c in e:
                if c.tag == "Text":
                    current.append(c.text or "")
                elif c.tag == "ControlChar" and c.get("code") == "10":
                    lines.append("".join(current))
                    current.clear()
            return
        for c in e:
            rec(c)

    rec(para_elem)
    lines.append("".join(current))
    return lines


def _cell_text(cell_elem: ET.Element) -> str:
    paragraphs = cell_elem.findall("Paragraph")
    if not paragraphs:
        parts = [t.text or "" for t in cell_elem.iter("Text")]
        text = "".join(parts)
        text = re.sub(r"[ \t]+", " ", text)
        return text.strip()

    lines: List[str] = []
    for p in paragraphs:
        lines.extend(_paragraph_lines(p))

    text = "\n".join(lines)
    text = "\n".join(re.sub(r"[ \t]+", " ", line).strip() for line in text.split("\n"))
    text = text.strip()
    return text


def parse_hwp_tables(hwp_path: str) -> List[HwpTable]:
    xml_str = _run_hwp5proc_xml(hwp_path)
    root = ET.fromstring(xml_str)

    tables: List[HwpTable] = []
    for table_ctrl in root.iter("TableControl"):
        body = table_ctrl.find("TableBody")
        if body is None:
            continue
        rows_elems = body.findall("TableRow")
        if not rows_elems:
            continue

        grid: List[List[str]] = []
        for row_elem in rows_elems:
            cells = row_elem.findall("TableCell")
            row_vals = [_cell_text(c) for c in cells]
            grid.append(row_vals)

        if not grid:
            continue

        headers = grid[0]
        data_rows = grid[1:]
        tables.append(HwpTable(headers=headers, rows=data_rows))

    return tables


def parse_hwp_main_table(hwp_path: str, table_index: Optional[int] = None) -> List[Dict[str, str]]:
    tables = parse_hwp_tables(hwp_path)
    if not tables:
        raise RuntimeError(f"{hwp_path} 안에서 표를 찾지 못했습니다.")

    if table_index is not None:
        chosen = tables[table_index]
    else:
        chosen = max(tables, key=lambda t: len(t.rows))

    return chosen.to_records()


_LOCATION_PATTERN = re.compile(r"^\((.+)\)$")


def _paragraph_own_text(elem: ET.Element) -> str:
    parts: List[str] = []

    def rec(e: ET.Element) -> None:
        if e.tag == "TableControl":
            return
        if e.tag == "Text":
            parts.append(e.text or "")
        for c in e:
            rec(c)

    rec(elem)
    text = "".join(parts)
    text = re.sub(r"[ \t]+", " ", text)
    return text.strip()


def parse_hwp_tables_with_location(hwp_path: str) -> List[Tuple[str, "HwpTable"]]:
    xml_str = _run_hwp5proc_xml(hwp_path)
    root = ET.fromstring(xml_str)

    results: List[Tuple[str, HwpTable]] = []
    context_paragraphs: List[str] = []

    def flush_location() -> str:
        location = ""
        for text in context_paragraphs:
            m = _LOCATION_PATTERN.match(text.strip())
            if m:
                location = m.group(1).strip()
        return location

    def extract_table(table_ctrl: ET.Element) -> Optional[HwpTable]:
        body = table_ctrl.find("TableBody")
        if body is None:
            return None
        rows_elems = body.findall("TableRow")
        if not rows_elems:
            return None
        grid: List[List[str]] = []
        for row_elem in rows_elems:
            cells = row_elem.findall("TableCell")
            grid.append([_cell_text(c) for c in cells])
        if not grid:
            return None
        return HwpTable(headers=grid[0], rows=grid[1:])

    def walk(elem: ET.Element, in_table: bool) -> None:
        tag = elem.tag
        if tag == "TableControl" and not in_table:
            location = flush_location()
            context_paragraphs.clear()
            table = extract_table(elem)
            if table is not None:
                results.append((location, table))
            for child in elem:
                walk(child, True)
            return
        if tag == "Paragraph" and not in_table:
            text = _paragraph_own_text(elem)
            if text:
                context_paragraphs.append(text)
            for child in elem:
                walk(child, in_table)
            return
        for child in elem:
            walk(child, in_table)

    walk(root, False)
    return results


def parse_hwp_all_tables(hwp_path: str, require_spec_no: bool = True) -> List[Dict[str, str]]:
    tables_with_loc = parse_hwp_tables_with_location(hwp_path)
    if not tables_with_loc:
        raise RuntimeError(f"{hwp_path} 안에서 표를 찾지 못했습니다.")

    records: List[Dict[str, str]] = []
    for location, table in tables_with_loc:
        col_map = table._map_columns()
        if require_spec_no and col_map.get("spec_no") is None:
            continue
        for rec in table.to_records():
            rec["location"] = location
            records.append(rec)
    return records