# -*- coding: utf-8 -*-
"""
hwp_parser.py
=============
.hwp(한글) 파일 안에 들어있는 "표(Table)"를 파싱해서
[{컬럼명: 값, ...}, ...] 형태의 리스트로 돌려주는 모듈.

동작 원리
---------
1. `hwp5proc xml <파일>` 명령을 실행해서 hwp 내부 구조를 XML로 뽑아낸다.
   (pyhwp 패키지에 포함된 CLI 도구. `pip install pyhwp olefile` 로 설치)
2. XML에서 <TableControl> 안의 <TableRow>/<TableCell> 을 순서대로 읽으며
   각 셀의 텍스트를 복원한다. (셀 하나 = <Paragraph>/<LineSeg>/<Text> 로
   줄바꿈 되어 쪼개져 있으므로 전부 이어붙여서 원문을 복원)
3. 첫 번째 행(row 0)을 헤더로 보고, 그 헤더 텍스트를 기준으로
   "제품 및 물질", "규격번호"/"규격코드", "규격명",
   "구성요소" 등 우리가 필요로 하는 컬럼을 자동으로 찾는다.

주의
----
- hwp5proc는 pyhwp 패키지가 설치되어 있어야 사용 가능합니다.
    pip install olefile pyhwp
- 배포용(distribution) 암호화된 hwp 파일은 지원하지 않습니다.
- 표가 여러 개면 가장 큰(행 수가 많은) 표를 기본으로 사용합니다.
  (필요하면 table_index 로 원하는 표를 지정할 수 있음)
"""

from __future__ import annotations

import re
import subprocess
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from typing import List, Dict, Optional


# hwp 헤더 텍스트 -> 우리가 쓸 표준 컬럼명 매핑용 키워드
# (hwp 파일마다 줄바꿈/문구가 조금씩 다를 수 있어 "포함 여부"로 판단)
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
    headers: List[str]           # 원본 헤더 텍스트 (열 순서대로)
    rows: List[List[str]]        # 데이터 행들 (열 순서대로)

    def to_records(self) -> List[Dict[str, str]]:
        """헤더 키워드를 기준으로 표준 컬럼명이 붙은 dict 리스트로 변환."""
        col_map = self._map_columns()
        records = []
        for row in self.rows:
            rec = {}
            for std_name, idx in col_map.items():
                rec[std_name] = row[idx].strip() if idx is not None and idx < len(row) else ""
            rec["_raw_row"] = row  # 원본 행도 같이 보관 (디버깅용)
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
    """hwp5proc xml 명령을 실행해서 XML 문자열을 반환."""
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


def _cell_text(cell_elem: ET.Element) -> str:
    """<TableCell> 하위의 모든 <Text> 노드를 문서 순서대로 이어붙여 원문 복원."""
    parts = [t.text or "" for t in cell_elem.iter("Text")]
    text = "".join(parts)
    # hwp 원문에 있는 불필요한 공백/개행 정리 (내용 자체는 보존)
    text = re.sub(r"[ \t]+", " ", text)
    text = text.strip()
    return text


def parse_hwp_tables(hwp_path: str) -> List[HwpTable]:
    """hwp 파일 안의 모든 표를 [HwpTable, ...] 로 반환."""
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


def _looks_like_data_table(t: "HwpTable") -> bool:
    """헤더에 우리가 찾는 컬럼 키워드가 하나라도 있으면 '데이터 표'로 간주.

    KOLAS 인정신청분야 문서는 대분류/중분류별로 페이지가 나뉘어 있어,
    같은 헤더(순번/제품 및 물질/규격번호/규격명/...)를 가진 표가
    여러 개(문서마다 다르지만 보통 5~15개) 반복해서 나온다. 표 하나만
    고르면 나머지 표에 있는 규격번호가 전부 "한글에 없음"으로 잘못
    잡히므로, 헤더가 맞는 표는 전부 모아서 합쳐야 한다.
    """
    header_norm = "".join(t.headers).replace(" ", "")
    return any(
        any(kw.replace(" ", "") in header_norm for kw in keywords)
        for keywords in HEADER_KEYWORDS.values()
    )


def parse_hwp_main_table(hwp_path: str, table_index: Optional[int] = None) -> List[Dict[str, str]]:
    """
    가장 일반적인 사용법: hwp 파일에서 표를 찾아 표준화된 레코드로 반환.

    table_index 를 지정하면 그 표 하나만 사용한다.
    지정하지 않으면, 우리가 찾는 컬럼(규격번호/규격명 등) 헤더를 가진
    표를 "전부" 찾아서 하나로 합친다. (KOLAS 인정신청분야 문서처럼
    대분류별로 표가 여러 개로 나뉘어 있는 경우, 표 하나만 골랐을 때
    나머지 표의 항목이 통째로 누락되는 문제를 방지하기 위함)
    """
    tables = parse_hwp_tables(hwp_path)
    if not tables:
        raise RuntimeError(f"{hwp_path} 안에서 표를 찾지 못했습니다.")

    if table_index is not None:
        chosen_tables = [tables[table_index]]
    else:
        chosen_tables = [t for t in tables if _looks_like_data_table(t)]
        if not chosen_tables:
            # 헤더 키워드로 하나도 못 찾으면 기존 방식(가장 큰 표)으로 폴백
            chosen_tables = [max(tables, key=lambda t: len(t.rows))]

    records: List[Dict[str, str]] = []
    for t in chosen_tables:
        records.extend(t.to_records())
    return records


if __name__ == "__main__":
    import sys
    import json

    if len(sys.argv) < 2:
        print("사용법: python hwp_parser.py <파일.hwp>")
        sys.exit(1)

    recs = parse_hwp_main_table(sys.argv[1])
    print(json.dumps(recs, ensure_ascii=False, indent=2))