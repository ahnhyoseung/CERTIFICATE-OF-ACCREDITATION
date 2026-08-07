# -*- coding: utf-8 -*-
"""
compare_hwp_excel.py
=====================
한글(HWP) 문서를 "기준(source of truth)"으로 삼아 엑셀 파일과 비교하고,
엑셀을 한글 내용에 맞게 동일화(sync)하는 스크립트.

비교 대상 4개 항목 (요청사항 그대로):
    1) 한글 "제품 및 물질"      <-> 엑셀 "제품 및 물질"
    2) 한글 "규격번호"          <-> 엑셀 "규격코드"
    3) 한글 "규격명"            <-> 엑셀 "규격명"
    4) 한글 "구성요소,특성(시험범위)" <-> 엑셀 "구성요소,특성(시험범위)(한글)"

매칭(짝짓기) 방법
------------------
행을 매칭할 고유 키가 마땅치 않은 경우가 많으므로, 기본적으로
"규격번호(=규격코드)"를 키로 사용합니다. 규격번호가 같은데 나머지
값이 다르면 "내용 불일치"로, 규격번호 자체가 한쪽에만 있으면
"누락/추가"로 분류합니다.

사용법
------
    python compare_hwp_excel.py \
        --hwp "2_최종_인정신청분야_및_범위-_주_엠아이티지.hwp" \
        --xls "_주_엠아이티지_인정분야.xls" \
        --out-dir ./output

결과물 (out-dir 안에 생성):
    - diff_report.md         : 사람이 읽기 좋은 비교 리포트 (Markdown)
    - diff_report.csv        : 같은 내용을 엑셀에서 열어볼 수 있는 CSV
    - synced.xlsx             : 한글 내용대로 값이 보정된 엑셀 (변경 셀 노란색 강조)

Claude API를 활용한 유사도 판정 (선택 사항)
--------------------------------------------
공백/기호 차이만 있는데 "다르다"고 잘못 판정되는 것을 줄이기 위해,
--use-claude 옵션을 주면 애매한 케이스(문자열이 다른데 얼마나 다른지
애매한 경우)에 대해 Anthropic Claude API를 호출해서
"의미상 동일한지" 여부를 한 번 더 확인합니다.
    - 환경변수 ANTHROPIC_API_KEY 필요
    - pip install anthropic
"""

from __future__ import annotations

import argparse
import csv
import os
import re
import sys
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill

from hwp_parser import parse_hwp_main_table


# --------------------------------------------------------------------------
# 1. 엑셀 컬럼 -> 표준 이름 매핑
#    (엑셀 헤더 문구가 프로젝트마다 조금씩 다를 수 있어 키워드 방식으로 탐색)
# --------------------------------------------------------------------------
EXCEL_HEADER_KEYWORDS = {
    "product": ["제품", "물질"],
    "spec_no": ["규격코드"],
    "spec_name": ["규격명"],
    "component": ["구성요소", "특성"],
}

# 비교 결과 각 필드에 대응하는, synced.xlsx 에 실제로 값을 써넣을 엑셀 컬럼 후보
# (엑셀 헤더 자동탐색 실패 시 이 이름으로 폴백)
FALLBACK_EXCEL_COLS = {
    "product": "제품 및 물질",
    "spec_no": "규격코드",
    "spec_name": "규격명",
    "component": "구성요소,특성(시험범위)(한글)",
}


def find_excel_columns(columns: List[str]) -> Dict[str, Optional[str]]:
    col_map: Dict[str, Optional[str]] = {k: None for k in EXCEL_HEADER_KEYWORDS}
    for col in columns:
        col_str = str(col)
        norm = col_str.replace(" ", "")
        for std_name, keywords in EXCEL_HEADER_KEYWORDS.items():
            if col_map[std_name] is not None:
                continue
            # "구성요소,특성(시험범위)(영문)" 같은 영문 컬럼은 제외
            if "영문" in norm:
                continue
            if all(kw in norm for kw in keywords):
                col_map[std_name] = col_str
    # 폴백 처리
    for k, v in col_map.items():
        if v is None and FALLBACK_EXCEL_COLS[k] in columns:
            col_map[k] = FALLBACK_EXCEL_COLS[k]
    return col_map


# --------------------------------------------------------------------------
# 2. 텍스트 정규화 (비교용) - 공백/괄호/기호 차이로 인한 오탐 줄이기
# --------------------------------------------------------------------------
def normalize(text: object) -> str:
    if text is None:
        return ""
    s = str(text)
    s = s.replace("\r\n", " ").replace("\n", " ").replace("\r", " ")
    s = re.sub(r"\s+", " ", s)
    # 전각/반각, 유사 기호 통일 (예: 물결표, 콜론 앞뒤 공백 등)
    s = s.replace("：", ":").replace("～", "~")
    s = s.strip()
    return s


def normalize_spec_no(text: object) -> str:
    """규격번호/규격코드는 공백을 아예 없애고 대문자로 통일해서 매칭 키로 사용."""
    s = normalize(text)
    s = s.upper().replace(" ", "")
    return s


# --------------------------------------------------------------------------
# 3. 비교 결과 자료구조
# --------------------------------------------------------------------------
@dataclass
class FieldDiff:
    field: str
    hwp_value: str
    excel_value: str
    same: bool


@dataclass
class RowResult:
    spec_no_key: str
    status: str  # "matched" | "hwp_only" | "excel_only"
    hwp_record: Optional[Dict[str, str]] = None
    excel_row_index: Optional[int] = None  # 엑셀 데이터프레임 상의 행 index
    field_diffs: List[FieldDiff] = field(default_factory=list)

    @property
    def has_mismatch(self) -> bool:
        return any(not d.same for d in self.field_diffs)


# --------------------------------------------------------------------------
# 4. 메인 비교 로직
# --------------------------------------------------------------------------
def load_excel(xls_path: str) -> Tuple[pd.DataFrame, Dict[str, Optional[str]]]:
    engine = "xlrd" if xls_path.lower().endswith(".xls") else None
    df = pd.read_excel(xls_path, engine=engine, dtype=str)
    df = df.fillna("")
    col_map = find_excel_columns(list(df.columns))
    missing = [k for k, v in col_map.items() if v is None]
    if missing:
        raise RuntimeError(
            f"엑셀에서 다음 컬럼을 찾지 못했습니다: {missing}\n"
            f"실제 컬럼 목록: {list(df.columns)}"
        )
    return df, col_map


def compare(hwp_records: List[Dict[str, str]], df: pd.DataFrame, col_map: Dict[str, str]) -> List[RowResult]:
    # 엑셀: 규격코드 정규화 키 -> 행 index 리스트 (중복 있을 수 있음)
    excel_key_to_indices: Dict[str, List[int]] = {}
    for idx, row in df.iterrows():
        key = normalize_spec_no(row[col_map["spec_no"]])
        excel_key_to_indices.setdefault(key, []).append(idx)

    results: List[RowResult] = []
    used_excel_indices = set()

    for rec in hwp_records:
        key = normalize_spec_no(rec.get("spec_no", ""))
        if not key:
            continue
        candidates = excel_key_to_indices.get(key, [])
        if not candidates:
            results.append(RowResult(spec_no_key=key, status="hwp_only", hwp_record=rec))
            continue

        excel_idx = candidates[0]  # 동일 규격번호가 여러 개면 첫 번째만 사용
        used_excel_indices.add(excel_idx)
        row = df.loc[excel_idx]

        diffs = [
            FieldDiff(
                field="제품 및 물질",
                hwp_value=rec.get("product", ""),
                excel_value=str(row[col_map["product"]]),
                same=normalize(rec.get("product", "")) == normalize(row[col_map["product"]]),
            ),
            FieldDiff(
                field="규격번호/규격코드",
                hwp_value=rec.get("spec_no", ""),
                excel_value=str(row[col_map["spec_no"]]),
                same=normalize_spec_no(rec.get("spec_no", "")) == normalize_spec_no(row[col_map["spec_no"]]),
            ),
            FieldDiff(
                field="규격명",
                hwp_value=rec.get("spec_name", ""),
                excel_value=str(row[col_map["spec_name"]]),
                same=normalize(rec.get("spec_name", "")) == normalize(row[col_map["spec_name"]]),
            ),
            FieldDiff(
                field="구성요소,특성(시험범위)",
                hwp_value=rec.get("component", ""),
                excel_value=str(row[col_map["component"]]),
                same=normalize(rec.get("component", "")) == normalize(row[col_map["component"]]),
            ),
        ]

        results.append(
            RowResult(
                spec_no_key=key,
                status="matched",
                hwp_record=rec,
                excel_row_index=excel_idx,
                field_diffs=diffs,
            )
        )

    # 엑셀에만 있고 한글엔 없는 행들
    for idx, row in df.iterrows():
        if idx in used_excel_indices:
            continue
        key = normalize_spec_no(row[col_map["spec_no"]])
        results.append(RowResult(spec_no_key=key, status="excel_only", excel_row_index=idx))

    return results


# --------------------------------------------------------------------------
# 5. (선택) Claude API로 애매한 텍스트 차이 재판정
# --------------------------------------------------------------------------
def refine_with_claude(results: List[RowResult]) -> None:
    """
    normalize() 만으로는 "다르다"고 나오지만, 실제로는 의미가 같은
    (띄어쓰기/기호 차이 정도) 애매한 케이스를 Claude에게 다시 물어봐서
    same 여부를 보정한다. field_diffs 를 in-place로 수정.
    """
    try:
        import anthropic
    except ImportError:
        print("[경고] anthropic 패키지가 없어 --use-claude 를 건너뜁니다. "
              "`pip install anthropic` 후 다시 시도하세요.", file=sys.stderr)
        return

    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        print("[경고] ANTHROPIC_API_KEY 환경변수가 없어 --use-claude 를 건너뜁니다.",
              file=sys.stderr)
        return

    client = anthropic.Anthropic(api_key=api_key)

    for r in results:
        if r.status != "matched":
            continue
        for d in r.field_diffs:
            if d.same:
                continue
            # 완전히 다른 문자열(길이 차이가 너무 큰 경우)은 굳이 물어보지 않음 -> 비용 절감
            if not d.hwp_value or not d.excel_value:
                continue
            len_ratio = len(d.hwp_value) / max(1, len(d.excel_value))
            if len_ratio < 0.5 or len_ratio > 2.0:
                continue

            prompt = (
                "다음 두 텍스트는 시험 규격/성분 관련 문서에서 나온 값입니다. "
                "표기나 단위, 띄어쓰기 차이가 아니라 '실질적인 값(숫자, 단위, 조성 등)'이 "
                "다른지 판단해주세요. 실질적으로 동일하면 SAME, 다르면 DIFFERENT 라고만 "
                "한 단어로 답하세요.\n\n"
                f"텍스트 A (한글 문서 기준):\n{d.hwp_value}\n\n"
                f"텍스트 B (엑셀):\n{d.excel_value}"
            )
            try:
                msg = client.messages.create(
                    model="claude-sonnet-4-6",
                    max_tokens=10,
                    messages=[{"role": "user", "content": prompt}],
                )
                answer = "".join(
                    block.text for block in msg.content if block.type == "text"
                ).strip().upper()
                if answer.startswith("SAME"):
                    d.same = True
            except Exception as e:  # 네트워크 오류 등은 조용히 무시하고 원래 판정 유지
                print(f"[경고] Claude 호출 실패, 원래 판정 유지: {e}", file=sys.stderr)


# --------------------------------------------------------------------------
# 6. 리포트 출력
# --------------------------------------------------------------------------
def write_reports(results: List[RowResult], out_dir: str) -> Tuple[str, str]:
    os.makedirs(out_dir, exist_ok=True)
    md_path = os.path.join(out_dir, "diff_report.md")
    csv_path = os.path.join(out_dir, "diff_report.csv")

    matched = [r for r in results if r.status == "matched"]
    mismatched = [r for r in matched if r.has_mismatch]
    hwp_only = [r for r in results if r.status == "hwp_only"]
    excel_only = [r for r in results if r.status == "excel_only"]

    with open(md_path, "w", encoding="utf-8") as f:
        f.write("# 한글(HWP) vs 엑셀 비교 리포트\n\n")
        f.write(f"- 매칭된 규격번호 수: {len(matched)}\n")
        f.write(f"- 내용이 다른 규격번호 수: {len(mismatched)}\n")
        f.write(f"- 한글에만 있는 규격번호 수(엑셀 누락): {len(hwp_only)}\n")
        f.write(f"- 엑셀에만 있는 규격번호 수(한글에 없음): {len(excel_only)}\n\n")

        if mismatched:
            f.write("## 내용 불일치 상세\n\n")
            for r in mismatched:
                f.write(f"### 규격번호: {r.hwp_record.get('spec_no', '')}\n\n")
                f.write("| 항목 | 한글(기준) | 엑셀 | 일치여부 |\n")
                f.write("|---|---|---|---|\n")
                for d in r.field_diffs:
                    mark = "✅" if d.same else "❌"
                    f.write(f"| {d.field} | {d.hwp_value} | {d.excel_value} | {mark} |\n")
                f.write("\n")

        if hwp_only:
            f.write("## 한글에는 있는데 엑셀에 없는 규격번호\n\n")
            for r in hwp_only:
                f.write(f"- {r.hwp_record.get('spec_no', '')} : {r.hwp_record.get('spec_name', '')}\n")
            f.write("\n")

        if excel_only:
            f.write("## 엑셀에는 있는데 한글에 없는 규격번호 (한글 문서에 해당 대/중분류가 없을 수도 있음)\n\n")
            for r in excel_only:
                f.write(f"- {r.spec_no_key}\n")
            f.write("\n")

    with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(["규격번호", "상태", "항목", "한글(기준)", "엑셀", "일치여부"])
        for r in results:
            if r.status == "matched":
                for d in r.field_diffs:
                    writer.writerow([
                        r.hwp_record.get("spec_no", ""),
                        "일치" if not r.has_mismatch else "불일치",
                        d.field, d.hwp_value, d.excel_value,
                        "일치" if d.same else "불일치",
                    ])
            elif r.status == "hwp_only":
                writer.writerow([r.hwp_record.get("spec_no", ""), "엑셀누락", "", "", "", ""])
            else:
                writer.writerow([r.spec_no_key, "한글없음", "", "", "", ""])

    return md_path, csv_path


# --------------------------------------------------------------------------
# 7. 엑셀 동일화(sync) - 한글 값으로 덮어쓰기 + 변경 셀 노란색 강조
# --------------------------------------------------------------------------
def write_synced_excel(xls_path: str, results: List[RowResult], col_map: Dict[str, str], out_dir: str) -> str:
    # openpyxl은 legacy .xls를 못 읽으므로, pandas로 읽은 뒤 xlsx로 새로 저장
    df = pd.read_excel(xls_path, engine=("xlrd" if xls_path.lower().endswith(".xls") else None), dtype=str)
    df = df.fillna("")

    changed_cells = []  # (row_idx(0-base, 데이터 시작 기준), col_name)

    for r in results:
        if r.status != "matched" or not r.has_mismatch:
            continue
        for d in r.field_diffs:
            std_name = {
                "제품 및 물질": "product",
                "규격번호/규격코드": "spec_no",
                "규격명": "spec_name",
                "구성요소,특성(시험범위)": "component",
            }[d.field]
            if d.same:
                continue
            excel_col = col_map[std_name]
            df.at[r.excel_row_index, excel_col] = d.hwp_value
            changed_cells.append((r.excel_row_index, excel_col))

    out_path = os.path.join(out_dir, "synced.xlsx")
    df.to_excel(out_path, index=False)

    # 변경된 셀 노란색 강조 + 폰트 통일
    wb = load_workbook(out_path)
    ws = wb.active
    header = [cell.value for cell in ws[1]]
    fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    for row_idx, col_name in changed_cells:
        col_idx = header.index(col_name) + 1  # 1-base
        excel_row = row_idx + 2  # 1행은 헤더, pandas index는 0부터, 데이터는 2행부터
        cell = ws.cell(row=excel_row, column=col_idx)
        cell.fill = fill

    for row in ws.iter_rows():
        for cell in row:
            cell.font = Font(name="Arial", size=cell.font.size or 10)

    wb.save(out_path)
    return out_path


# --------------------------------------------------------------------------
# 8. 파일 선택 (경로를 안 주면 탐색기 창을 띄워서 고르게 함)
# --------------------------------------------------------------------------
def pick_file_with_dialog(title: str, filetypes: List[Tuple[str, str]]) -> str:
    try:
        import tkinter as tk
        from tkinter import filedialog
    except ImportError:
        raise RuntimeError(
            "tkinter를 사용할 수 없습니다. --hwp/--xls 옵션으로 경로를 직접 지정해주세요."
        )
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    path = filedialog.askopenfilename(title=title, filetypes=filetypes)
    root.destroy()
    if not path:
        raise SystemExit("파일을 선택하지 않아 종료합니다.")
    return path


def resolve_hwp_xls_from_dropped(paths: List[str]) -> Tuple[str, str]:
    """드래그&드롭으로 받은 파일 경로 목록에서 hwp/xls를 자동으로 구분."""
    hwp_path = next((p for p in paths if p.lower().endswith(".hwp")), None)
    xls_path = next((p for p in paths if p.lower().endswith((".xls", ".xlsx"))), None)
    if not hwp_path or not xls_path:
        raise SystemExit(
            "드래그한 파일 중 .hwp 하나와 .xls/.xlsx 하나가 반드시 있어야 합니다.\n"
            f"받은 파일: {paths}"
        )
    return hwp_path, xls_path


# --------------------------------------------------------------------------
# 9. CLI
# --------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(description="HWP(기준) vs Excel 비교/동일화 도구")
    parser.add_argument("--hwp", help="기준이 되는 .hwp 파일 경로")
    parser.add_argument("--xls", help="비교할 .xls/.xlsx 파일 경로")
    parser.add_argument("--out-dir", default="./output", help="결과물 저장 폴더")
    parser.add_argument("--use-claude", action="store_true",
                         help="애매한 불일치를 Claude API로 재판정 (ANTHROPIC_API_KEY 필요)")
    # 드래그&드롭용: hwp/xls 파일을 스크립트(또는 .bat) 위에 끌어다 놓으면
    # 여기로 파일 경로가 그대로 인자로 들어온다.
    parser.add_argument("dropped", nargs="*", help=argparse.SUPPRESS)
    args = parser.parse_args()

    if args.hwp and args.xls:
        hwp_path, xls_path = args.hwp, args.xls
    elif args.dropped:
        hwp_path, xls_path = resolve_hwp_xls_from_dropped(args.dropped)
    else:
        # 아무 경로도 안 줬으면 탐색기 창을 띄워서 고르게 함
        print("경로가 지정되지 않아 파일 선택 창을 엽니다...")
        hwp_path = pick_file_with_dialog(
            "기준이 되는 .hwp 파일 선택", [("HWP 파일", "*.hwp"), ("모든 파일", "*.*")]
        )
        xls_path = pick_file_with_dialog(
            "비교할 .xls/.xlsx 파일 선택", [("Excel 파일", "*.xls;*.xlsx"), ("모든 파일", "*.*")]
        )

    print(f"[1/5] HWP 파싱 중: {hwp_path}")
    hwp_records = parse_hwp_main_table(hwp_path)
    print(f"      -> {len(hwp_records)}개 행 추출")

    print(f"[2/5] 엑셀 로딩 중: {xls_path}")
    df, col_map = load_excel(xls_path)
    print(f"      -> 컬럼 매핑: {col_map}")

    print("[3/5] 비교 중...")
    results = compare(hwp_records, df, col_map)

    if args.use_claude:
        print("[3.5/5] Claude API로 애매한 케이스 재판정 중...")
        refine_with_claude(results)

    print("[4/5] 리포트 작성 중...")
    md_path, csv_path = write_reports(results, args.out_dir)
    print(f"      -> {md_path}")
    print(f"      -> {csv_path}")

    print("[5/5] 엑셀 동일화(synced.xlsx) 작성 중...")
    synced_path = write_synced_excel(xls_path, results, col_map, args.out_dir)
    print(f"      -> {synced_path}")

    print("\n완료.")
    # 드래그&드롭(.bat)으로 실행했을 때 창이 바로 닫히는 것을 방지
    try:
        input("\n결과 확인 후 아무 키나 눌러 창을 닫으세요...")
    except EOFError:
        pass


if __name__ == "__main__":
    main()
