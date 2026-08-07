# 한글(HWP) ↔ 엑셀 비교/동일화 도구

한글 문서(`최종 인정신청분야 및 범위`)를 **기준**으로 삼아, 엑셀 문서
(`인정분야목록`)와 아래 4개 항목을 비교하고, 다르면 한글 값으로
엑셀을 맞춰주는 파이썬 도구입니다.

비교 항목:
1. 제품 및 물질
2. 규격번호(한글) ↔ 규격코드(엑셀)
3. 규격명
4. 구성요소,특성(시험범위)

## 폴더 구성

```
kolas_compare/
├── hwp_parser.py           # .hwp 표를 파싱하는 모듈
├── compare_hwp_excel.py    # 비교 + 리포트 + 엑셀 동일화 메인 스크립트
└── README.md
```

## 설치 (VS Code 터미널에서)

```bash
pip install olefile pyhwp pandas openpyxl xlrd anthropic
```

- `pyhwp`  : .hwp 파일을 읽기 위한 필수 패키지 (hwp5proc 명령 제공)
- `xlrd`   : 구버전 .xls 파일을 읽기 위해 필요 (.xlsx만 쓴다면 생략 가능)
- `anthropic` : `--use-claude` 옵션(선택)을 쓸 때만 필요

## 실행 방법 (3가지 중 편한 것으로)

### 방법 A. 경로를 직접 지정 (VS Code 터미널)

```bash
python compare_hwp_excel.py \
  --hwp "2_최종_인정신청분야_및_범위-_주_엠아이티지.hwp" \
  --xls "_주_엠아이티지_인정분야.xls" \
  --out-dir ./output
```

### 방법 B. 파일 선택 창(탐색기)으로 고르기

경로 없이 그냥 실행하면 탐색기 창이 2번(hwp 선택 → xls 선택) 뜹니다.

```bash
python compare_hwp_excel.py
```

### 방법 C. 드래그 앤 드롭 (Windows)

`run_compare.bat` 파일 위에 **.hwp 파일 1개 + .xls(또는 .xlsx) 파일 1개, 총 2개를
동시에 선택해서** 마우스로 끌어다 놓으면 자동으로 실행됩니다. (`.bat`을 그냥
더블클릭만 하면 방법 B처럼 파일 선택 창이 뜹니다.)

Claude API로 애매한 불일치(단위/띄어쓰기 차이 등)를 한 번 더 검증하고 싶다면:

```bash
export ANTHROPIC_API_KEY=sk-ant-...
python compare_hwp_excel.py --hwp ... --xls ... --use-claude
```

## 결과물 (`./output` 폴더)

| 파일 | 내용 |
|---|---|
| `diff_report.md` | 규격번호별로 4개 항목이 일치하는지 표로 정리한 리포트 |
| `diff_report.csv` | 같은 내용을 엑셀에서 열어볼 수 있는 CSV |
| `synced.xlsx` | 한글 문서 값으로 보정된 엑셀. **변경된 셀은 노란색**으로 표시됨 |

## 매칭(짝짓기) 방식

- 기본적으로 **규격번호(한글) = 규격코드(엑셀)** 를 키로 사용해서 행을 짝짓습니다.
  (공백 제거, 대문자 통일 후 비교)
- 규격번호가 한쪽에만 있으면 "한글에만 있음(엑셀 누락)" 또는
  "엑셀에만 있음(한글에 없음)"으로 리포트에 별도 표시됩니다.
- 같은 규격번호인데 나머지 값이 다르면 "내용 불일치"로 표시되고,
  `synced.xlsx` 에서는 한글 값으로 덮어써집니다.

## 표(테이블)가 여러 개인 hwp 파일

`hwp_parser.py` 는 기본적으로 "행 수가 가장 많은 표"를 자동으로 고릅니다.
파일 안에 표가 여러 개이고 원하는 표가 다르면:

```python
from hwp_parser import parse_hwp_main_table
records = parse_hwp_main_table("파일.hwp", table_index=1)  # 0부터 시작
```

## 엑셀 헤더 문구가 달라도 동작하도록

`compare_hwp_excel.py` 의 `EXCEL_HEADER_KEYWORDS` 는 헤더에 특정
키워드가 포함되어 있는지로 컬럼을 찾습니다. 예를 들어 "규격코드"라는
글자가 포함된 컬럼을 자동으로 찾으므로, 컬럼 순서가 바뀌어도 대체로
잘 동작합니다. 자동 인식이 실패하면 에러 메시지에 실제 컬럼 목록이
출력되니, `FALLBACK_EXCEL_COLS` 값을 실제 헤더명에 맞게 고쳐주세요.

## 주의사항

- 배포용(권한 제한/암호화)으로 저장된 .hwp 파일은 pyhwp가 열지 못할 수 있습니다.
- 규격번호가 완전히 같아야 매칭됩니다. 규격번호 표기 자체가 다르면
  (예: 오탈자) 자동으로 짝지어지지 않으니, `diff_report.md` 의
  "엑셀에만 있음"/"한글에만 있음" 목록을 사람이 한 번 확인하는 것을
  권장합니다.
