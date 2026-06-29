# -*- coding: utf-8 -*-
"""
schema.json 으로 입력용 엑셀 양식을 생성한다.

생성되는 워크북: KOCARC_입력양식.xlsx
  - [사용법]   : 간단 안내
  - [환자목록] : 환자 1명 = 1행. 환자키 + 환자등록(시작) 입력 칸
  - [영역별 시트] : 공통영역, 예방역학 ... 각 시트는 환자키로 연결
  - [코드북]   : 코드값을 갖는 모든 항목의 (코드 = 의미) 사전

각 영역 시트 구조:
  1행 = 한글 라벨(질문)
  2행 = 필드명(서버 키)  ← 봇이 읽는 줄. 수정 금지.
  3행~ = 환자별 데이터 (A열 = 환자키)
"""
import os
import re
import sys
import json
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

BASE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(BASE, "KOCARC_입력양식.xlsx")

# 같은 질문의 단일-선택 체크박스(부/모/형제 ...)들을 드롭다운 1칸으로 묶을 때,
# 2행(필드명 줄)에 적는 마커. 봇이 이 마커를 보고 라벨->필드명 매칭을 처리한다.
# kocarc_bot.GROUP_PREFIX 와 반드시 동일해야 함.
GROUP_PREFIX = "@GROUP:"


def resource_path(name):
    """읽기전용 동봉 파일(schema.json). PyInstaller 번들 대응."""
    base = getattr(sys, "_MEIPASS", BASE)
    return os.path.join(base, name)

# 엑셀 시트 이름은 31자 제한 + 일부 문자 불가
SHEET_NAMES = {
    "common": "공통영역",
    "prevent": "예방역학",
    "community": "지역사회",
    "relief": "구급단계",
    "in_hosp": "병원단계",
    "alive_after": "소생후단계",
    "heart": "심장검사",
    "y_child": "소아소생술",
    "comment": "CommentLog",
}

HEADER_FILL = PatternFill("solid", fgColor="6BA68C")
NAME_FILL = PatternFill("solid", fgColor="EAF1ED")
KEY_FILL = PatternFill("solid", fgColor="FFF2CC")
SECTION_FILL = PatternFill("solid", fgColor="D9E8DF")
WHITE_BOLD = Font(bold=True, color="FFFFFF", size=10)
NAME_FONT = Font(color="888888", size=8)
THIN = Side(style="thin", color="BBBBBB")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)


def load_schema(schema_path=None):
    path = schema_path or resource_path("schema.json")
    with open(path, encoding="utf-8") as f:
        return json.load(f)


# 필드명 꼬리 -> 단위 (쪼개진 칸 구분용)
NAME_UNIT = [("_YY", "년"), ("_MM", "월"), ("_DD", "일"),
             ("_HOUR", "시"), ("_MIN", "분"), ("_DATE", "날짜")]
SHORT_UNITS = ("년", "월", "일", "시", "분", "세", "초")


def with_context(label, ctx):
    """질병명/구획 같은 상위 컨텍스트를 라벨 앞에 붙인다. (예: '고혈압 - 치료')"""
    ctx = (ctx or "").strip()
    if ctx and ctx not in label:
        return f"{ctx} - {label}"
    return label


def display_label(f):
    """사람이 읽기 쉬운 컬럼 라벨(상위 컨텍스트 포함)."""
    return with_context(_display_core(f), f.get("context"))


def _display_core(f):
    """컬럼 라벨 본문(컨텍스트 제외).
    - 쪼개진 일시/생년월일: (년)(월)(일)(날짜)(시)(분)
    - 단일 체크박스/라디오: 그 칸의 의미(부/모/미상/ND 등)를 붙임
    - 단위/미상 보조칸: (단위)(미상)"""
    base = f.get("label") or f["name"]
    up = f["name"].upper()
    # 1) 날짜/시각/생년월일 분할 단위
    for sfx, u in NAME_UNIT:
        if up.endswith(sfx):
            return f"{base} ({u})"
    # 2) 단일 체크박스/라디오: 옆에 적힌 의미를 붙임
    if f["type"] in ("radio", "checkbox") and len(f.get("choices") or []) == 1:
        lab = (f["choices"][0].get("label") or "").strip()
        if lab:
            return f"{base} ({lab})"
    # 3) 단위/미상 텍스트칸
    if up.endswith("_UNIT"):
        return f"{base} (단위)"
    if up.endswith("_UK"):
        return f"{base} (미상)"
    # 4) 짧은 꼬리표(시/분/년/월/일/세)
    s = (f.get("suffix") or "").strip()
    if s in SHORT_UNITS:
        return f"{base} ({s})"
    return base


def labeled_fields(fields):
    """user_entry 필드들에 표시 라벨을 부여하고, 같은 라벨은 [필드명]으로 구분."""
    out, seen = [], {}
    for f in fields:
        if not f["user_entry"]:
            continue
        d = display_label(f)
        if d in seen:
            seen[d] += 1
            d = f"{d} [{f['name']}]"
        else:
            seen[d] = 1
        out.append((f, d))
    return out


def dv_for_field(f):
    """선택지가 적은 코드필드에 한해 드롭다운(라벨 기준) 생성."""
    labels = []
    if f["type"] in ("radio", "checkbox"):
        labels = [c["label"] for c in f["choices"] if c["label"]]
    elif f["type"] == "select":
        # 자리표시(placeholder) 옵션 제외
        labels = [o["label"] for o in f["options"]
                  if o["value"] not in ("", "XXXX", "XX", "--", "XX:-", "--:-")
                  and o["label"] not in ("Y", "M", "D", "-")]
    if not labels or len(labels) > 20:
        return None
    if any("," in x or '"' in x for x in labels):
        return None
    formula = '"' + ",".join(labels) + '"'
    if len(formula) > 255:
        return None
    dv = DataValidation(type="list", formula1=formula, allow_blank=True)
    dv.error = "코드북의 값 중에서 선택하세요"
    dv.errorTitle = "값 확인"
    return dv


def is_single_cb(f):
    """value 없는 '해당시 체크' 단일 체크박스(부/모/형제 같은 칸) 여부."""
    return (f["type"] == "checkbox"
            and len(f.get("choices") or []) == 1
            and bool((f["choices"][0].get("label") or "").strip()))


def group_prefix(name):
    """그룹 공통 접두어. 끝 숫자(HTNTX1->HTNTX) 또는 끝 _세그먼트(FHX_CA_F->FHX_CA)를 떼어낸다."""
    p = re.sub(r"\d+$", "", name)
    if p == name and "_" in name:
        p = name.rsplit("_", 1)[0]
    return p


def grouped_columns(fields):
    """user_entry 필드를 '컬럼 단위'로 묶는다.
    같은 질문(label)·같은 접두어의 단일 체크박스가 2개 이상 연속하면 하나의 그룹으로.
    반환 항목: ("single", field) 또는 ("group", 라벨, [member_fields])"""
    uf = [f for f in fields if f["user_entry"]]
    cols, i = [], 0
    while i < len(uf):
        f = uf[i]
        if is_single_cb(f):
            j = i
            while (j < len(uf) and is_single_cb(uf[j])
                   and uf[j]["label"] == f["label"]
                   and group_prefix(uf[j]["name"]) == group_prefix(f["name"])):
                j += 1
            if j - i >= 2:
                cols.append(("group", f.get("label") or f["name"], uf[i:j]))
                i = j
                continue
        cols.append(("single", f))
        i += 1
    return cols


def dv_for_group(members):
    """그룹 드롭다운: 멤버 라벨(부/모/...) 목록. 드롭다운을 주되,
    복수선택(쉼표 입력)을 막지 않도록 검증 오류는 띄우지 않는다."""
    labels = [(m["choices"][0].get("label") or "").strip() for m in members]
    labels = [x for x in labels if x]
    if not labels or len(labels) > 30:
        return None
    if any("," in x or '"' in x for x in labels):
        return None
    formula = '"' + ",".join(labels) + '"'
    if len(formula) > 255:
        return None
    dv = DataValidation(type="list", formula1=formula,
                        allow_blank=True, showErrorMessage=False)
    return dv


def write_columns(ws, fields, n_rows):
    """area/patient 시트 공통: B열부터 컬럼(단일/그룹)을 쓰고 환자키를 채운다.
    반환: 입력 칸(컬럼) 개수."""
    seen, col = {}, 2
    for spec in grouped_columns(fields):
        if spec[0] == "single":
            f = spec[1]
            disp = display_label(f)
            if disp in seen:
                seen[disp] += 1
                disp = f"{disp} [{f['name']}]"
            else:
                seen[disp] = 1
            _write_header(ws, col, disp, f["name"])
            dv = dv_for_field(f)
        else:
            _, glabel, members = spec
            disp = with_context(glabel, members[0].get("context"))
            if disp in seen:
                seen[disp] += 1
                # 같은 질문의 라디오(예/아니오/미상) 옆 세부 체크박스 묶음 → "(상세)"
                alt = f"{disp} (상세)"
                if alt in seen:
                    alt = f"{disp} [{group_prefix(members[0]['name'])}]"
                disp = alt
                seen[disp] = seen.get(disp, 0) + 1
            else:
                seen[disp] = 1
            name_text = GROUP_PREFIX + ",".join(m["name"] for m in members)
            _write_header(ws, col, disp, name_text)
            dv = dv_for_group(members)
        if dv is not None:
            ws.add_data_validation(dv)
            colL = get_column_letter(col)
            dv.add(f"{colL}3:{colL}{2 + n_rows}")
        col += 1
    for i in range(n_rows):
        ws.cell(row=3 + i, column=1, value=i + 1)
    ws.freeze_panes = "B3"
    return col - 2


def _write_header(ws, col, label, name_text):
    """1행=라벨, 2행=필드명(또는 그룹마커) 헤더 작성."""
    c1 = ws.cell(row=1, column=col, value=label)
    c1.fill = HEADER_FILL
    c1.font = WHITE_BOLD
    c1.alignment = CENTER
    c1.border = BORDER
    c2 = ws.cell(row=2, column=col, value=name_text)
    c2.fill = NAME_FILL
    c2.font = NAME_FONT
    c2.alignment = CENTER
    c2.border = BORDER
    ws.column_dimensions[get_column_letter(col)].width = max(12, min(28, len(str(label)) + 4))


def write_field_header(ws, col, f, label=None):
    """1행=라벨, 2행=필드명 헤더 작성."""
    if label is None:
        label = display_label(f)
    _write_header(ws, col, label, f["name"])


def add_key_columns(ws, label_text):
    c1 = ws.cell(row=1, column=1, value=label_text)
    c1.fill = KEY_FILL
    c1.font = Font(bold=True, size=10)
    c1.alignment = CENTER
    c1.border = BORDER
    c2 = ws.cell(row=2, column=1, value="__KEY__")
    c2.fill = NAME_FILL
    c2.font = NAME_FONT
    c2.alignment = CENTER
    c2.border = BORDER
    ws.column_dimensions["A"].width = 10


def build_area_sheet(wb, area, sheet_title, n_rows=30):
    ws = wb.create_sheet(sheet_title)
    add_key_columns(ws, "환자키")
    return write_columns(ws, area["fields"], n_rows)  # 입력 칸 개수


def build_codebook(wb, schema):
    ws = wb.create_sheet("코드북")
    heads = ["영역", "질문(라벨)", "필드명", "유형", "허용값 (코드 = 의미)"]
    for j, h in enumerate(heads, start=1):
        c = ws.cell(row=1, column=j, value=h)
        c.fill = HEADER_FILL
        c.font = WHITE_BOLD
        c.alignment = CENTER
        c.border = BORDER
    widths = [14, 34, 20, 10, 70]
    for j, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(j)].width = w

    def put_row(r, row):
        for j, val in enumerate(row, start=1):
            c = ws.cell(row=r, column=j, value=val)
            c.border = BORDER
            c.alignment = Alignment(vertical="center", wrap_text=(j in (2, 5)))

    r = 2
    for area in schema["areas"]:
        for spec in grouped_columns(area["fields"]):
            if spec[0] == "group":
                _, glabel, members = spec
                labels = [(m["choices"][0].get("label") or "").strip() for m in members]
                desc = "택1 또는 복수(쉼표): " + " | ".join(labels)
                gdisp = with_context(glabel, members[0].get("context"))
                put_row(r, [area["title"], gdisp, "(드롭다운)", "드롭다운", desc])
                r += 1
                continue
            f = spec[1]
            pairs = []
            if f["type"] in ("radio", "checkbox"):
                pairs = [(c["value"], c["label"]) for c in f["choices"]]
            elif f["type"] == "select":
                pairs = [(o["value"], o["label"]) for o in f["options"]
                         if o["value"] not in ("", "XXXX", "XX", "--", "XX:-", "--:-")]
            if not pairs:
                continue  # 자유입력(텍스트/날짜/숫자)은 코드북에 안 실음
            if len(pairs) > 40:
                desc = f"(목록 {len(pairs)}개: 연/시각 등 숫자 선택)"
            else:
                desc = " | ".join(f"{v} = {t}" for v, t in pairs)
            put_row(r, [area["title"], display_label(f), f["name"], f["type"], desc])
            r += 1
    ws.freeze_panes = "A2"


def build_usage(wb):
    ws = wb.create_sheet("사용법")
    lines = [
        "KOCARC 환자 자동등록 입력 양식",
        "",
        "■ 작성 방법",
        "1) [환자목록] 시트에 환자를 한 줄씩 적습니다. (환자키 = 1,2,3 ... 자동)",
        "2) 각 영역 시트(공통영역, 예방역학 ...)에서 같은 '환자키' 행에 값을 채웁니다.",
        "3) 값은 [코드북] 시트를 참고하세요. 코드(예: M) 또는 의미(예: Male) 둘 다 입력 가능합니다.",
        "4) 해당 없는 칸은 비워두면 됩니다. (빈칸은 입력하지 않음)",
        "",
        "■ 시트 구조",
        " - 1행 = 질문(한글 라벨)",
        " - 2행 = 필드명(프로그램이 읽는 줄) → 절대 수정하지 마세요.",
        " - 3행부터 = 환자 데이터",
        "",
        "■ 날짜/시각",
        " - 날짜 칸은 2026-06-15 형식(YYYY-MM-DD)으로 적으세요.",
        " - 시각은 _HOUR(시), _MIN(분) 칸에 각각 적습니다. (코드북 참고)",
        "",
        "■ 주의",
        " - 비밀번호는 이 파일에 적지 마세요.",
        " - 자동 등록은 실제 연구 DB에 저장되므로, 처음에는 연습 환자 1~2명으로 시험하세요.",
    ]
    for i, t in enumerate(lines, start=1):
        c = ws.cell(row=i, column=1, value=t)
        if i == 1:
            c.font = Font(bold=True, size=14, color="2E6B4F")
        elif t.startswith("■"):
            c.font = Font(bold=True, size=11, color="2E6B4F")
    ws.column_dimensions["A"].width = 95


def build(out_path=None, schema_path=None, log=print):
    out_path = out_path or OUT
    schema = load_schema(schema_path)
    areas = {a["key"]: a for a in schema["areas"]}

    wb = Workbook()
    wb.remove(wb.active)  # 기본 시트 제거

    build_usage(wb)

    # 환자목록 = 환자등록(시작) 영역
    pa = areas.get("patient_add")
    ws = wb.create_sheet("환자목록")
    add_key_columns(ws, "환자키")
    pa_count = write_columns(ws, pa["fields"], 30)

    # 나머지 영역 시트 (입력칸이 없는 영역은 시트 생성 안 함)
    counts = {}
    for key in ["common", "prevent", "community", "relief", "in_hosp",
                "alive_after", "heart", "y_child", "comment"]:
        if key not in areas:
            continue
        if not any(f["user_entry"] for f in areas[key]["fields"]):
            continue
        counts[key] = build_area_sheet(wb, areas[key], SHEET_NAMES[key])

    build_codebook(wb, schema)

    wb.save(out_path)
    log(f"엑셀 양식 저장: {out_path}")
    total = pa_count + sum(counts.values())
    log(f"전체 입력 칸: {total}개 (환자목록 {pa_count} + 영역들)")
    return out_path


def main():
    build()


if __name__ == "__main__":
    main()
