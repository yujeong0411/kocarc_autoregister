# -*- coding: utf-8 -*-
"""
KOCARC eCRF 폼 스키마 추출기.
저장된 HTML 페이지들을 파싱하여 영역별 입력 폼 구조를 schema.json 으로 출력한다.

각 필드에 대해 다음을 추출:
  - name      : 폼 필드 이름 (서버로 전송되는 키)
  - type      : text / select / radio / checkbox / textarea / hidden
  - label     : 화면에 보이는 질문 라벨 (예: "성별")
  - section   : 소속 구획 제목 (예: "핵심변수")
  - readonly  : 회색 자동표시 칸 여부 (사람이 입력 안 함)
  - hidden    : 숨김 시스템 필드 여부
  - system    : 봇이 자동 처리하는 시스템 필드 (PAT_ID, WORKMODE 등)
  - user_entry: 사람이 엑셀에 채워야 하는 칸 여부
  - value     : HTML에 들어있던 예시(현재) 값
  - widget    : date / time_hour / time_min / number / text 등 입력 형태 힌트
  - options   : select 의 (value,label) 목록
  - choices   : radio/checkbox 의 (value,label) 목록
"""
import os
import re
import json
from bs4 import BeautifulSoup, Comment


def norm(s):
    """공백/줄바꿈/&nbsp; 정리 및 짝이 맞지 않는 꼬리/머리 괄호 제거.
    (균형 잡힌 괄호 '나이(만)', '기타(직접입력)' 는 보존)"""
    if not s:
        return ""
    s = s.replace("\xa0", " ")
    s = re.sub(r"\s+", " ", s).strip()
    # 뒤따르는 하위입력 표시 '(' 제거
    while s.endswith("(") and s.count("(") > s.count(")"):
        s = s[:-1].strip()
    # 짝 없는 닫힘 ')' 제거 (예: '구급차)', '미상)')
    while s.endswith(")") and s.count(")") > s.count("("):
        s = s[:-1].strip()
    # 맨 앞 짝 없는 열림 '(' 제거
    while s.startswith("(") and s.count("(") > s.count(")"):
        s = s[1:].strip()
    return s

BASE = os.path.dirname(os.path.abspath(__file__))

# 파일 -> (영역키, 한글 제목)
AREAS = [
    ("patient_add", "환자등록(시작)",       "new_case.html"),
    ("common",      "1.공통영역",            "new_case1.html"),
    ("prevent",     "2.예방역학영역",        "new_case2.html"),
    ("community",   "3.지역사회영역",        "new_case3.html"),
    ("relief",      "4.구급단계영역",        "new_case4.html"),
    ("in_hosp",     "5.병원단계영역",        "new_case5.html"),
    ("alive_after", "6.소생후단계영역",      "new_case6.html"),
    ("heart",       "7.심장검사영역",        "new_case7.html"),
    ("y_child",     "8.소아소생술영역",      "new_case8.html"),
    ("comment",     "Comment Log",          "new_case9.html"),
]

# 구조 탐지(detect_trees)로 못 잡는 조건부 트리를 수동 등록.
# (부모·자식이 다른 셀/행에 있거나, 자식 일부 옵션이 항상 활성이라 자동탐지 제외되는 경우)
# 형식: {영역키: {부모필드: {부모값: 자식필드}}}
MANUAL_TREES = {
    # 음주: ALCOHOL=1(있음) 이면 ALCOHOL_DOSE(빈도) 활성 — 다른 행이라 자동탐지 못 함.
    "prevent": {"ALCOHOL": {"1": "ALCOHOL_DOSE"}},
    # 병원치료결과=생존퇴원(1) 이면 퇴원처(HOSP_RESULT_OUT: 자택/타병원/호스피스/요양시설)
    # 를 부모 드롭다운으로 통합 → '생존퇴원 - 자택' 식 잎. (사망퇴원/입원중은 자식 없음)
    "common": {"HOSP_RESULT": {"1": "HOSP_RESULT_OUT"}},
    # CAG 결과: 자동탐지가 못 잡음(부모 CAG_RE 자체가 CAG_DONE에 의해 disabled 로 시작).
    # 3=Significant Stenosis→SS(1VD..), 4=Vasospasm→VASOS(LAD..) 를 부모칸 드롭다운으로 통합.
    # (88=기타는 텍스트 CAG_RE_ETC 라 트리 제외 → '기타' 잎 + 별도 자유입력칸 유지)
    "heart": {
        "CAG_RE1": {"3": "CAG_RE_SS1", "4": "CAG_RE_VASOS1"},
        "CAG_RE2": {"3": "CAG_RE_SS2", "4": "CAG_RE_VASOS2"},
        "CAG_RE3": {"3": "CAG_RE_SS3", "4": "CAG_RE_VASOS3"},
    },
    # 소아 심정지 원인: 3단계 중첩 라디오를 드롭다운 1칸으로 통합(사용자 확정).
    #   ARREST_CA=Medical(1)→ARREST_CA_MEDI, Asphyxial(6)→ARREST_CA_ASPH
    #   ARREST_CA_MEDI=Presumed Cardiac(1)→_PC, Other Medical(3)→_OME (손자 레벨)
    # 자식이 또 부모라 tree_leaves/bot tree_targets 가 재귀로 펼침·복원.
    "y_child": {
        "ARREST_CA": {"1": "ARREST_CA_MEDI", "6": "ARREST_CA_ASPH"},
        "ARREST_CA_MEDI": {"1": "ARREST_CA_MEDI_PC", "3": "ARREST_CA_MEDI_OME"},
    },
}

# 파싱이 라벨을 잘못 가져오는 선택지 라벨 교정 (중첩 텍스트로 오염되는 경우).
# 형식: {영역키: {필드: {value: 교정라벨}}}
CHOICE_OVERRIDES = {
    # 다중출동: HTML이 '단일출동 다중출동(펌블런스/구급차) 미상' 구조라 value=1 라벨이
    # '단일출동 다중출동'으로 오염됨. 사용자 혼동 방지 위해 명확히 교정(봇도 이 라벨로 매칭).
    "relief": {"MULTI_MOVE": {"1": "단일출동", "2": "다중출동-펌블런스",
                              "3": "다중출동-구급차", "99": "미상"}},
}

# 같은 셀에 묶여 라벨이 뭉뚱그려진 필드를 명확히 지정. 형식: {영역키: {필드: 라벨}}
FIELD_LABEL_OVERRIDES = {
    # Na-K-Cl 한 셀에 값칸 3개 → 각각 Na / K / Cl 로 분명하게.
    "in_hosp": {"NA": "Na (mmol/L)", "K": "K (mmol/L)", "CL": "Cl (mmol/L)"},
    # GCS 검사결과 3칸 = E(개안)/V(언어)/M(운동) — 직후·72시간째 동일.
    "alive_after": {
        "GCS_E_RESULT": "E (개안)", "GCS_V_RESULT": "V (언어)", "GCS_M_RESULT": "M (운동)",
        "GCS72_E_RESULT": "E (개안)", "GCS72_V_RESULT": "V (언어)", "GCS72_M_RESULT": "M (운동)",
    },
}

# 같은 상위구획(CAG 등)에 시기별로 똑같은 필드가 반복될 때, 컨텍스트에 시기를 넣어 구분.
# (예: CAG 결과/기타가 24h·24-72h·7일-퇴원전 3번 → [CAG_RE2] 대신 시기명으로 표시)
FIELD_CONTEXT_OVERRIDES = {
    "heart": {
        "CAG_RE1": "CAG (24시간 이내)", "CAG_RE_ETC1": "CAG (24시간 이내)",
        "CAG_RE2": "CAG (24-72시간 이내)", "CAG_RE_ETC2": "CAG (24-72시간 이내)",
        "CAG_RE3": "CAG (7일-퇴원 전)", "CAG_RE_ETC3": "CAG (7일-퇴원 전)",
    },
    # 병원퇴원/사망 일시: 라벨에 이미 '(생존/사망퇴원인 경우)'가 있어 '병원치료결과 -'
    # 접두어가 중복이라 컬럼이 너무 길어짐 → context 비워 접두어 제거.
    # (@DT 합친칸 헤더는 _DATE 필드의 라벨/context만 사용하므로 _DATE만 비우면 됨.)
    "common": {"HOSP_OUT_DATE": "", "HOSP_DIE_DATE": ""},
}

# 봇이 자동으로 채우는 시스템 필드 (사람이 입력하지 않음)
SYSTEM_FIELDS = {
    "PAT_ID", "WORKMODE", "goURL", "HOSP_CD", "HOSP_NM", "RAND_GRP",
    "sVISIT_DATE", "G_LOCAL_PAT_ID",
    # 입력완료 체크용 시스템 필드 (데이터 아님)
    "COMPLETE_YN", "COMPLETE_YN2",
}

# 환자등록(앞단계)에서 자동으로 넘어오는 칸: 해당 영역에서 다시 입력할 필요 없음
CARRIED_OVER = {
    "common": {"NEW_PAT_NM", "SEX", "DOB_YY", "DOB_MM", "DOB_DD",
               "ER_DATE", "ER_CLOCK_HOUR", "ER_CLOCK_MIN"},
}

LABEL_CLASSES = ("tdlb", "boxtitle", "tdl", "tdcb", "tdcb2")
SECTION_CLASSES = ("tdcb", "tdcb2")


def read_html(path):
    with open(path, "rb") as f:
        raw = f.read()
    # 저장된 페이지는 UTF-8 (브라우저 소스보기 저장). UTF-8을 먼저 시도.
    for enc in ("utf-8", "euc-kr", "cp949"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", "replace")


def trailing_text(el):
    """입력칸 바로 뒤에 붙은 짧은 텍스트(시/분/년/월/일/세 등)."""
    for sib in el.next_siblings:
        if getattr(sib, "name", None) in ("input", "select", "textarea", "br"):
            break
        t = sib if isinstance(sib, str) else sib.get_text(" ", strip=True)
        t = norm(t or "")
        if t:
            return t
    return ""


def cls_of(tag):
    return " ".join(tag.get("class", []) or [])


SKIP_CONTEXT = {"핵심변수", "선택변수"}


def find_label(el):
    """입력칸의 질문 라벨과 상위 컨텍스트를 찾는다. 반환: (label, context)

    칸이 중첩 테이블 안에 있어도, 조상 셀(td)들을 바깥쪽으로 거슬러 올라가며
    각 셀의 앞 형제 라벨셀(tdlb/boxtitle)을 찾는다. 그래도 없으면 문서상
    가장 가까운 앞쪽 라벨셀을 사용(rowspan 등 대응).

    중첩 표(예: 과거력의 '고혈압' > '병원에서 진단/치료')에서는
    안쪽 라벨을 label, 바깥쪽 라벨(질병명 등)을 context 로 분리한다."""
    td = el.find_parent("td")
    if td is None:
        return "", ""
    # 1) 조상 td들을 안쪽→바깥쪽으로: 각 td의 '같은 줄 앞 형제' 중 라벨셀
    #    라벨 셀 class 가 일관되지 않음(tdlb/boxtitle/tdcb/tdcb2 혼용).
    #    같은 줄의 앞 형제일 때만 tdcb 도 라벨로 인정(구획 제목은 별도 행이라 안 잡힘).
    found = []  # 안쪽 -> 바깥쪽 순
    cur = td
    while cur is not None:
        for prev in cur.find_previous_siblings("td"):
            c = cls_of(prev)
            if any(k in c for k in ("tdlb", "boxtitle", "tdcb")):
                txt = norm(prev.get_text(" ", strip=True))
                if txt:
                    found.append(txt)
                    break
        cur = cur.find_parent("td")
    if found:
        label = found[0]
        # 바깥쪽 라벨 중 구획제목(핵심/선택변수)·라벨중복을 뺀 것을 컨텍스트로
        outer = [t for t in found[1:] if t not in SKIP_CONTEXT and t != label]
        context = outer[-1] if outer else ""
        return label, context
    # 2) 폴백: 문서상 가장 가까운 앞쪽 라벨셀
    for cell in el.find_all_previous("td"):
        c = cls_of(cell)
        if any(k in c for k in ("tdlb", "boxtitle")):
            txt = norm(cell.get_text(" ", strip=True))
            if txt:
                return txt, ""
    return "", ""


def find_section(el):
    """입력칸 위쪽에서 가장 가까운 구획 제목(tdcb)을 찾는다."""
    for cell in el.find_all_previous("td"):
        c = cls_of(cell)
        if "tdcb" in c:
            txt = norm(cell.get_text(" ", strip=True))
            if txt:
                return txt
    return ""


def strip_nav(s):
    """'모름 if 있음 ↓' / '미상 if 예 ↓' 같은 화면 이동안내를 라벨에서 제거."""
    s = re.sub(r"\s*if\b[^↓]*↓\s*", " ", s)
    return s.replace("↓", "").strip()


def strip_unit_ph(s):
    """단위 선택 드롭다운이 라벨로 딸려온 '( 선택 mmol/L mg/dL )' 류를 제거.
    (혈액검사 값칸 라벨을 'iCa ( 선택 mmol/L mg/dL )' → 'iCa' 로 깔끔히.)"""
    if not s:
        return s
    return norm(re.sub(r"\(\s*선택[^)]*\)", "", s))


def value_label(el):
    """radio/checkbox 입력 바로 뒤에 붙은 설명 텍스트."""
    parts = []
    for sib in el.next_siblings:
        nm = getattr(sib, "name", None)
        if nm in ("input", "select", "textarea", "br"):
            break
        if isinstance(sib, str):
            t = sib.strip()
        else:
            t = sib.get_text(" ", strip=True)
        if t:
            parts.append(t)
    return strip_nav(norm(" ".join(parts)))


def guess_widget(name, el, ftype):
    onclick = (el.get("onclick") or "")
    up = name.upper()
    # _DATE/_HOUR/_MIN 뒤에 번호가 붙어도(예: SSEP_DATE1) 일시 위젯으로 인식
    if "Calendar" in onclick or re.search(r"_DATE\d*$", up):
        return "date"
    if re.search(r"_HOUR\d*$", up):
        return "time_hour"
    if re.search(r"_MIN\d*$", up):
        return "time_min"
    oc = (el.get("onkeyup") or "") + (el.get("onblur") or "")
    if "onlynumber" in oc or "checkNumber" in oc:
        return "number"
    return ftype


def region_of_fields(scope):
    """문서 순서대로 '핵심변수'/'선택변수' 구획을 추적해, 각 필드가 어느 구획에
    처음 등장하는지(core/optional) 판정한다."""
    region = "core"
    field_region = {}
    for el in scope.find_all(["td", "input", "select", "textarea"]):
        if el.name == "td":
            t = norm(el.get_text(" ", strip=True))
            if t == "선택변수":
                region = "optional"
            elif t == "핵심변수":
                region = "core"
        else:
            nm = el.get("name")
            if nm and nm not in field_region:
                field_region[nm] = region
    return field_region


def detect_trees(scope):
    """라디오→라디오 조건부 트리 탐지.

    한 셀(td) 안에서 '처음부터 활성인 라디오'가 정확히 1개(부모)이고,
    '전부 disabled 로 시작하는 라디오'가 1개 이상(자식)일 때만 트리로 본다.
    (사이트는 부모 특정값 클릭 전까지 자식칸을 disabled 로 둔다.)
    자식의 트리거 부모값 = 자식 첫 input 직전에 등장한 부모 라디오의 value.

    반환: (tree_map, child_names)
      tree_map    = {부모필드: {부모값: 자식필드}}
      child_names = {모든 자식 필드}  (별도 컬럼 대신 부모칸에 합쳐짐)
    """
    tree_map, child_names = {}, set()
    for td in scope.find_all("td"):
        radios = [r for r in td.find_all("input", attrs={"type": "radio"})
                  if r.find_parent("td") is td]  # 직계 라디오만
        if len(radios) < 2:
            continue
        names = []
        for r in radios:
            nm = r.get("name")
            if nm and nm not in names:
                names.append(nm)
        disabled = {nm: all(r.has_attr("disabled")
                            for r in radios if r.get("name") == nm)
                    for nm in names}
        enabled = [nm for nm in names if not disabled[nm]]
        kids = [nm for nm in names if disabled[nm]]
        if len(enabled) != 1 or not kids:
            continue  # 부모 1개 + 자식 1개이상 인 셀만
        parent = enabled[0]
        cur_pv, mapping = None, {}
        for r in radios:
            nm = r.get("name")
            if nm == parent:
                cur_pv = r.get("value", "")
            elif nm in kids and cur_pv is not None:
                # ponytail: 부모값 1개당 자식 1개(1:1)만 지원. 현재 사이트 전부 1:1.
                mapping.setdefault(cur_pv, nm)
        if mapping:
            tree_map[parent] = mapping
            child_names.update(mapping.values())
    return tree_map, child_names


def extract_area(key, title, fname):
    path = os.path.join(BASE, fname)
    soup = BeautifulSoup(read_html(path), "lxml")
    # 주석(<!-- -->)으로 비활성화된 옵션/필드는 실제 사이트에 없다.
    # 주석 안 HTML이 라벨로 딸려 들어가는 것도 막기 위해 통째로 제거.
    for cmt in soup.find_all(string=lambda s: isinstance(s, Comment)):
        cmt.extract()
    form = soup.find("form")
    action = (form.get("action") if form else "") or ""
    action = action.split("?", 1)[0]  # 쿼리스트링(?G_LOCAL_PAT_ID=… 등 PHI) 제거
    enctype = form.get("enctype", "") if form else ""
    scope = form if form else soup
    field_region = region_of_fields(scope)
    tree_map, tree_children = detect_trees(scope)
    # 수동 등록 트리 병합
    for pname, mp in MANUAL_TREES.get(key, {}).items():
        tree_map.setdefault(pname, {}).update(mp)
        tree_children.update(mp.values())

    fields = {}   # name -> field dict (순서 유지)
    order = []
    for el in scope.find_all(["input", "select", "textarea"]):
        name = el.get("name")
        if not name:
            continue
        tag = el.name
        ftype = el.get("type", tag).lower() if tag == "input" else tag
        is_hidden = (ftype == "hidden")
        is_readonly = el.has_attr("readonly")
        is_disabled = el.has_attr("disabled")

        if name not in fields:
            fld = {
                "name": name,
                "type": ftype,
                "label": "",
                "context": "",
                "section": "",
                "readonly": is_readonly,
                "hidden": is_hidden,
                "system": (name in SYSTEM_FIELDS) or is_hidden,
                "value": "",
                "widget": ftype,
                "suffix": "",
                "options": [],
                "choices": [],
            }
            fields[name] = fld
            order.append(name)
        fld = fields[name]

        if not is_hidden:
            if not fld["label"]:
                fld["label"], fld["context"] = find_label(el)
                fld["label"] = strip_unit_ph(fld["label"])
                fld["context"] = strip_unit_ph(fld["context"])
            if not fld["section"]:
                fld["section"] = find_section(el)
            fld["widget"] = guess_widget(name, el, ftype)
            if ftype not in ("radio", "checkbox") and not fld["suffix"]:
                fld["suffix"] = trailing_text(el)

        if ftype in ("radio", "checkbox"):
            v = el.get("value", "")
            lab = value_label(el)
            pair = {"value": v, "label": lab}
            if pair not in fld["choices"]:
                fld["choices"].append(pair)
            if el.has_attr("checked") and not fld["value"]:
                fld["value"] = v
        elif tag == "select":
            opts = []
            for o in el.find_all("option"):
                ov = o.get("value", "")
                ot = norm(o.get_text(strip=True))
                # 자리표시(placeholder) 옵션 제외: XXXX, XX, --, 빈값
                opts.append({"value": ov, "label": ot})
                if o.has_attr("selected") and not fld["value"]:
                    fld["value"] = ov
            fld["options"] = opts
        else:  # text / textarea / hidden 등
            # 개인정보 보호: 자유입력 칸의 예시값(환자 실제값)은 저장하지 않음
            pass

    # 선택지 라벨 교정 (오염된 라벨 바로잡기)
    for fn, mp in CHOICE_OVERRIDES.get(key, {}).items():
        f = fields.get(fn)
        if f:
            for ch in f["choices"]:
                if ch["value"] in mp:
                    ch["label"] = mp[ch["value"]]

    # 필드 라벨 직접 지정 (같은 셀에 묶인 값칸 구분 등)
    for fn, lab in FIELD_LABEL_OVERRIDES.get(key, {}).items():
        if fn in fields:
            fields[fn]["label"] = lab

    # 필드 컨텍스트 직접 지정 (시기별 반복 구획 구분 등)
    for fn, ctx in FIELD_CONTEXT_OVERRIDES.get(key, {}).items():
        if fn in fields:
            fields[fn]["context"] = ctx

    # 단위 선택칸(_UNIT)은 값칸과 같은 셀에 있어 라벨이 옆 항목으로 오검출됨 →
    # 짝 값칸('{이름}' = _UNIT 뗀 것)의 라벨을 가져와 명확히 한다. (예: ICA_UNIT ← ICA)
    for n in list(fields):
        if n.endswith("_UNIT"):
            base = fields.get(n[:-5])
            if base and base["label"]:
                fields[n]["label"] = base["label"]
                fields[n]["context"] = base["context"]

    # user_entry 판정
    out_fields = []
    for n in order:
        f = fields[n]
        f["region"] = field_region.get(n, "core")
        # readonly 라도 '달력으로 입력하는 날짜칸'은 사용자가 채우는 값이므로 포함
        f["user_entry"] = (not f["hidden"]) and (
            (not f["readonly"]) or f["widget"] == "date")
        # 선택변수 구획은 제외 (핵심변수만 입력)
        if f["region"] == "optional":
            f["user_entry"] = False
        # 파일 업로드 칸: 엑셀 일괄자동화로 환자별 파일 첨부 불가 → 제외
        if f["type"] == "file":
            f["user_entry"] = False
        # 시스템 필드(COMPLETE_YN 등) 제외
        if f["name"] in SYSTEM_FIELDS:
            f["user_entry"] = False
            f["system"] = True
        # 앞단계에서 자동으로 넘어오는 칸은 제외
        if f["name"] in CARRIED_OVER.get(key, set()):
            f["user_entry"] = False
            f["system"] = True
        # 조건부 트리: 부모칸에 트리 정보 기록, 자식칸은 부모칸에 합쳐지므로 컬럼 제외
        if f["name"] in tree_map:
            f["tree"] = tree_map[f["name"]]
        if f["name"] in tree_children:
            f["user_entry"] = False
            f["tree_child"] = True
        out_fields.append(f)

    return {
        "key": key,
        "title": title,
        "file": fname,
        "action": action,
        "enctype": enctype,
        "fields": out_fields,
    }


def main():
    areas = []
    for key, title, fname in AREAS:
        if not os.path.exists(os.path.join(BASE, fname)):
            print(f"[건너뜀] {fname} 없음")
            continue
        area = extract_area(key, title, fname)
        areas.append(area)

    schema = {"areas": areas}
    with open(os.path.join(BASE, "schema.json"), "w", encoding="utf-8") as f:
        json.dump(schema, f, ensure_ascii=False, indent=2)

    # 요약 출력
    print(f"{'영역':22s} {'전체':>5s} {'입력칸':>6s} {'자동/숨김':>8s}")
    print("-" * 50)
    g_total = g_user = 0
    for a in areas:
        total = len(a["fields"])
        user = sum(1 for f in a["fields"] if f["user_entry"])
        g_total += total
        g_user += user
        print(f"{a['title']:22s} {total:5d} {user:6d} {total-user:8d}")
    print("-" * 50)
    print(f"{'합계':22s} {g_total:5d} {g_user:6d} {g_total-g_user:8d}")
    print("\nschema.json 저장 완료")


if __name__ == "__main__":
    main()
