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
from bs4 import BeautifulSoup


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
    ("common",      "1.공통영역",            "new_case3.html"),
    ("prevent",     "2.예방역학영역",        "new_case4.html"),
    ("community",   "3.지역사회영역",        "new_case5.html"),
    ("relief",      "4.구급단계영역",        "new_case6.html"),
    ("in_hosp",     "5.병원단계영역",        "new_case7.html"),
    ("alive_after", "6.소생후단계영역",      "new_case8.html"),
    ("heart",       "7.심장검사영역",        "new_case9.html"),
    ("y_child",     "8.소아소생술영역",      "new_case10.html"),
    ("comment",     "Comment Log",          "new_case11.html"),
]

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
    return norm(" ".join(parts))


def guess_widget(name, el, ftype):
    onclick = (el.get("onclick") or "")
    if "Calendar" in onclick or name.upper().endswith("_DATE"):
        return "date"
    up = name.upper()
    if up.endswith("_HOUR"):
        return "time_hour"
    if up.endswith("_MIN"):
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


def extract_area(key, title, fname):
    path = os.path.join(BASE, fname)
    soup = BeautifulSoup(read_html(path), "lxml")
    form = soup.find("form")
    action = form.get("action") if form else ""
    enctype = form.get("enctype", "") if form else ""
    scope = form if form else soup
    field_region = region_of_fields(scope)

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
        # 시스템 필드(COMPLETE_YN 등) 제외
        if f["name"] in SYSTEM_FIELDS:
            f["user_entry"] = False
            f["system"] = True
        # 앞단계에서 자동으로 넘어오는 칸은 제외
        if f["name"] in CARRIED_OVER.get(key, set()):
            f["user_entry"] = False
            f["system"] = True
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
