# -*- coding: utf-8 -*-
"""Extract all form fields from saved KOCARC eCRF HTML pages."""
import os, re, json
from bs4 import BeautifulSoup

FILES = {
    "patient_add (환자등록)": "new_case.html",
    "common (공통영역/기초정보)": "new_case3.html",
    "prevent (예방역학)": "new_case4.html",
    "community (지역사회)": "new_case5.html",
    "relief (구급단계)": "new_case6.html",
    "in_hosp (병원단계)": "new_case7.html",
    "alive_after (소생후)": "new_case8.html",
    "heart (심장검사)": "new_case9.html",
    "y_child (소아소생술)": "new_case10.html",
    "comment (Comment Log)": "new_case11.html",
}

BASE = os.path.dirname(os.path.abspath(__file__))

def read_html(path):
    with open(path, "rb") as f:
        raw = f.read()
    for enc in ("euc-kr", "cp949", "utf-8"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", "replace")

summary = {}
report = []
for label, fname in FILES.items():
    path = os.path.join(BASE, fname)
    if not os.path.exists(path):
        continue
    soup = BeautifulSoup(read_html(path), "lxml")
    form = soup.find("form")
    action = form.get("action") if form else "(no form)"

    fields = {}  # name -> info
    scope = form if form else soup
    for el in scope.find_all(["input", "select", "textarea"]):
        name = el.get("name")
        if not name:
            continue
        tag = el.name
        ftype = el.get("type", tag) if tag == "input" else tag
        info = fields.setdefault(name, {"type": ftype, "values": [], "options": []})
        if ftype in ("radio", "checkbox"):
            v = el.get("value", "")
            if v not in info["values"]:
                info["values"].append(v)
        if tag == "select":
            opts = []
            for o in el.find_all("option"):
                opts.append((o.get("value", ""), o.get_text(strip=True)))
            info["options"] = opts

    # counts by category
    cats = {}
    for n, i in fields.items():
        cats[i["type"]] = cats.get(i["type"], 0) + 1
    # for radio/checkbox count distinct field names not value rows
    summary[label] = {"action": action, "n_fields": len(fields), "by_type": cats}

    report.append(f"\n{'='*80}\n## {label}   ->  action={action}\n{'='*80}")
    report.append(f"필드 개수(이름 기준): {len(fields)}  | 유형별: {cats}")
    for n, i in fields.items():
        line = f"  - {n}  [{i['type']}]"
        if i["type"] in ("radio", "checkbox") and i["values"]:
            line += f"  values={i['values']}"
        if i["options"]:
            shown = i["options"][:12]
            line += "  options=" + ", ".join(f"{v}:{t}" for v, t in shown)
            if len(i["options"]) > 12:
                line += f" ...(+{len(i['options'])-12})"
        report.append(line)

with open(os.path.join(BASE, "form_fields_report.txt"), "w", encoding="utf-8") as f:
    f.write("\n".join(report))

print("=== SUMMARY (필드 개수) ===")
total = 0
for label, s in summary.items():
    total += s["n_fields"]
    print(f"{label:30s} fields={s['n_fields']:3d}  action={s['action']}")
print(f"{'TOTAL':30s} fields={total}")
print("\n상세는 form_fields_report.txt 에 저장됨")
