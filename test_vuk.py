# -*- coding: utf-8 -*-
"""@VUK(값+미상 통합) 및 MULTI_MOVE 라벨 교정 검증.
- 봇 resolver 가 교정된 MULTI_MOVE 라벨을 올바른 값으로 매칭하는지
- fill_valuk 분기 로직(미상→체크박스, 그 외→값입력)이 맞는지
실행: uv run python test_vuk.py"""
import json
import os
from kocarc_bot import Resolver, VUK_PREFIX

BASE = os.path.dirname(os.path.abspath(__file__))
with open(os.path.join(BASE, "schema.json"), encoding="utf-8") as f:
    R = Resolver(json.load(f))

# MULTI_MOVE: 교정된 라벨 -> 값
assert R.to_form_value("relief", "MULTI_MOVE", "단일출동") == "1"
assert R.to_form_value("relief", "MULTI_MOVE", "다중출동-펌블런스") == "2"
assert R.to_form_value("relief", "MULTI_MOVE", "다중출동-구급차") == "3"
assert R.to_form_value("relief", "MULTI_MOVE", "미상") == "99"
assert R.to_form_value("relief", "MULTI_MOVE", "2") == "2"  # 코드 직접입력도 허용


def valuk_action(marker, raw, area="relief"):
    """fill_valuk 분기와 동일한 판정(드라이버 없이 로직만 검증)."""
    parts = marker[len(VUK_PREFIX):].split(",")
    val_field, uk_field = parts[0], parts[1]
    keyword = parts[2] if len(parts) > 2 else "미상"
    if str(raw).strip().lower() == keyword.lower():
        return ("check_uk", uk_field)
    return ("set_val", val_field, R.to_form_value(area, val_field, raw))


# 미상(_UK) 통합칸
m1 = "@VUK:DISASTER_NUM,DISASTER_NUM_UK,미상"
assert valuk_action(m1, "미상") == ("check_uk", "DISASTER_NUM_UK")
assert valuk_action(m1, "123456789012") == ("set_val", "DISASTER_NUM", "123456789012")
m2 = "@VUK:PRE_DEFIB_N,PRE_DEFIB_N_UK,미상"
assert valuk_action(m2, 3) == ("set_val", "PRE_DEFIB_N", "3")
assert valuk_action(m2, "미상") == ("check_uk", "PRE_DEFIB_N_UK")

# ND(_YN) 통합칸 (병원단계 혈액검사) — 'ND' 입력시 체크박스, 숫자는 값입력
m3 = "@VUK:WBC,WBC_YN,ND"
assert valuk_action(m3, "ND", "in_hosp") == ("check_uk", "WBC_YN")
assert valuk_action(m3, "nd", "in_hosp") == ("check_uk", "WBC_YN")   # 대소문자 무시
assert valuk_action(m3, "7.2", "in_hosp") == ("set_val", "WBC", "7.2")

# 수동 페어링(DEFIB_CNT+DEFIB_UK, 이름규칙 밖) + Na('ND -' 꼬리 제거되어 ND로 병합)
m4 = "@VUK:DEFIB_CNT,DEFIB_UK,미상"
assert valuk_action(m4, "미상", "in_hosp") == ("check_uk", "DEFIB_UK")
assert valuk_action(m4, 2, "in_hosp") == ("set_val", "DEFIB_CNT", "2")
m5 = "@VUK:NA,NA_YN,ND"
assert valuk_action(m5, "ND", "in_hosp") == ("check_uk", "NA_YN")
assert valuk_action(m5, "140", "in_hosp") == ("set_val", "NA", "140")

print("OK: VUK/MULTI_MOVE 검증 통과")
