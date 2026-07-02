# -*- coding: utf-8 -*-
"""공통영역 종속 일시칸 회색 규칙 진리표 검증.
build_template.apply_grey_rules 의 엑셀 수식과 '동일한 로직'을 파이썬으로
재현해, 대표 시나리오에서 흰색(입력)/회색(비움)이 임상 규칙과 맞는지 확인한다.
수식을 고치면 이 표도 같이 고쳐야 한다. 실행: uv run python test_grey.py"""

# 엑셀 수식과 1:1 대응 (grey=True 면 회색). f=심폐소생술1, g=심폐소생술2 통합칸 값.
def stop_grey(f, g):       # 중단일시
    return f == "" or f[:3] == "미시행"
def any_grey(f, g):        # Any ROSC 일시
    return g != "시행 - Sustained ROSC없이 Any ROSC만"
def sus_grey(f, g):        # Sustained ROSC 일시
    return g != "시행 - Sustained ROSC" and f != "시행 20분이내 중단 - 자발순환회복"
def warn(f, g):            # 모순: 심폐1=20분이상 시행인데 심폐2=미시행
    return f == "20분이상 시행" and g == "미시행"
def die_grey(q):           # 응급실 사망일시 (q=응급실 결과 통합칸)
    return q[:2] != "사망"
def batch_grey(q):         # 생존-입원/퇴원 종속 batch (HOSP_AD.../HOSP_RESULT/FU6M_STAT)
    return q not in ("생존 - 입원", "생존 - 퇴원")
def hosp_out_grey(hr):     # 병원퇴원일시(+퇴원형태) (hr=병원치료결과)
    return hr != "생존퇴원"
def hosp_die_grey(hr):     # 병원사망일시
    return hr != "사망퇴원"
def fu6_die_grey(fs):      # 6개월 후 사망일시 (fs=6개월 후 생존)
    return fs != "사망"
def f6_cpc_grey(fs):       # 6개월 후 신경학적 상태
    return fs != "생존"
def hosp_cpc_grey(q, hr):  # 병원퇴원시 신경학적상태 (결과 q + 병원치료결과 hr)
    return batch_grey(q) or hr == "입원중"
def tx_grey(dx):           # 예방역학 치료여부: 병원진단≠예면 회색
    return dx != "예"
def method_grey(tx):       # 예방역학 치료방법: 치료여부≠예면 회색
    return tx != "예"
def onset_grey(w):         # 지역사회 발생추정시각: 목격여부≠목격이면 회색
    return w != "목격"
def etc_grey(parent):      # 지역사회 기타 자유입력칸: 부모≠기타면 회색
    return parent != "기타"
def relief_grey(pval, white):  # 구급단계 종속칸: 부모 활성값≠이면 회색
    return pval != white
def dose_grey(group_val, drug):  # 약물 총 용량: 그룹칸에 해당 약 없으면 회색
    return drug not in group_val

# (심폐1, 심폐2, 기대[중단, Any, Sustained] — True=흰색(입력), 기대경고)
CASES = [
    ("미시행 - DOA",                    "미시행",                              (False, False, False), False),
    ("미시행 - 내원당시 자발순환회복 상태", "미시행",                           (False, False, False), False),
    ("시행 20분이내 중단 - 자발순환회복",  "시행 - Sustained ROSC",              (True,  False, True ), False),
    ("시행 20분이내 중단 - 소생술중단사망","시행 - Any ROSC없이 사망",           (True,  False, False), False),
    ("20분이상 시행",                    "시행 - Sustained ROSC없이 Any ROSC만",(True,  True,  False), False),
    ("20분이상 시행",                    "시행 - Sustained ROSC",              (True,  False, True ), False),
    ("20분이상 시행",                    "미시행",                              (True,  False, False), True ),  # 모순
    ("",                                 "",                                    (False, False, False), False),
]

# (응급실 결과, 기대[응급실사망일시, 생존입원batch] — True=흰색)
RESULT_CASES = [
    ("생존 - 입원",        (False, True )),
    ("생존 - 전원",        (False, False)),
    ("생존 - 퇴원",        (False, True )),  # 입원/퇴원 둘 다 batch 흰색
    ("CPR하면서 전원",     (False, False)),
    ("사망 - 사망",        (True,  False)),
    ("사망 - 가망없는퇴원", (True,  False)),
    ("",                   (False, False)),
]

# (병원치료결과, [병원퇴원일시 white, 병원사망일시 white])
HOSP_CASES = [
    ("생존퇴원", (True,  False)),
    ("사망퇴원", (False, True )),
    ("입원중",   (False, False)),
    ("",         (False, False)),
]
# (6개월 후 생존, [6개월 사망일시 white, 6개월 신경학적 white])
FU6_CASES = [
    ("사망",           (True,  False)),
    ("생존",           (False, True )),
    ("미상(연락두절)",  (False, False)),
    ("",               (False, False)),
]
# (응급실결과, 병원치료결과, 병원퇴원시신경 white)
CPC_CASES = [
    ("생존 - 입원", "생존퇴원", True ),
    ("생존 - 입원", "사망퇴원", True ),
    ("생존 - 입원", "입원중",   False),  # 입원중 → 회색
    ("생존 - 퇴원", "",         True ),  # batch활성·병원결과 미선택 → 흰색
    ("사망 - 사망", "",         False),  # batch 비활성 → 회색
    ("",            "입원중",   False),
]

# 예방역학 치료 연쇄 (병원진단, 치료여부, [치료여부 white, 방법 white])
PREVENT_CASES = [
    ("예",   "예",    (True,  True )),
    ("예",   "아니오", (True,  False)),
    ("아니오", "",     (False, False)),
    ("미상",  "",     (False, False)),
    ("",     "",      (False, False)),
]
# 지역사회 발생추정시각 (목격여부, 흰색?)
COMMUNITY_CASES = [("목격", True), ("비목격", False), ("미상", False), ("", False)]
# 지역사회 기타 자유입력칸 (부모값, 흰색?)  — WIT_PERSON/ONSET_LOC == 기타일 때만
ETC_CASES = [("기타", True), ("일반인", False), ("집/거주시설", False), ("미상", False), ("", False)]
# 구급단계 종속칸 (부모값, 흰색조건값, 흰색?)
RELIEF_CASES = [
    ("시행", "시행", True), ("미시행", "시행", False), ("미상", "시행", False), ("", "시행", False),
    ("적용", "적용", True), ("미적용", "적용", False),
    ("회복", "회복", True), ("미회복", "회복", False),
    ("심정지리듬", "심정지리듬", True), ("일반인소생(정상동율동)", "심정지리듬", False),
    ("기타", "기타", True), ("Asystole", "기타", False),
    ("확보함", "확보함", True), ("아니오", "확보함", False),
    ("사용함", "사용함", True), ("아니오", "사용함", False),
]
# 약물별 총 용량 (약물상세 그룹값, 약이름, 흰색?)
DOSE_CASES = [
    ("에피네프린,바소프레신", "에피네프린", True),
    ("에피네프린,바소프레신", "아미오다론", False),
    ("에피네프린,바소프레신", "바소프레신", True),
    ("", "에피네프린", False),
]
# 병원단계 종속칸 (부모값, 흰색조건값, 흰색?)  — relief_grey 재사용
IN_HOSP_CASES = [
    ("타병원경유", "타병원경유", True), ("직접내원", "타병원경유", False), ("", "타병원경유", False),  # #1 PRE_ROSC(교차시트)
    ("회복", "회복", True), ("미회복", "회복", False), ("자발순환회복 후 타병원 내원", "회복", False),  # #2 회복시각
    ("심정지리듬", "심정지리듬", True), ("생존후 등록병원 내원 (ROSC)", "심정지리듬", False),  # #3 HOSP_ECG_RH
    ("기타", "기타", True), ("E-tube", "기타", False), ("IV", "기타", False),  # #6/#7/#9 기타칸
]
# 에피네프린 두 칸(#4): 미상/미사용 그룹칸 값 있으면 회색 (그룹값, 흰색?)
EPINE_CASES = [("", True), ("미상", False), ("미사용", False), ("미상, 미사용", False)]
# 단위 드롭다운: 짝 값칸이 ND/미상이면 회색 (값칸값, 흰색?)
def unit_grey(v): return v in ("ND", "미상")
UNIT_CASES = [("ND", False), ("미상", False), ("140", True), ("1.2", True), ("", True)]
# Steroid 종류(기타): 종류=기타일 때만 흰색 / 투여량단위: 종류=미시행·미선택이면 회색
def steroid_etc_grey(k): return k != "기타"
def steroid_unit_grey(k): return k in ("", "미시행")
STEROID_CASES = [  # (종류값, [종류기타 흰색?, 투여량단위 흰색?])
    ("Hydrocortisone", (False, True)),
    ("MethylPD",       (False, True)),
    ("Dexamethasone",  (False, True)),
    ("기타",            (True,  True)),
    ("미시행",          (False, False)),
    ("",               (False, False)),
]
# 소생후단계 시트 게이팅: 공통 심폐1=시행20분이내중단-자발순환회복 또는 심폐2=시행-Sustained ROSC
def alive_gate_grey(c1, c2):
    return not (c1 == "시행 20분이내 중단 - 자발순환회복" or c2 == "시행 - Sustained ROSC")
def proc_grey(v):  return v in ("", "미시행", "미상")        # 시술류(REPER/CLOT/ANGIO/CA/CAB)
def hyper_grey(v): return v != "ROSC 24시간 이내"             # 승압제 종류(24시간 이내만 활성)
ALIVE_GATE_CASES = [  # (심폐1, 심폐2, 흰색?)
    ("시행 20분이내 중단 - 자발순환회복", "미시행", True),
    ("미시행 - DOA", "시행 - Sustained ROSC", True),
    ("20분이상 시행", "시행 - Sustained ROSC없이 Any ROSC만", False),
    ("", "", False),
]
PROC_CASES = [("등록병원 시행", True), ("타병원 시행 후 내원", True),
              ("미시행", False), ("미상", False), ("", False)]
EXEC_CASES = [("시행함", True), ("미시행", False), ("미상", False), ("", False)]  # relief_grey(v,"시행함")
# 목표체온조절 하위(저온법4종·일시·온도·속도) 전부: 시행일 때만 흰색 relief_grey(v,"시행")
LAW_TEMP_CASES = [("시행", True), ("미시행", False), ("미상", False), ("", False)]
HYPER_CASES = [("ROSC 24시간 이내", True), ("ROSC 24시간 이후", False), ("미사용", False), ("", False)]

# 소아소생술: 12개월 후 생존=생존/사망일 때만 12개월 후 PCPC 활성, 과거력=기타만 기타칸 활성
def y12_pcpc_grey(live): return live not in ("생존", "사망")
Y12_PCPC_CASES = [("생존", True), ("사망", True), ("미상(연락두절)", False), ("", False)]
YCHILD_ETC_CASES = [("기타", True), ("없음", False), ("신경계질환", False), ("", False)]
# 병원 퇴원시 PCPC=PCPC6(사망)/미상/미입력이면 6·12개월 추적관찰 회색
def y_fu_grey(out): return out in ("PCPC 6 (사망)", "미상", "")
Y_FU_CASES = [("PCPC 1 (뇌기능정상)", True), ("PCPC 5 (코마/식물인간)", True),
              ("PCPC 6 (사망)", False), ("미상", False), ("", False)]

def main():
    for f, g, (w_stop, w_any, w_sus), w_warn in CASES:
        assert (not stop_grey(f, g)) == w_stop, ("중단", f, g)
        assert (not any_grey(f, g)) == w_any,   ("Any", f, g)
        assert (not sus_grey(f, g)) == w_sus,   ("Sus", f, g)
        assert warn(f, g) == w_warn,            ("경고", f, g)
    for q, (w_die, w_batch) in RESULT_CASES:
        assert (not die_grey(q)) == w_die,     ("사망일시", q)
        assert (not batch_grey(q)) == w_batch, ("batch", q)
    for hr, (w_out, w_hdie) in HOSP_CASES:
        assert (not hosp_out_grey(hr)) == w_out,  ("병원퇴원", hr)
        assert (not hosp_die_grey(hr)) == w_hdie, ("병원사망", hr)
    for fs, (w_fdie, w_cpc) in FU6_CASES:
        assert (not fu6_die_grey(fs)) == w_fdie, ("6M사망", fs)
        assert (not f6_cpc_grey(fs)) == w_cpc,   ("6M신경", fs)
    for q, hr, w_cpc in CPC_CASES:
        assert (not hosp_cpc_grey(q, hr)) == w_cpc, ("병원신경", q, hr)
    for dx, tx, (w_tx, w_method) in PREVENT_CASES:
        assert (not tx_grey(dx)) == w_tx,        ("치료여부", dx)
        assert (not method_grey(tx)) == w_method, ("치료방법", tx)
    for w, w_onset in COMMUNITY_CASES:
        assert (not onset_grey(w)) == w_onset, ("발생시각", w)
    for p, w_etc in ETC_CASES:
        assert (not etc_grey(p)) == w_etc, ("기타칸", p)
    for pval, white, w_ok in RELIEF_CASES:
        assert (not relief_grey(pval, white)) == w_ok, ("구급단계", pval, white)
    for gval, drug, w_ok in DOSE_CASES:
        assert (not dose_grey(gval, drug)) == w_ok, ("약물용량", gval, drug)
    for pval, white, w_ok in IN_HOSP_CASES:
        assert (not relief_grey(pval, white)) == w_ok, ("병원단계", pval, white)
    for gv, w_ok in EPINE_CASES:  # 그룹값 있으면 회색
        assert (gv == "") == w_ok, ("에피네프린", gv)
    for v, w_ok in UNIT_CASES:
        assert (not unit_grey(v)) == w_ok, ("단위", v)
    for k, (w_etc, w_unit) in STEROID_CASES:
        assert (not steroid_etc_grey(k)) == w_etc,   ("Steroid종류기타", k)
        assert (not steroid_unit_grey(k)) == w_unit, ("Steroid단위", k)
    for c1, c2, w_ok in ALIVE_GATE_CASES:
        assert (not alive_gate_grey(c1, c2)) == w_ok, ("소생후게이팅", c1, c2)
    for v, w_ok in PROC_CASES:
        assert (not proc_grey(v)) == w_ok, ("시술류", v)
    for v, w_ok in EXEC_CASES:
        assert (not relief_grey(v, "시행함")) == w_ok, ("시행함류", v)
    for v, w_ok in LAW_TEMP_CASES:
        assert (not relief_grey(v, "시행")) == w_ok, ("목표체온", v)
    for v, w_ok in HYPER_CASES:
        assert (not hyper_grey(v)) == w_ok, ("승압제", v)
    for v, w_ok in Y12_PCPC_CASES:
        assert (not y12_pcpc_grey(v)) == w_ok, ("소아12M_PCPC", v)
    for p, w_ok in YCHILD_ETC_CASES:
        assert (not etc_grey(p)) == w_ok, ("소아과거력기타", p)
    for v, w_ok in Y_FU_CASES:
        assert (not y_fu_grey(v)) == w_ok, ("소아추적관찰", v)
    n = (len(CASES) + len(RESULT_CASES) + len(HOSP_CASES) + len(FU6_CASES)
         + len(CPC_CASES) + len(PREVENT_CASES) + len(COMMUNITY_CASES) + len(ETC_CASES)
         + len(RELIEF_CASES) + len(DOSE_CASES) + len(IN_HOSP_CASES) + len(EPINE_CASES)
         + len(UNIT_CASES) + len(STEROID_CASES) + len(ALIVE_GATE_CASES) + len(PROC_CASES)
         + len(EXEC_CASES) + len(LAW_TEMP_CASES) + len(HYPER_CASES)
         + len(Y12_PCPC_CASES) + len(YCHILD_ETC_CASES) + len(Y_FU_CASES))
    print(f"OK: {n} 시나리오 회색/흰색/경고 진리표 통과")

if __name__ == "__main__":
    main()
