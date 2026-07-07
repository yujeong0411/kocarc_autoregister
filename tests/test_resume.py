# -*- coding: utf-8 -*-
"""resume_decision 검증. 중복생성 방지가 핵심:
이미 생성된(=pat_id 기록된) 환자는 절대 'create' 가 나오면 안 된다.
실행: uv run python tests/test_resume.py"""
import os
import sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from kocarc_bot import resume_decision, all_done


def row(status, pat_id=""):
    return {"status": status, "pat_id": pat_id}


CASES = [
    # (이전행, 기대 action, 기대 pat_id)
    (None,                              "create", None),   # 첫 실행
    (row("create_fail"),               "create", None),   # 검증에 막혀 생성 안 됨 → 재시도 안전
    (row("done", "14-1"),              "done",   "14-1"), # 완료 → 건너뜀
    (row("partial:heart", "14-2"),     "reuse",  "14-2"), # 일부 실패 → 재생성 금지, 영역 재저장
    (row("error:foo", "14-3"),         "reuse",  "14-3"), # 오류지만 생성됨 → 재사용
    (row("no_pat_id"),                 "manual", None),   # 저장됐을 수 있으나 id 불명 → 수동
    (row("error:foo", ""),             "manual", None),   # 생성 여부 불명 → 수동(재생성 금지)
]


def demo():
    for prev, exp_action, exp_pat in CASES:
        action, pat_id = resume_decision(prev)
        assert action == exp_action, f"{prev} → action {action} != {exp_action}"
        assert pat_id == exp_pat, f"{prev} → pat_id {pat_id!r} != {exp_pat!r}"
    # 안전 불변식: pat_id 가 기록돼 있으면 절대 create 가 아니어야 한다.
    for status in ("partial:x", "error:y", "done"):
        act, _ = resume_decision(row(status, "14-9"))
        assert act != "create", f"중복생성 위험: {status} 에서 create 반환"

    # all_done: 엑셀 전 환자가 done 일 때만 True(→ progress.csv 자동삭제 트리거).
    patients = {"1": {}, "2": {}, "3": {}}      # 엑셀에 환자 3명
    prog_all = {"1": row("done","a"), "2": row("done","b"), "3": row("done","c")}
    prog_part = {"1": row("done","a"), "2": row("partial:heart","b")}  # 3 미기록
    assert all_done(patients, prog_all) is True
    assert all_done(patients, prog_part) is False        # 2 partial + 3 없음
    assert all_done(patients, {}) is False               # 아무것도 안 됨
    assert all_done({}, prog_all) is False               # 환자 없음 → 삭제 안 함
    assert all_done(patients, {"1":row("done"),"2":row("done"),"3":row("done")}) is True
    print(f"OK: {len(CASES)} cases + 중복생성 불변식 + all_done 통과")


if __name__ == "__main__":
    demo()
