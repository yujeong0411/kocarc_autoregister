# -*- coding: utf-8 -*-
"""@DT 통합 일시칸 파싱 검증. 봇이 '2026-06-15 05:02'를 (날짜,시,분)으로
정확히 분해하는지 확인(0채움/날짜만/엑셀 datetime 문자열 대응).
실행: uv run python tests/test_dt.py"""
import os
import sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from kocarc_bot import datetime_targets

CASES = [
    ("2026-06-15 05:02",    ("2026-06-15", "05", "02")),
    ("2026-6-15 5:2",       ("2026-06-15", "05", "02")),   # 0채움
    ("2026-06-15 05:02:00", ("2026-06-15", "05", "02")),   # 엑셀 datetime str
    ("2026/06/15 14:30",    ("2026-06-15", "14", "30")),   # 슬래시
    ("2026-06-15",          ("2026-06-15", None, None)),   # 날짜만
    ("202606150502",        ("2026-06-15", "05", "02")),   # 구분자없이 12자리
    (202606150502,          ("2026-06-15", "05", "02")),   # 엑셀 숫자값
    ("20260615",            ("2026-06-15", None, None)),   # 8자리 날짜만
    ("2026061505",          (None, None, None)),           # 10자리(모호) → 거부
    ("미상",                 (None, None, None)),           # 미상 → 날짜파싱 안됨(봇이 _UK 체크)
    ("",                    (None, None, None)),
]


def main():
    for raw, expect in CASES:
        got = datetime_targets(raw)
        assert got == expect, (raw, got, expect)
    print(f"OK: {len(CASES)} 일시 파싱 통과")


if __name__ == "__main__":
    main()
