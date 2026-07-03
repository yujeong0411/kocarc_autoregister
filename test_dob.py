# -*- coding: utf-8 -*-
"""@DOB 생년월일 통합칸 파싱 검증. 봇이 'YYYY-MM-DD'를 (년,월,일) 셀렉트 value로
정확히 분해하는지 확인(년4자리·월/일 2자리 0채움, 슬래시/점 구분자, 형식오류=None).
실행: uv run python test_dob.py"""
from kocarc_bot import dob_targets

CASES = [
    ("1970-05-15", ("1970", "05", "15")),
    ("1970-5-9",   ("1970", "05", "09")),   # 0채움
    ("2001.12.31", ("2001", "12", "31")),   # 점 구분자
    ("1988/1/1",   ("1988", "01", "01")),   # 슬래시
    ("19700515",   ("1970", "05", "15")),   # 구분자 없이 8자리
    (19700515,     ("1970", "05", "15")),   # 엑셀 숫자값
    ("1970515",    None),                    # 7자리(월/일 모호) → 거부
    ("모름",        None),                    # 형식 안 맞음
    ("",           None),
]


def main():
    for raw, expect in CASES:
        got = dob_targets(raw)
        assert got == expect, (raw, got, expect)
    print(f"OK: {len(CASES)} 생년월일 파싱 통과")


if __name__ == "__main__":
    main()
