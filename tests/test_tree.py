# -*- coding: utf-8 -*-
"""조건부 트리(@TREE) 라운드트립 자체검증 (브라우저 불필요).
build_template 이 만든 잎 라벨을 kocarc_bot 이 다시 (부모값, 자식값)으로
정확히 복원하는지 확인한다. 실행: uv run python tests/test_tree.py"""
import os
import sys
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
from build_template import tree_leaves, load_schema
from kocarc_bot import Resolver, tree_targets, parse_tree_marker


def main():
    schema = load_schema(os.path.join(ROOT, "schema.json"))
    resolver = Resolver(schema)
    checked = 0
    for area in schema["areas"]:
        byname = {f["name"]: f for f in area["fields"]}
        for f in area["fields"]:
            if not f.get("tree"):
                continue
            leaves, marker = tree_leaves(byname, f)
            pname, tree = parse_tree_marker(marker)
            assert pname == f["name"], marker
            assert leaves, f["name"]
            for c in f["choices"]:
                pv, plab = c["value"], (c["label"] or "").strip()
                if not plab:
                    continue
                child = byname.get(tree.get(pv))
                if child:
                    for cc in child["choices"]:
                        clab = (cc["label"] or "").strip()
                        if not clab:
                            continue
                        leaf = f"{plab} - {clab}"
                        tg = tree_targets(resolver, area["key"], marker, leaf)
                        assert tg[0] == (f["name"], pv), (leaf, tg)
                        assert tg[1] == (child["name"], cc["value"]), (leaf, tg)
                        checked += 1
                else:
                    tg = tree_targets(resolver, area["key"], marker, plab)
                    assert tg == [(f["name"], pv)], (plab, tg)
                    checked += 1
    print(f"OK: {checked} 개 잎 라운드트립 통과")


if __name__ == "__main__":
    main()
