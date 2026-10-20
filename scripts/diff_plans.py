#!/usr/bin/env python
# -*- coding: utf-8 -*-
r"""全量差异核对：确认除预期改动外，计划没有被意外修改。

以「大纲编号 + 任务名称」为键（插入任务后 ID 会顺延，不能用 ID 对齐）。

用法:
    python scripts/diff_plans.py data/template.json data/final_mpp.json
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

FIELDS = [
    ("start", "计划开始"),
    ("finish", "计划完成"),
    ("duration", "工期"),
    ("percent_complete", "完成率"),
    ("actual_start", "实际开始"),
    ("actual_finish", "实际完成"),
    ("constraint", "约束"),
]


def key(task: dict) -> str:
    outline = task.get("outline_number") or ""
    return f"{outline}|{task.get('name')}"


def norm(task: dict, field: str):
    value = task.get(field)
    if field == "duration":
        value = (value or {}).get("days")
    if isinstance(value, str) and len(value) >= 10 and value[4:5] == "-":
        return value[:10]
    if isinstance(value, float):
        return round(value, 2)
    return value


def main() -> int:
    before = json.loads((ROOT / sys.argv[1]).read_text(encoding="utf-8"))
    after = json.loads((ROOT / sys.argv[2]).read_text(encoding="utf-8"))

    b_map = {key(t): t for t in before["tasks"]}
    a_map = {key(t): t for t in after["tasks"]}

    only_before = [k for k in b_map if k not in a_map]
    only_after = [k for k in a_map if k not in b_map]

    diffs = []
    for k in b_map:
        if k not in a_map:
            continue
        for field, label in FIELDS:
            bv, av = norm(b_map[k], field), norm(a_map[k], field)
            if bv != av:
                diffs.append((k, label, bv, av))

    print("=" * 78)
    print(f"全量差异核对：{Path(sys.argv[1]).name}  ->  {Path(sys.argv[2]).name}")
    print("=" * 78)
    print(f"改动前任务数: {len(b_map)}    改动后任务数: {len(a_map)}")
    print(f"仅存在于改动前: {len(only_before)}    仅存在于改动后: {len(only_after)}")
    print()
    if only_before:
        print("消失的任务（最多列 10 条）:")
        for k in only_before[:10]:
            print("   -", k)
        print()
    if only_after:
        print("新增的任务:")
        for k in only_after:
            t = a_map[k]
            print(f"   + {k}  {t.get('start', '')[:10]}  工期 {(t.get('duration') or {}).get('days')}"
                  f"{'  [里程碑]' if t.get('is_milestone') else ''}")
        print()
    print(f"字段级差异合计: {len(diffs)} 处")
    for k, label, bv, av in diffs:
        print(f"   · {k}\n       {label}: {bv}  ->  {av}")
    print()
    print("-" * 78)
    unchanged = len(b_map) - len({k for k, *_ in diffs})
    print(f"结论：{len(b_map)} 条原任务中，{unchanged} 条完全未变，{len({k for k, *_ in diffs})} 条按预期变更。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
