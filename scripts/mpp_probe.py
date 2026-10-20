#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""用 MPXJ 读取 Microsoft Project 文件（.mpp/.mpt/.xml），导出结构化 JSON + 控制台摘要。

用法:
    python scripts/mpp_probe.py "<file.mpp>" [--json data/dump.json]

说明:
    - 通过 jpype 直接调用 MPXJ 的 Java API（org.mpxj.reader.UniversalProjectReader）
    - 不修改源文件，只读
    - 依赖项目内 lib/site 下的 jpype 与 mpxj jar（见 scripts/bootstrap_deps.py）
"""
from __future__ import annotations

import argparse
import json
import os
import shutil
import sys
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
# jpype 的原生库在含中文/空格的路径下无法加载，因此依赖目录允许用环境变量指向纯 ASCII 路径
LIB_SITE = Path(os.environ.get("MPXJ_LIB", ROOT / "lib" / "site"))

# ---- 让 jpype 可被找到 ----------------------------------------------------
if str(LIB_SITE) not in sys.path:
    sys.path.insert(0, str(LIB_SITE))

import jpype  # noqa: E402


def find_jars() -> list[str]:
    jars = sorted(str(p) for p in LIB_SITE.rglob("*.jar"))
    if not jars:
        raise SystemExit(f"未找到任何 jar，请按 README 安装 requirements-mpxj.txt 并设置 MPXJ_LIB（查找目录 {LIB_SITE}）")
    return jars


def find_jvm() -> str | None:
    """定位 jvm.dll：JAVA_HOME 优先，其次 PATH 上的 java.exe 反推，最后常见安装目录。"""
    candidates: list[Path] = []
    java_home = os.environ.get("JAVA_HOME")
    if java_home:
        candidates.append(Path(java_home))
    java_exe = shutil.which("java")
    if java_exe:
        # <jdk>/bin/java.exe -> <jdk>
        candidates.append(Path(java_exe).resolve().parent.parent)
    for base in (
        Path(os.environ.get("LOCALAPPDATA", "")) / "Programs" / "Eclipse Adoptium",
        Path("C:/Program Files/Java"),
        Path("C:/Program Files/Eclipse Adoptium"),
    ):
        if base.is_dir():
            candidates.extend(sorted(p for p in base.iterdir() if p.is_dir()))
    for root in candidates:
        for rel in ("bin/server/jvm.dll", "bin/client/jvm.dll", "jre/bin/server/jvm.dll"):
            dll = root / rel
            if dll.exists():
                return str(dll)
    return None


def start_jvm() -> None:
    if jpype.isJVMStarted():
        return
    jars = find_jars()
    jvm = find_jvm()
    if not jvm:
        raise SystemExit("未找到 jvm.dll，请设置 JAVA_HOME 指向 JDK 目录")
    print(f"[jvm] jvm.dll: {jvm}")
    print(f"[jvm] classpath 中 jar 数量: {len(jars)}")
    jpype.startJVM(jvmpath=jvm, classpath=jars, convertStrings=True)


# ---- 取值辅助（字段缺失不应中断整份导出）---------------------------------
def safe(obj, method: str, default=None):
    try:
        fn = getattr(obj, method)
        return fn()
    except Exception:
        return default


def as_date(value) -> str | None:
    if value is None:
        return None
    try:
        ms = int(value.getTime())
        return datetime.fromtimestamp(ms / 1000).strftime("%Y-%m-%d")
    except Exception:
        return str(value)


def as_number(value):
    if value is None:
        return None
    try:
        f = float(value)
        return int(f) if f == int(f) else round(f, 2)
    except Exception:
        try:
            return float(value.doubleValue())
        except Exception:
            return None


def as_duration(value):
    """MPXJ Duration -> {'days': x, 'hours': y, 'text': '...'}"""
    if value is None:
        return None
    out = {}
    out["value"] = as_number(safe(value, "getDuration"))
    out["units"] = str(safe(value, "getUnits", ""))
    days = safe(value, "getDuration")
    if days is not None:
        try:
            out["days"] = round(float(days), 2)
        except Exception:
            pass
    hours = safe(value, "getDuration", None)
    try:
        out["hours"] = round(float(value.getDuration(2)), 2)  # TimeUnit.HOURS == 2
    except Exception:
        pass
    txt = safe(value, "toString")
    if txt:
        out["text"] = str(txt)
    return out or None


def as_link(value):
    if value is None:
        return None
    return {
        "id": as_number(safe(value, "getID")),
        "name": str(safe(value, "getName", "") or ""),
    }


def read_project(path: Path) -> dict:
    UniversalProjectReader = jpype.JClass("org.mpxj.reader.UniversalProjectReader")
    project = UniversalProjectReader().read(str(path))

    props = safe(project, "getProjectProperties")
    default_cal = safe(props, "getDefaultCalendar") if props else None

    data: dict = {
        "source_file": str(path),
        "file_size_kb": round(path.stat().st_size / 1024, 1),
        "read_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "project": {
            "title": str(safe(props, "getProjectTitle", "") or ""),
            "subject": str(safe(props, "getSubject", "") or ""),
            "manager": str(safe(props, "getManager", "") or ""),
            "company": str(safe(props, "getCompany", "") or ""),
            "author": str(safe(props, "getAuthor", "") or ""),
            "start_date": as_date(safe(props, "getStartDate")),
            "finish_date": as_date(safe(props, "getFinishDate")),
            "schedule_from": str(safe(props, "getScheduleFrom", "") or ""),
            "default_calendar": str(safe(default_cal, "getName", "") or "") if default_cal else None,
            "last_saved": as_date(safe(props, "getLastSaved")),
        },
        "calendars": [],
        "resources": [],
        "tasks": [],
        "relations": [],
        "warnings": [],
    }

    # 日历
    cals = safe(project, "getCalendars")
    if cals:
        for cal in cals:
            entry = {
                "id": as_number(safe(cal, "getUniqueID")),
                "name": str(safe(cal, "getName", "") or ""),
                "minutes_per_day": as_number(safe(cal, "getMinutesPerDay")),
                "minutes_per_week": as_number(safe(cal, "getMinutesPerWeek")),
                "days_per_week": as_number(safe(cal, "getDaysPerWeek")),
            }
            # 一周内哪些天为工作日
            try:
                working = []
                Day = jpype.JClass("java.time.DayOfWeek")
                for d in Day.values():
                    wd = safe(cal, "isWorkingDay", None)
                    if wd is not None:
                        working.append(str(d))
                entry["note"] = "工作日判定见 project 默认日历设置"
            except Exception as exc:
                data["warnings"].append(f"calendar detail: {exc}")
            data["calendars"].append(entry)

    # 资源
    resources = safe(project, "getResources")
    if resources:
        for res in resources:
            rid = as_number(safe(res, "getUniqueID"))
            if rid == 0:
                continue
            data["resources"].append({
                "id": rid,
                "name": str(safe(res, "getName", "") or ""),
                "type": str(safe(res, "getType", "") or ""),
                "max_units": as_number(safe(res, "getMaxUnits")),
            })

    # 任务
    tasks = safe(project, "getTasks")
    by_id: dict[int, dict] = {}
    if tasks:
        for t in tasks:
            tid = as_number(safe(t, "getID"))
            parent = safe(t, "getParentTask")
            predecessors = []
            rel_list = safe(t, "getPredecessors")
            if rel_list:
                for rel in rel_list:
                    pred = safe(rel, "getPredecessorTask")
                    predecessors.append({
                        "predecessor_id": as_number(safe(pred, "getID")),
                        "predecessor_name": str(safe(pred, "getName", "") or ""),
                        "type": str(safe(rel, "getType", "") or ""),
                        "lag": as_duration(safe(rel, "getLag")),
                    })
            assignees = []
            assignment_details = []
            assigns = safe(t, "getResourceAssignments")
            if assigns:
                for a in assigns:
                    r = safe(a, "getResource")
                    if r is not None:
                        assignees.append(str(safe(r, "getName", "") or ""))
                        assignment_details.append({
                            "resource_name": str(safe(r, "getName", "") or ""),
                            "units_percent": as_number(safe(a, "getUnits")),
                        })
            entry = {
                "id": tid,
                "unique_id": as_number(safe(t, "getUniqueID")),
                "wbs": str(safe(t, "getWBS", "") or ""),
                "outline_number": str(safe(t, "getOutlineNumber", "") or ""),
                "outline_level": as_number(safe(t, "getOutlineLevel")),
                "name": str(safe(t, "getName", "") or ""),
                "parent_id": as_number(safe(parent, "getID")) if parent is not None else None,
                "parent_name": str(safe(parent, "getName", "") or "") if parent is not None else None,
                "is_summary": bool(safe(t, "getSummary", False)),
                "is_milestone": bool(safe(t, "getMilestone", False)),
                "start": as_date(safe(t, "getStart")),
                "finish": as_date(safe(t, "getFinish")),
                "actual_start": as_date(safe(t, "getActualStart")),
                "actual_finish": as_date(safe(t, "getActualFinish")),
                "duration": as_duration(safe(t, "getDuration")),
                "remaining_duration": as_duration(safe(t, "getRemainingDuration")),
                "percent_complete": as_number(safe(t, "getPercentageComplete")),
                "percent_work_complete": as_number(safe(t, "getPercentageWorkComplete")),
                "critical": bool(safe(t, "getCritical", False)),
                "total_slack": as_duration(safe(t, "getTotalSlack")),
                "free_slack": as_duration(safe(t, "getFreeSlack")),
                "deadline": as_date(safe(t, "getDeadline")),
                "constraint": str(safe(t, "getConstraintType", "") or ""),
                "constraint_date": as_date(safe(t, "getConstraintDate")),
                "priority": str(safe(t, "getPriority", "") or ""),
                "notes": str(safe(t, "getNotes", "") or "")[:500],
                "assignees": assignees,
                "assignments": assignment_details,
                "predecessors": predecessors,
                "predecessor_count": len(predecessors),
            }
            if tid is not None:
                by_id[tid] = entry
            data["tasks"].append(entry)

    # 任务间关系（从 task.getSuccessors 也统计一遍，便于核对）
    project_relations = []
    proj_rel = safe(project, "getRelations")
    if proj_rel:
        for rel in proj_rel:
            pred = safe(rel, "getPredecessorTask")
            succ = safe(rel, "getSuccessorTask")
            project_relations.append({
                "from_id": as_number(safe(pred, "getID")) if pred is not None else None,
                "from_name": str(safe(pred, "getName", "") or "") if pred is not None else None,
                "to_id": as_number(safe(succ, "getID")) if succ is not None else None,
                "to_name": str(safe(succ, "getName", "") or "") if succ is not None else None,
                "type": str(safe(rel, "getType", "") or ""),
                "lag": as_duration(safe(rel, "getLag")),
            })

    # 里程碑兜底判定：未设里程碑标记但工期为 0 的非摘要任务
    for t in data["tasks"]:
        if not t["is_milestone"] and not t["is_summary"]:
            dur = (t["duration"] or {}).get("days")
            if dur == 0:
                t["zero_duration"] = True
        else:
            t["zero_duration"] = False

    for t in data["tasks"]:
        for pred in t["predecessors"]:
            data["relations"].append({
                "from_id": pred["predecessor_id"],
                "from_name": pred["predecessor_name"],
                "to_id": t["id"],
                "to_name": t["name"],
                "type": pred["type"],
                "lag": pred["lag"],
            })

    data["project_relations"] = project_relations
    data["stats"] = {
        "task_count": len(data["tasks"]),
        "summary_count": sum(1 for t in data["tasks"] if t["is_summary"]),
        "leaf_count": sum(1 for t in data["tasks"] if not t["is_summary"]),
        "milestone_count": sum(1 for t in data["tasks"] if t["is_milestone"]),
        "zero_duration_count": sum(1 for t in data["tasks"] if t.get("zero_duration")),
        "relation_count": len(data["relations"]),
        "project_relation_count": len(project_relations),
        "resource_count": len(data["resources"]),
        "calendar_count": len(data["calendars"]),
        "critical_count": sum(1 for t in data["tasks"] if t["critical"]),
        "max_outline_level": max((t["outline_level"] or 0) for t in data["tasks"]) if data["tasks"] else 0,
        "top_level_count": sum(1 for t in data["tasks"] if (t["outline_level"] or 0) == 1),
        "company": data["project"]["company"],
    }
    return data


def print_summary(data: dict) -> None:
    p = data["project"]
    s = data["stats"]
    print("=" * 78)
    print("项目文件摘要")
    print("=" * 78)
    print(f"文件      : {data['source_file']}")
    print(f"大小      : {data['file_size_kb']} KB")
    print(f"项目名称  : {p['title']}")
    print(f"公司/作者 : {p['company']} / {p['manager']}")
    print(f"计划期间  : {p['start_date']} ~ {p['finish_date']}")
    print(f"排程方式  : {p['schedule_from']}   默认日历: {p['default_calendar']}")
    print()
    print(f"任务总数  : {s['task_count']}（摘要 {s['summary_count']} / 明细 {s['leaf_count']}）")
    print(f"里程碑    : {s['milestone_count']}")
    print(f"依赖关系  : {s['relation_count']}")
    print(f"关键任务  : {s['critical_count']}")
    print(f"资源      : {s['resource_count']}    日历: {s['calendar_count']}")
    print(f"层级深度  : {s['max_outline_level']}    一级 WBS 数: {s['top_level_count']}")
    print()
    print("--- 日历 ---")
    for c in data["calendars"]:
        print(f"  {c['name']}: {c['minutes_per_day']} 分钟/天, {c['minutes_per_week']} 分钟/周")
    print()
    print("--- 一级 WBS ---")
    for t in data["tasks"]:
        if (t["outline_level"] or 0) == 1:
            kids = sum(1 for x in data["tasks"] if x["parent_id"] == t["id"])
            flag = " [里程碑]" if t["is_milestone"] else ""
            print(f"  {t['id']:>4} {t['name'][:44]:<46} 子任务 {kids:>3}  "
                  f"{t['start']}~{t['finish']}  {t['percent_complete']}%{flag}")
    print()
    print("--- 里程碑清单 ---")
    for t in data["tasks"]:
        if t["is_milestone"]:
            print(f"  {t['id']:>4} L{t['outline_level']} {t['name'][:40]:<42} {t['start']}  {t['percent_complete']}%")
    print()
    print("--- 有依赖关系的任务（示例前 25 条）---")
    shown = 0
    for t in data["tasks"]:
        if t["predecessors"] and shown < 25:
            preds = ", ".join(
                f"#{pr['predecessor_id']}({pr['type']}{'/lag=' + str(pr['lag']['text']) if pr['lag'] else ''})"
                for pr in t["predecessors"]
            )
            print(f"  {t['id']:>4} {t['name'][:34]:<36} <- {preds}")
            shown += 1
    if data["warnings"]:
        print()
        print("--- 警告 ---")
        for w in data["warnings"][:10]:
            print("  " + w)


def write_markdown(data: dict, out: Path) -> None:
    """输出 UTF-8 可读大纲，便于人工核对解析结果是否忠实于原文件。"""
    p, s = data["project"], data["stats"]
    lines: list[str] = []
    lines.append(f"# {p['title'] or Path(data['source_file']).stem}")
    lines.append("")
    lines.append("## 一、文件与项目信息")
    lines.append("")
    lines.append(f"- 源文件：`{data['source_file']}`")
    lines.append(f"- 文件大小：{data['file_size_kb']} KB")
    lines.append(f"- 项目标题：{p['title']}")
    lines.append(f"- 公司 / 负责人：{p['company']} / {p['manager'] or '（未设置）'}")
    lines.append(f"- 计划期间：{p['start_date']} ~ {p['finish_date']}")
    lines.append(f"- 默认日历：{p['default_calendar']}")
    lines.append(f"- 最后保存：{p['last_saved']}")
    lines.append("")
    lines.append("## 二、规模统计")
    lines.append("")
    lines.append("| 指标 | 数值 |")
    lines.append("|---|---|")
    lines.append(f"| 任务总数 | {s['task_count']} |")
    lines.append(f"| 摘要（父）任务 | {s['summary_count']} |")
    lines.append(f"| 明细（叶）任务 | {s['leaf_count']} |")
    lines.append(f"| 里程碑（已标记） | {s['milestone_count']} |")
    lines.append(f"| 零工期任务 | {s.get('zero_duration_count', 0)} |")
    lines.append(f"| 依赖关系（任务侧） | {s['relation_count']} |")
    lines.append(f"| 依赖关系（文件级） | {s.get('project_relation_count', 0)} |")
    lines.append(f"| 资源 | {s['resource_count']} |")
    lines.append(f"| 日历 | {s['calendar_count']} |")
    lines.append(f"| 标记为关键的任务 | {s['critical_count']} |")
    lines.append(f"| 最大层级深度 | {s['max_outline_level']} |")
    lines.append(f"| 一级 WBS 数量 | {s['top_level_count']} |")
    lines.append("")
    lines.append("## 三、日历")
    lines.append("")
    lines.append("| 名称 | 分钟/天 | 分钟/周 | 折算 |")
    lines.append("|---|---|---|---|")
    for c in data["calendars"]:
        mpw = c["minutes_per_week"] or 0
        days = round(mpw / 480, 2) if mpw else None
        lines.append(f"| {c['name']} | {c['minutes_per_day']} | {mpw} | 约 {days} 天/周（按 8h/天） |")
    lines.append("")
    lines.append("## 四、资源")
    lines.append("")
    lines.append("| ID | 名称 | 类型 |")
    lines.append("|---|---|---|")
    for r in data["resources"]:
        lines.append(f"| {r['id']} | {r['name']} | {r['type']} |")
    lines.append("")
    lines.append("## 五、完整任务大纲（层级 / 日期 / 工期 / 完成率 / 前置）")
    lines.append("")
    lines.append("| ID | 层级 | 任务名称 | 开始 | 完成 | 工期(天) | 完成率 | 负责人 | 前置 | 属性 |")
    lines.append("|---|---|---|---|---|---|---|---|---|---|")
    for t in data["tasks"]:
        level = t["outline_level"] or 0
        indent = "&nbsp;" * 4 * max(level - 1, 0)
        name = (t["name"] or "").replace("|", "\\|")
        dur = (t["duration"] or {}).get("days")
        preds = ", ".join(
            f"{pr['predecessor_id']}{pr['type']}{'' if not pr['lag'] else ' lag=' + str(pr['lag'].get('text'))}"
            for pr in t["predecessors"]
        ) or ""
        flags = []
        if t["is_summary"]:
            flags.append("摘要")
        if t["is_milestone"]:
            flags.append("里程碑")
        if t.get("zero_duration"):
            flags.append("零工期")
        if t["critical"]:
            flags.append("关键")
        if t["constraint"]:
            flags.append(f"约束:{t['constraint']}")
        owner = ", ".join(t["assignees"]) or ""
        lines.append(
            f"| {t['id']} | {level} | {indent}{name} | {(t['start'] or '')[:10]} | "
            f"{(t['finish'] or '')[:10]} | {dur if dur is not None else ''} | "
            f"{t['percent_complete'] if t['percent_complete'] is not None else ''} | {owner} | {preds} | {','.join(flags)} |"
        )
    lines.append("")
    if data.get("project_relations"):
        lines.append("## 六、依赖关系明细")
        lines.append("")
        lines.append("| 前置任务 | 类型 | 后续任务 | 滞后 |")
        lines.append("|---|---|---|---|")
        for r in data["project_relations"]:
            lines.append(f"| {r['from_id']} {r['from_name']} | {r['type']} | {r['to_id']} {r['to_name']} | "
                         f"{(r['lag'] or {}).get('text', '')} |")
        lines.append("")
    if data.get("warnings"):
        lines.append("## 七、解析警告")
        lines.append("")
        for w in data["warnings"]:
            lines.append(f"- {w}")
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text("\n".join(lines), encoding="utf-8")


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("mpp_file")
    ap.add_argument("--json", default=None, help="JSON 输出路径（默认 data/<文件名>.json）")
    ap.add_argument("--markdown", default=None, help="UTF-8 大纲输出路径（默认 docs/<文件名>_大纲.md）")
    ap.add_argument("--no-markdown", action="store_true", help="仅输出 JSON，不生成默认大纲")
    args = ap.parse_args()

    path = Path(args.mpp_file)
    if not path.exists():
        raise SystemExit(f"文件不存在: {path}")

    start_jvm()
    data = read_project(path)
    print_summary(data)

    out = Path(args.json) if args.json else ROOT / "data" / (path.stem + ".json")
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"结构化数据已写入: {out}")

    if not args.no_markdown:
        md = Path(args.markdown) if args.markdown else ROOT / "docs" / (path.stem + "_大纲.md")
        write_markdown(data, md)
        print(f"可读大纲已写入: {md}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
