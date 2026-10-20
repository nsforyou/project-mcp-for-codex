"""Safe construction of a new native MPP from MSPDI Project XML.

This module is intentionally only for creating a *new* plan.  Existing MPP
files must keep using the Change Set / COM in-place workflow, which preserves
their Project view layer.  A fresh plan has no view layer to preserve.
"""
from __future__ import annotations

import gc
import re
import json
import argparse
import sys
import xml.etree.ElementTree as ET
from datetime import datetime
from pathlib import Path

import pythoncom
import win32com.client


NS = {"p": "http://schemas.microsoft.com/project"}


def _text(node, tag: str, default: str = "") -> str:
    value = node.find(f"p:{tag}", NS)
    return value.text if value is not None and value.text else default


def _when(value: str):
    # Passing a naive Python datetime to COM shifts it by the host timezone.
    # Project parses this local, unambiguous timestamp without UTC conversion.
    return datetime.fromisoformat(value).strftime("%Y-%m-%d %H:%M") if value else None


def _minutes(duration):
    if not duration:
        return 0.0
    factors = {"m": 1, "h": 60, "d": 480, "w": 2400}
    unit = str(duration.get("units", "")).lower()
    if unit not in factors:
        raise ValueError(f"不能核验的工期单位: {unit}")
    return float(duration["value"]) * factors[unit]


def _iso_minutes(value):
    match = re.fullmatch(r"PT(?:(\d+(?:\.\d+)?)H)?(?:(\d+(?:\.\d+)?)M)?(?:(\d+(?:\.\d+)?)S)?", value)
    if not match:
        raise ValueError(f"不支持的 ISO 工期: {value}")
    return float(match[1] or 0) * 60 + float(match[2] or 0) + float(match[3] or 0) / 60


def _links(node):
    """Accept MSPDI plus the legacy PredecessorTasks container used by drafts."""
    found = {}
    for link in node.findall("p:PredecessorLink", NS):
        found[int(_text(link, "PredecessorUID"))] = (
            int(_text(link, "Type", "1")), int(_text(link, "LinkLag", "0")))
    for link in node.findall("p:PredecessorTasks/p:Predecessor", NS):
        uid = int(_text(link, "UID"))
        value = (int(_text(link, "Type", "1")), int(_text(link, "LinkLag", "0")))
        if uid in found and found[uid] != value:
            raise ValueError(f"冲突的前置关系 UID={uid}")
        found[uid] = value
    return found


def _nodes(root):
    tasks = [n for n in root.findall("p:Tasks/p:Task", NS) if _text(n, "UID") != "0"]
    resources = [n for n in root.findall("p:Resources/p:Resource", NS) if _text(n, "UID") != "0"]
    return tasks, resources, root.findall("p:Assignments/p:Assignment", NS)


def inspect_xml(xml_path: str | Path) -> dict:
    path = Path(xml_path).resolve()
    root = ET.parse(path).getroot()
    tasks, resources, assignments = _nodes(root)
    relations = sum(len(_links(t)) for t in tasks)
    milestones = sum(_text(t, "Milestone") == "1" for t in tasks)
    if not tasks:
        raise ValueError("XML 中没有任务，拒绝创建空计划")
    if root.tag != "{http://schemas.microsoft.com/project}Project":
        raise ValueError("不是 Microsoft Project MSPDI XML")
    if root.find("p:Calendars", NS) is not None:
        raise ValueError("当前创建器仅支持默认标准日历；自定义日历尚未实现，拒绝静默丢失")
    uids = [int(_text(t, "UID")) for t in tasks]
    rids = [int(_text(r, "UID")) for r in resources]
    if len(set(uids)) != len(uids) or len(set(rids)) != len(rids):
        raise ValueError("任务或资源 UID 重复")
    if len({_text(r, 'Name') for r in resources}) != len(resources):
        raise ValueError("资源名称重复，无法可靠分配")
    previous_level = 0
    graph = {}
    for task in tasks:
        uid = int(_text(task, "UID"))
        level = int(_text(task, "OutlineLevel", "1"))
        if level < 1 or level > previous_level + 1:
            raise ValueError(f"非法 WBS 层级 UID={uid}")
        previous_level = level
        for field in ("ConstraintType", "CalendarUID", "PercentComplete"):
            if _text(task, field, "0") not in ("0", "-1"):
                raise ValueError(f"当前创建器尚不支持 {field}，UID={uid}")
        for field in ("ActualStart", "ActualFinish", "ActualWork", "Deadline"):
            if _text(task, field):
                raise ValueError(f"当前创建器尚不支持 {field}，UID={uid}")
        if _text(task, "Duration"):
            _iso_minutes(_text(task, "Duration"))
        if task.find("p:Baseline", NS) is not None:
            raise ValueError("当前创建器不导入基线")
        graph[uid] = list(_links(task))
        for pred, (kind, lag) in _links(task).items():
            if pred not in uids or kind not in range(4) or pred == uid:
                raise ValueError(f"非法前置关系 UID={uid}, predecessor={pred}")
    visiting, visited = set(), set()
    def visit(uid):
        if uid in visiting:
            raise ValueError("任务依赖成环")
        if uid in visited:
            return
        visiting.add(uid)
        for pred in graph[uid]:
            visit(pred)
        visiting.remove(uid)
        visited.add(uid)
    for uid in uids:
        visit(uid)
    pairs = set()
    for assignment in assignments:
        pair = (int(_text(assignment, "TaskUID")), int(_text(assignment, "ResourceUID")))
        if pair[0] not in uids or pair[1] not in rids or pair in pairs:
            raise ValueError("无效或重复的资源分配")
        pairs.add(pair)
        if float(_text(assignment, "Units", "1")) <= 0:
            raise ValueError("资源 Units 必须为正数")
    return {
        "xml": str(path), "tasks": len(tasks), "resources": len(resources),
        "assignments": len(assignments), "relations": relations,
        "milestones": milestones,
        "warnings": ["新建计划使用标准日历、手动排程；不导入原 MPP 视图/成本/实际进度。",
                     "资源容量优先使用 XML MaxUnits；缺失时采用 1，绝不从名称猜测人数。"],
    }


def _build_new_mpp_worker(xml_path: str | Path, output_path: str | Path) -> dict:
    """Create a native MPP without ever opening an existing user MPP.

    The parent holds the COM lock and rejects existing Project processes.
    DispatchEx alone is not an isolation guarantee for a single-instance app.
    The destination must not exist.
    """
    source = Path(xml_path).resolve()
    output = Path(output_path).resolve()
    if source.suffix.lower() != ".xml" or not source.is_file():
        raise ValueError(f"需要有效的 Project XML 文件: {source}")
    if output.suffix.lower() != ".mpp":
        raise ValueError("输出文件必须以 .mpp 结尾")
    if output.exists():
        raise FileExistsError(f"拒绝覆盖已有 MPP: {output}")
    output.parent.mkdir(parents=True, exist_ok=True)

    expected = inspect_xml(source)
    root = ET.parse(source).getroot()
    task_nodes, resource_nodes, assignment_nodes = _nodes(root)
    if not task_nodes:
        raise ValueError("XML 中没有任务，拒绝创建空计划")

    app = project = None
    pythoncom.CoInitialize()
    try:
        app = win32com.client.DispatchEx("MSProject.Application")
        app.Visible = False
        app.DisplayAlerts = False
        try:
            app.AskToUpdateLinks = False
        except Exception:
            pass
        app.FileNew()
        project = app.ActiveProject
        if project.Tasks.Count or project.Resources.Count:
            raise RuntimeError("隔离的 Project 实例不是空白计划；已中止")
        start = _when(_text(root, "StartDate"))
        if start:
            project.ProjectStart = start

        task_by_uid: dict[int, object] = {}
        rows: list[tuple[int, int]] = []
        for node in task_nodes:
            uid = int(_text(node, "UID"))
            task = project.Tasks.Add(_text(node, "Name", f"任务{uid}"))
            task.Manual = True
            if _text(node, "Notes"):
                task.Notes = _text(node, "Notes")
            if _text(node, "Duration") and _text(node, "Summary") != "1":
                task.Duration = _iso_minutes(_text(node, "Duration"))
            task_start, task_finish = _when(_text(node, "Start")), _when(_text(node, "Finish"))
            if task_start:
                task.Start = task_start
            if task_finish:
                task.Finish = task_finish
            if _text(node, "Milestone") == "1":
                task.Milestone = True
            task_by_uid[uid] = task
            rows.append((task.ID, int(_text(node, "OutlineLevel", "1"))))

        # Restore WBS in contiguous batches.  Project's OutlineIndent is slow
        # per-row but reliable for a selected range.
        for level in range(1, max(level for _, level in rows)):
            for index, (_, row_level) in enumerate(rows):
                if row_level != level or index + 1 >= len(rows) or rows[index + 1][1] <= level:
                    continue
                last = index + 1
                while last + 1 < len(rows) and rows[last + 1][1] > level:
                    last += 1
                app.SelectRow(rows[index + 1][0], False, last - index)
                app.OutlineIndent()

        resource_by_uid: dict[int, object] = {}
        for node in resource_nodes:
            uid = int(_text(node, "UID"))
            resource = project.Resources.Add(_text(node, "Name", f"资源{uid}"))
            resource.MaxUnits = float(_text(node, "MaxUnits", "1"))
            resource_by_uid[uid] = resource

        # Assignments.Add is unreliable in Project 2019 COM here.  Assigning
        # ResourceNames is native Project behaviour and creates assignments
        # consistently; then set the requested crew units.
        by_task: dict[int, list[tuple[int, float]]] = {}
        for node in assignment_nodes:
            task_uid, resource_uid = int(_text(node, "TaskUID")), int(_text(node, "ResourceUID"))
            if task_uid not in task_by_uid or resource_uid not in resource_by_uid:
                raise ValueError("XML assignment references an unknown task or resource")
            by_task.setdefault(task_uid, []).append((resource_uid, float(_text(node, "Units", "1"))))
        for task_uid, requested in by_task.items():
            task = task_by_uid[task_uid]
            task.ResourceNames = ", ".join(resource_by_uid[rid].Name for rid, _ in requested)
            for assignment in task.Assignments:
                for resource_uid, units in requested:
                    if assignment.ResourceName == resource_by_uid[resource_uid].Name:
                        assignment.Units = units
                        break

        for node in task_nodes:
            uid = int(_text(node, "UID"))
            predecessors = [f"{task_by_uid[pred].ID}{['FF','FS','SF','SS'][kind]}{lag / 10:+g}m"
                            for pred, (kind, lag) in _links(node).items()]
            if predecessors:
                task_by_uid[uid].Predecessors = ",".join(predecessors)

        # Resource assignments and dependencies may recalculate task dates.
        # Restore the draft's explicit leaf dates last, including 0-duration
        # milestones, then verify the saved file independently.
        for node in task_nodes:
            task = task_by_uid[int(_text(node, "UID"))]
            if task.Summary:
                continue
            if _text(node, "Start"):
                task.Start = _when(_text(node, "Start"))
            if _text(node, "Finish"):
                task.Finish = _when(_text(node, "Finish"))
            if _text(node, "Milestone") == "1":
                task.Duration = 0
                task.Milestone = True
        if project.Tasks.Count != len(task_nodes) or project.Resources.Count != len(resource_nodes):
            raise RuntimeError("写入前的任务/资源数量不一致，拒绝保存")
        app.FileSaveAs(str(output), 0)  # pjMPP; only valid for a fresh plan
        return {"output": str(output), "tasks": project.Tasks.Count,
                "resources": project.Resources.Count, "assignments": len(assignment_nodes),
                "relations": expected["relations"]}
    finally:
        if app is not None:
            try:
                app.FileCloseEx(0)
            except Exception:
                pass
            app.Quit(0)
        project = app = None
        gc.collect()
        pythoncom.CoUninitialize()


def verify_new_mpp(source, output):
    from com_bridge import mpxj_read
    data = mpxj_read(str(output))
    tasks = [t for t in data["tasks"] if t.get("id") != 0]
    nodes, resources, assignments = _nodes(ET.parse(source).getroot())
    checks = []
    checks.append({"field": "task_count", "ok": len(tasks) == len(nodes)})
    checks.append({"field": "resource_count", "ok": len(data["resources"]) == len(resources)})
    actual_resources = {r['name']: r for r in data['resources']}
    for resource in resources:
        name = _text(resource, 'Name')
        expected_capacity = float(_text(resource, 'MaxUnits', '1')) * 100
        actual_capacity = actual_resources.get(name, {}).get('max_units')
        checks.append({'field': 'MaxUnits', 'resource': name, 'expected': expected_capacity,
                       'actual': actual_capacity, 'ok': expected_capacity == actual_capacity})
    for node, task in zip(nodes, tasks):
        for field, actual in (("Name", task["name"]), ("OutlineLevel", str(task["outline_level"]))):
            checks.append({"task": task["id"], "field": field, "ok": _text(node, field) == actual})
        if not task.get("is_summary"):
            for field in ("Start", "Finish"):
                expected = _text(node, field)
                if expected:
                    actual = task.get(field.lower()) or ""
                    checks.append({"task": task["id"], "field": field,
                                   "expected": expected[:16], "actual": actual[:16],
                                   "ok": expected[:16] == actual[:16]})
        checks.append({"task": task["id"], "field": "Milestone",
                       "ok": bool(task.get("is_milestone")) == (_text(node, "Milestone") == "1")})
        uid_to_row = {int(_text(n, "UID")): i + 1 for i, n in enumerate(nodes)}
        expected_links = sorted((uid_to_row[uid], ['FF','FS','SF','SS'][kind], lag / 10)
                                for uid, (kind, lag) in _links(node).items())
        actual_links = sorted((p["predecessor_id"], p["type"], _minutes(p.get("lag")))
                              for p in task.get("predecessors", []))
        checks.append({"task": task["id"], "field": "Dependencies", "expected": expected_links,
                       "actual": actual_links, "ok": expected_links == actual_links})
        resource_names = {int(_text(r, "UID")): _text(r, "Name") for r in resources}
        requested = sorted((resource_names[int(_text(a, "ResourceUID"))], float(_text(a, "Units", "1")) * 100)
                           for a in assignments if _text(a, "TaskUID") == _text(node, "UID"))
        actual_assignments = sorted((a["resource_name"], a["units_percent"]) for a in task.get("assignments", []))
        checks.append({"task": task["id"], "field": "Assignments", "expected": requested,
                       "actual": actual_assignments, "ok": requested == actual_assignments})
        if _text(node, "Duration") and not task.get("is_summary"):
            expected_duration = _iso_minutes(_text(node, "Duration"))
            actual_duration = _minutes(task.get("duration"))
            checks.append({"task": task["id"], "field": "Duration", "expected_minutes": expected_duration,
                           "actual_minutes": actual_duration, "ok": abs(expected_duration - actual_duration) < 0.01})
    checks.append({"field": "relation_count", "ok": len(data["relations"]) == inspect_xml(source)["relations"]})
    return {"ok": all(c["ok"] for c in checks), "checks": checks}


def main():
    import os
    from com_runtime import exclusive_project
    parser = argparse.ArgumentParser()
    parser.add_argument("source")
    parser.add_argument("output")
    args = parser.parse_args()
    # Normal invocation takes the lock itself; bridge workers inherit their
    # parent's locked session and must not recursively acquire it.
    if os.environ.get('PROJECT_COM_PARENT') == str(os.getppid()):
        result = _build_new_mpp_worker(args.source, args.output)
    else:
        with exclusive_project():
            result = _build_new_mpp_worker(args.source, args.output)
    result["verification"] = verify_new_mpp(args.source, args.output)
    print(json.dumps(result, ensure_ascii=False))
    return 0 if result["verification"]["ok"] else 2


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
