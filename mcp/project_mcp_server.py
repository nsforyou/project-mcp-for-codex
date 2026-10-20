#!/usr/bin/env python
# -*- coding: utf-8 -*-
r"""Microsoft Project MCP Server（stdio）。

工具分层：
    READ   —— 项目/任务/依赖/资源/日历/视图层/基线（MPXJ 为主，视图层走 COM）
    WRITE  —— 一律先 Preview 成 Change Set，经用户批准后由 Project COM 原生写入
    流程   —— READ → PROPOSE → PREVIEW → APPROVE → WRITE → RECALCULATE → VERIFY → SAVE

运行：
    python mcp/project_mcp_server.py            # stdio，供 MCP 客户端拉起

约定：
    - 默认强制备份后原生就地保存，以保留源 MPP 视图层
    - 每次执行前自动备份、执行后重开文件双读验证
"""
from __future__ import annotations

import json
import sys
import traceback
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import change_set as CS  # noqa: E402
from com_bridge import (BridgeError, ROOT, com_snapshot, export_excel, mpxj_read, create_plan)  # noqa: E402
from note_fields import read_fields, merge_fields
from live_project import run_live

PROJECT_NAME = "Project MCP for codex"
SERVER_NAME = "project-mcp-for-codex"
SERVER_VERSION = "0.3.0"
PROTOCOL_VERSION = "2024-11-05"

STATE: dict = {"active_file": None, "_cache": {}}


def log(msg: str) -> None:
    print(f"[mcp] {msg}", file=sys.stderr, flush=True)


# ---------------------------------------------------------------- 数据访问
def active_or(file: str | None) -> str:
    path = file or STATE.get("active_file")
    if not path:
        raise ValueError("尚未打开项目：请先调用 project_open 或传入 file 参数")
    if not Path(path).exists():
        raise ValueError(f"文件不存在: {path}")
    return str(path)


def read_plan(file: str | None = None, refresh: bool = False) -> dict:
    path = active_or(file)
    cached = STATE["_cache"].get(path)
    stat = Path(path).stat()
    signature = (stat.st_mtime_ns, stat.st_size)
    if cached and not refresh and cached[0] == signature:
        return cached[1]
    data = mpxj_read(path)
    STATE["_cache"][path] = (signature, data)
    return data


def task_view(t: dict) -> dict:
    return {
        "unique_id": t.get("unique_id"), "id": t.get("id"), "wbs": t.get("wbs"),
        "outline_level": t.get("outline_level"), "name": t.get("name"),
        "parent_name": t.get("parent_name"), "summary": t.get("is_summary"),
        "milestone": t.get("is_milestone"), "start": t.get("start"), "finish": t.get("finish"),
        "duration_days": (t.get("duration") or {}).get("days"),
        "percent_complete": t.get("percent_complete"),
        "actual_start": t.get("actual_start"), "actual_finish": t.get("actual_finish"),
        "constraint": t.get("constraint"), "constraint_date": t.get("constraint_date"),
        "critical": t.get("critical"), "total_slack_days": (t.get("total_slack") or {}).get("days"),
        "predecessors": [p.get("predecessor_id") for p in t.get("predecessors") or []],
        "assignees": t.get("assignees"), "notes": t.get("notes"),
    }


# ---------------------------------------------------------------- 工具实现
def t_create_plan(args: dict) -> dict:
    return create_plan(args["xml"], args["output"], preview=bool(args.get("preview", True)),
                       source_sha256=args.get("source_sha256", ""))


def t_get_note_fields(args: dict) -> dict:
    # Use COM's full Notes value, not the truncated MPXJ display excerpt.
    snap = com_snapshot(active_or(args.get('file')))
    if args.get('unique_id') is None:
        rows = []
        for task in snap['tasks']:
            parsed = read_fields(task.get('notes') or '')
            if args.get('field_name') and parsed['fields'].get(args['field_name']) != str(args.get('equals', '')):
                continue
            rows.append({'unique_id': task['unique_id'], 'name': task['name'], **parsed})
        return {'count': len(rows), 'tasks': rows}
    task = next((t for t in snap['tasks'] if t.get('unique_id') == int(args['unique_id'])), None)
    if task is None:
        raise ValueError('找不到任务')
    return {'unique_id': task['unique_id'], 'name': task['name'], **read_fields(task.get('notes') or '')}


def t_preview_note_fields(args: dict) -> dict:
    current = t_get_note_fields(args)
    proposed = merge_fields(current['notes'], args['fields'])
    result = t_preview_changes({'file': args.get('file'), 'title': args.get('title') or '更新备注中的任务字段',
        'ops': [{'op': 'set_task_fields', 'unique_id': int(args['unique_id']), 'fields': {'Notes': proposed},
                 'expected_notes': current['notes']}]})
    result['note_fields'] = read_fields(proposed)['fields']
    return result


def t_project_open(args: dict) -> dict:
    path = str(Path(args["file"]).resolve())
    if not Path(path).exists():
        raise ValueError(f"文件不存在: {path}")
    STATE["active_file"] = path
    STATE["_cache"].pop(path, None)
    plan = read_plan(path)
    p = plan["project"]
    return {
        "opened": path,
        "name": p.get("title") or Path(path).stem,
        "start": p.get("start_date"), "finish": p.get("finish_date"),
        "default_calendar": p.get("default_calendar"),
        "tasks": plan["stats"]["task_count"], "relations": plan["stats"]["relation_count"],
        "milestones": plan["stats"]["milestone_count"], "resources": plan["stats"]["resource_count"],
        "note": "只读打开，未修改文件",
    }


def t_project_close(args: dict) -> dict:
    STATE["active_file"] = None
    STATE["_cache"].clear()
    return {"closed": True}


def t_get_project_info(args: dict) -> dict:
    plan = read_plan(args.get("file"))
    return {"project": plan["project"], "stats": plan["stats"]}


def t_get_schedule_summary(args: dict) -> dict:
    plan = read_plan(args.get("file"))
    tasks = plan["tasks"]
    leaves = [t for t in tasks if not t.get("is_summary")]
    today = args.get("today")
    overdue = [t for t in leaves if (t.get("percent_complete") or 0) < 100
               and (t.get("finish") or "")[:10] and (t.get("finish") or "")[:10] < today] if today else []
    return {
        "title": plan["project"].get("title"),
        "start": plan["project"].get("start_date"), "finish": plan["project"].get("finish_date"),
        "task_count": len(tasks), "leaf_count": len(leaves),
        "level1": [{"uid": t["unique_id"], "wbs": t["wbs"], "name": t["name"],
                    "start": t["start"], "finish": t["finish"]}
                   for t in tasks if t.get("outline_level") == 1],
        "milestones": [{"uid": t["unique_id"], "name": t["name"], "date": t["start"]}
                       for t in tasks if t.get("is_milestone")],
        "critical_count": plan["stats"].get("critical_count"),
        **({"overdue_count": len(overdue)} if today else {}),
    }


def t_list_tasks(args: dict) -> dict:
    plan = read_plan(args.get("file"))
    tasks = plan["tasks"]
    if args.get("name_contains"):
        tasks = [t for t in tasks if args["name_contains"] in (t.get("name") or "")]
    if args.get("outline_level") is not None:
        tasks = [t for t in tasks if t.get("outline_level") == args["outline_level"]]
    if args.get("only_leaf"):
        tasks = [t for t in tasks if not t.get("is_summary")]
    if args.get("milestones_only"):
        tasks = [t for t in tasks if t.get("is_milestone")]
    limit = int(args.get("limit") or 100)
    return {"total_matched": len(tasks), "returned": min(limit, len(tasks)),
            "tasks": [task_view(t) for t in tasks[:limit]]}


def t_get_task(args: dict) -> dict:
    plan = read_plan(args.get("file"))
    uid = int(args["unique_id"])
    for t in plan["tasks"]:
        if int(t.get("unique_id") or -1) == uid:
            return task_view(t)
    raise ValueError(f"找不到 UniqueID={uid}")


def t_find_tasks(args: dict) -> dict:
    return t_list_tasks({**args, "name_contains": args.get("query") or args.get("name_contains")})


def t_get_dependencies(args: dict) -> dict:
    plan = read_plan(args.get("file"))
    uid = args.get("unique_id")
    links = plan.get("relations") or []
    if uid is not None:
        uid = int(uid)
        links = [l for l in links if l.get("from_id") == uid or l.get("to_id") == uid]
    return {"count": len(links), "relations": [{
        "from_uid": l.get("from_id"), "from_name": l.get("from_name"),
        "to_uid": l.get("to_id"), "to_name": l.get("to_name"),
        "type": l.get("type"), "lag": (l.get("lag") or {}).get("text"),
    } for l in links[: int(args.get("limit") or 200)]]}


def t_get_view_layer(args: dict) -> dict:
    """走 COM：MPXJ 读不到视图层。"""
    snap = com_snapshot(active_or(args.get("file")))
    vl = snap["view_layer"]
    return {
        "current_view": snap["project"].get("current_view"),
        "current_table": snap["project"].get("current_table"),
        "current_filter": snap["project"].get("current_filter"),
        "current_group": snap["project"].get("current_group"),
        "collections": {k: [x["name"] for x in v] for k, v in vl.items() if isinstance(v, list) and v and isinstance(v[0], dict)},
        "lists": {k: v for k, v in vl.items() if k.endswith("_list")},
        "note": "甘特条形样式在 COM 对象模型中没有读取接口，无法通过 MCP 读取",
    }


def t_list_calendars(args: dict) -> dict:
    plan = read_plan(args.get("file"))
    return {"calendars": plan["calendars"]}


def t_list_resources(args: dict) -> dict:
    plan = read_plan(args.get("file"))
    return {"resources": plan["resources"]}


def t_get_baselines(args: dict) -> dict:
    plan = read_plan(args.get("file"))
    rows = []
    for t in plan["tasks"]:
        if t.get("baseline_start") or t.get("baseline_finish"):
            rows.append({"unique_id": t["unique_id"], "wbs": t["wbs"], "name": t["name"],
                         "baseline_start": t.get("baseline_start"), "baseline_finish": t.get("baseline_finish")})
    return {"with_baseline": len(rows), "tasks": rows[:100],
            "note": "本文件基线为空" if not rows else ""}


def t_get_critical_path(args: dict) -> dict:
    plan = read_plan(args.get("file"))
    crit = [task_view(t) for t in plan["tasks"] if t.get("critical")]
    return {"count": len(crit), "tasks": crit,
            "note": "依赖关系为空时，关键路径仅反映 Project 的计算结果" if not plan["stats"]["relation_count"] else ""}


def t_preview_changes(args: dict) -> dict:
    cs = CS.build(title=args.get("title") or "未命名改动",
                  source_file=active_or(args.get("file")),
                  ops=args["ops"],
                  requested_by=args.get("requested_by") or "agent")
    return {"change_set_id": cs.id, "risk_level": cs.risk_level, "status": cs.status,
            "save_mode": cs.save_mode,
            "requires_approval": cs.risk_level in CS.APPROVAL_REQUIRED_RISK,
            "errors": cs.errors, "preview": cs.preview, "summary": cs.human_summary(),
            "note": "默认就地保存（保留视图层）；就地保存会覆盖源文件，执行前会自动备份，"
                    "需在 project_apply_changes 中传 allow_in_place=true。"}


def t_apply_changes(args: dict) -> dict:
    cs = CS.apply(change_set_id=args["change_set_id"],
                  approved_by=args.get("approved_by") or "",
                  note=args.get("note") or "",
                  keep_visible=bool(args.get("keep_visible")),
                  allow_in_place=bool(args.get("allow_in_place", True)))
    STATE["_cache"].pop(cs.output_file or "", None)
    view = (cs.verification or {}).get("view_layer") or {}
    return {"change_set_id": cs.id, "status": cs.status, "save_mode": cs.save_mode,
            "output_file": cs.output_file, "backup_file": cs.backup_file, "log_file": cs.log_file,
            "verification": cs.verification,
            "view_layer_preserved": view.get("preserved"),
            "view_layer_warning": view.get("warning"),
            "summary": cs.human_summary()}


def t_verify_changes(args: dict) -> dict:
    cs = CS.ChangeSet.load(args["change_set_id"])
    result = CS.verify(cs)
    cs.verification = result
    cs.status = "verified" if result.get("ok") else "applied_with_warnings"
    cs.save()
    return result


def t_rollback_changes(args: dict) -> dict:
    return CS.rollback(args["change_set_id"])


def t_list_change_sets(args: dict) -> dict:
    items = []
    for path in sorted(CS.CHANGESET_DIR.glob("CS-*.json")):
        d = json.loads(path.read_text(encoding="utf-8"))
        items.append({"id": d["id"], "title": d["title"], "risk": d["risk_level"],
                      "status": d["status"], "created_at": d["created_at"],
                      "output_file": d.get("output_file")})
    return {"count": len(items), "change_sets": items[-30:]}


def t_append_note(args: dict) -> dict:
    return t_preview_changes({
        "title": args.get("title") or "追加现场备注",
        "ops": [{"op": "append_task_note", "unique_id": int(args["unique_id"]),
                 "note": args["note"]}],
        "file": args.get("file"), "requested_by": args.get("requested_by"),
    })


def t_set_actual_progress(args: dict) -> dict:
    ops = []
    uid = int(args["unique_id"])
    if args.get("actual_start"):
        ops.append({"op": "set_actual_start", "unique_id": uid, "date": args["actual_start"],
                    "note": args.get("note")})
    if args.get("actual_finish"):
        ops.append({"op": "set_actual_finish", "unique_id": uid, "date": args["actual_finish"]})
    if args.get("percent_complete") is not None:
        ops.append({"op": "set_percent_complete", "unique_id": uid, "percent": args["percent_complete"]})
    if not ops:
        raise ValueError("至少提供 actual_start / actual_finish / percent_complete 之一")
    return t_preview_changes({"title": args.get("title") or f"更新进度 UID={uid}", "ops": ops,
                              "file": args.get("file")})


def t_export_excel(args: dict) -> dict:
    path = active_or(args.get("file"))
    plan = mpxj_read(path)
    tmp = ROOT / "mcp" / "work" / "export_source.json"
    tmp.parent.mkdir(parents=True, exist_ok=True)
    tmp.write_text(json.dumps(plan, ensure_ascii=False), encoding="utf-8")
    out = args.get("output") or str(ROOT / "exports" / f"{Path(path).stem}_任务清单.xlsx")
    export_excel(str(tmp), out)
    return {"output": out, "source": path}


# ---------------------------------------------------------------- 工具清单
TOOLS: list[dict] = [
    {"name": "project_live_read", "description": "读取已打开的活动 Project 内存状态；明确传 file。只连接既有进程，保留窗口，不保存。",
     "inputSchema": {"type": "object", "properties": {"file": {"type": "string"}}, "required": ["file"]},
     "handler": lambda args: run_live('read', args)},
    {"name": "project_live_preview", "description": "开着 Project 预览任务字段/备注字段改动；需先保存但无需关闭，预览不写入。",
     "inputSchema": {"type": "object", "properties": {"file": {"type": "string"}, "title": {"type": "string"},
         "ops": {"type": "array", "items": {"type": "object"}, "minItems": 1, "maxItems": 20}}, "required": ["file", "ops"]},
     "handler": lambda args: run_live('preview', args)},
    {"name": "project_live_apply", "description": "执行已确认在线变更，原生撤销组、备份和内存核验；保持 Project 打开，由用户查看后保存。",
     "inputSchema": {"type": "object", "properties": {"change_id": {"type": "string"}, "approved_by": {"type": "string"}},
                     "required": ["change_id", "approved_by"]},
     "handler": lambda args: run_live('apply', args)},
    {"name": "project_get_task_note_fields", "description": "读取完整备注的逐行字段；不传 unique_id 时批量读取，可用 field_name/equals 筛选工程量、状态等字段。",
     "inputSchema": {"type": "object", "properties": {"file": {"type": "string"}, "unique_id": {"type": "integer"},
         "field_name": {"type": "string"}, "equals": {"type": "string"}}}, "handler": t_get_note_fields},
    {"name": "project_preview_task_note_fields", "description": "预览更新备注中的结构化任务字段；保留自由说明文字，通过 Change Set 执行。",
     "inputSchema": {"type": "object", "properties": {"file": {"type": "string"}, "unique_id": {"type": "integer"},
         "fields": {"type": "object", "additionalProperties": {"type": ["string", "number", "boolean"]}},
         "title": {"type": "string"}}, "required": ["unique_id", "fields"]}, "handler": t_preview_note_fields},
    {"name": "project_create_from_xml", "description": "从基础 Project XML 创建新 MPP：默认只预览；执行需 source_sha256。拒绝覆盖，Project 须已关闭；生成后核验。",
     "inputSchema": {"type": "object", "properties": {
         "xml": {"type": "string"}, "output": {"type": "string"},
         "preview": {"type": "boolean", "default": True}, "source_sha256": {"type": "string"}},
         "required": ["xml", "output"]}, "handler": t_create_plan},
    {"name": "project_open", "description": "只读打开一个 Microsoft Project 文件并设为当前项目",
     "inputSchema": {"type": "object", "properties": {"file": {"type": "string"}}, "required": ["file"]},
     "handler": t_project_open},
    {"name": "project_close", "description": "关闭当前项目（不写入任何文件）",
     "inputSchema": {"type": "object", "properties": {}}, "handler": t_project_close},
    {"name": "project_get_project_info", "description": "读取项目基础信息与规模统计",
     "inputSchema": {"type": "object", "properties": {"file": {"type": "string"}}}, "handler": t_get_project_info},
    {"name": "project_get_schedule_summary", "description": "读取计划概要（一级计划、里程碑、可按 today 统计逾期）",
     "inputSchema": {"type": "object", "properties": {"file": {"type": "string"}, "today": {"type": "string"}}},
     "handler": t_get_schedule_summary},
    {"name": "project_list_tasks", "description": "按条件列出任务（名称/层级/仅明细/仅里程碑/条数上限）",
     "inputSchema": {"type": "object", "properties": {
         "file": {"type": "string"}, "name_contains": {"type": "string"},
         "outline_level": {"type": "integer"}, "only_leaf": {"type": "boolean"},
         "milestones_only": {"type": "boolean"}, "limit": {"type": "integer"}}},
     "handler": t_list_tasks},
    {"name": "project_find_tasks", "description": "按名称关键字查找任务（project_list_tasks 的简化入口）",
     "inputSchema": {"type": "object", "properties": {"query": {"type": "string"}, "file": {"type": "string"}},
                     "required": ["query"]},
     "handler": t_find_tasks},
    {"name": "project_get_task", "description": "按 UniqueID 读取单个任务（推荐用 UniqueID 作为稳定标识）",
     "inputSchema": {"type": "object", "properties": {"unique_id": {"type": "integer"}, "file": {"type": "string"}},
                     "required": ["unique_id"]},
     "handler": t_get_task},
    {"name": "project_get_dependencies", "description": "读取依赖关系（可按任务 UniqueID 过滤）",
     "inputSchema": {"type": "object", "properties": {"unique_id": {"type": "integer"},
                                                      "file": {"type": "string"}, "limit": {"type": "integer"}}},
     "handler": t_get_dependencies},
    {"name": "project_get_view_layer", "description": "读取视图层（视图/表格/筛选/分组/报表集合，走 COM）",
     "inputSchema": {"type": "object", "properties": {"file": {"type": "string"}}}, "handler": t_get_view_layer},
    {"name": "project_list_calendars", "description": "列出日历",
     "inputSchema": {"type": "object", "properties": {"file": {"type": "string"}}}, "handler": t_list_calendars},
    {"name": "project_list_resources", "description": "列出资源",
     "inputSchema": {"type": "object", "properties": {"file": {"type": "string"}}}, "handler": t_list_resources},
    {"name": "project_get_baselines", "description": "读取基线（只读；本文件基线为空）",
     "inputSchema": {"type": "object", "properties": {"file": {"type": "string"}}}, "handler": t_get_baselines},
    {"name": "project_get_critical_path", "description": "读取关键任务集合",
     "inputSchema": {"type": "object", "properties": {"file": {"type": "string"}}}, "handler": t_get_critical_path},
    {"name": "project_preview_changes", "description": "生成 Change Set 预览（不写入文件）：READ→PROPOSE→PREVIEW",
     "inputSchema": {"type": "object", "properties": {
         "title": {"type": "string"}, "ops": {"type": "array", "items": {"type": "object"}},
         "file": {"type": "string"}, "requested_by": {"type": "string"}}, "required": ["ops"]},
     "handler": t_preview_changes},
    {"name": "project_apply_changes",
     "description": ("执行已批准的 Change Set：备份→COM 原生写入→重算→保存→双读验证。"
                     "默认就地保存（allow_in_place=true，会覆盖源文件但完整保留甘特图视图层）；"
                     "allow_in_place=false 时另存为新文件，会丢失甘特图上的图形/文本对象与 OLE 链接。"),
     "inputSchema": {"type": "object", "properties": {
         "change_set_id": {"type": "string"}, "approved_by": {"type": "string"},
         "note": {"type": "string"}, "keep_visible": {"type": "boolean"},
         "allow_in_place": {"type": "boolean", "description": "就地保存并覆盖源文件（执行前自动备份），默认 true"}},
         "required": ["change_set_id", "approved_by"]},
     "handler": t_apply_changes},
    {"name": "project_verify_changes", "description": "重开输出文件核对 Change Set 是否逐项落盘",
     "inputSchema": {"type": "object", "properties": {"change_set_id": {"type": "string"}},
                     "required": ["change_set_id"]},
     "handler": t_verify_changes},
    {"name": "project_rollback_changes", "description": "回滚 Change Set（删除本次输出，确认源文件完好）",
     "inputSchema": {"type": "object", "properties": {"change_set_id": {"type": "string"}},
                     "required": ["change_set_id"]},
     "handler": t_rollback_changes},
    {"name": "project_list_change_sets", "description": "列出历史 Change Set 及状态",
     "inputSchema": {"type": "object", "properties": {}}, "handler": t_list_change_sets},
    {"name": "project_append_task_note", "description": "生成「追加现场备注」的 Change Set（不覆盖原备注）",
     "inputSchema": {"type": "object", "properties": {
         "unique_id": {"type": "integer"}, "note": {"type": "string"},
         "title": {"type": "string"}, "file": {"type": "string"}}, "required": ["unique_id", "note"]},
     "handler": t_append_note},
    {"name": "project_set_actual_progress", "description": "生成「更新实际进度」的 Change Set（实际开始/完成/完成率）",
     "inputSchema": {"type": "object", "properties": {
         "unique_id": {"type": "integer"}, "actual_start": {"type": "string"},
         "actual_finish": {"type": "string"}, "percent_complete": {"type": "number"},
         "title": {"type": "string"}, "note": {"type": "string"}, "file": {"type": "string"}},
         "required": ["unique_id"]},
     "handler": t_set_actual_progress},
    {"name": "project_export_excel", "description": "把当前项目任务清单导出为 Excel",
     "inputSchema": {"type": "object", "properties": {"file": {"type": "string"}, "output": {"type": "string"}}},
     "handler": t_export_excel},
]

TOOL_MAP = {t["name"]: t for t in TOOLS}


# ---------------------------------------------------------------- JSON-RPC
def send(payload: dict) -> None:
    sys.stdout.write(json.dumps(payload, ensure_ascii=False) + "\n")
    sys.stdout.flush()


def handle(request: dict) -> dict | None:
    method = request.get("method")
    rid = request.get("id")

    if method == "initialize":
        return {"jsonrpc": "2.0", "id": rid, "result": {
            "protocolVersion": PROTOCOL_VERSION,
            "capabilities": {"tools": {"listChanged": False}},
            "serverInfo": {"name": SERVER_NAME, "version": SERVER_VERSION},
        }}
    if method in ("notifications/initialized", "initialized"):
        return None
    if method == "ping":
        return {"jsonrpc": "2.0", "id": rid, "result": {}}
    if method == "tools/list":
        return {"jsonrpc": "2.0", "id": rid, "result": {
            "tools": [{"name": t["name"], "description": t["description"], "inputSchema": t["inputSchema"]}
                      for t in TOOLS]}}
    if method == "tools/call":
        params = request.get("params") or {}
        name = params.get("name")
        args = params.get("arguments") or {}
        tool = TOOL_MAP.get(name)
        if not tool:
            return {"jsonrpc": "2.0", "id": rid, "result": {
                "content": [{"type": "text", "text": f"未知工具: {name}"}], "isError": True}}
        try:
            result = tool["handler"](args)
            return {"jsonrpc": "2.0", "id": rid, "result": {
                "content": [{"type": "text", "text": json.dumps(result, ensure_ascii=False, indent=2, default=str)}],
                "isError": False}}
        except Exception as exc:  # 工具异常必须作为 isError 返回，不能中断服务
            log(f"tool {name} 失败: {exc}\n{traceback.format_exc()}")
            return {"jsonrpc": "2.0", "id": rid, "result": {
                "content": [{"type": "text", "text": f"{type(exc).__name__}: {exc}"}], "isError": True}}
    if rid is None:
        return None
    return {"jsonrpc": "2.0", "id": rid, "error": {"code": -32601, "message": f"Method not found: {method}"}}


def main() -> int:
    # MCP stdio is UTF-8 JSON-RPC. Windows console code pages otherwise corrupt
    # Chinese tool metadata and file paths when this server is launched by Codex.
    for stream in (sys.stdin, sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8")
    log(f"{SERVER_NAME} {SERVER_VERSION} 启动（stdio）")
    for line in sys.stdin:
        line = line.strip()
        if not line:
            continue
        try:
            request = json.loads(line)
        except json.JSONDecodeError as exc:
            send({"jsonrpc": "2.0", "id": None, "error": {"code": -32700, "message": str(exc)}})
            continue
        response = handle(request)
        if response is not None:
            send(response)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
