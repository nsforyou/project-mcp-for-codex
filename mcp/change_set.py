# -*- coding: utf-8 -*-
r"""Change Set：MCP 写入的事务单元。

流程：READ → PROPOSE → PREVIEW → USER APPROVE → WRITE → RECALCULATE → VERIFY → SAVE
安全约定：
    - 默认强制备份后就地保存，以保留原 MPP 视图层
    - 应用前自动备份源文件
    - 应用后重新打开输出文件做双读验证（COM），并与预期逐项核对
    - 全过程写 JSONL 日志
"""
from __future__ import annotations

import json
import time
import uuid
from dataclasses import dataclass, field, asdict
from pathlib import Path
from com_runtime import exclusive_project

from com_bridge import (ROOT, backup_file, com_apply, com_recalculate_and_verify, com_snapshot,
                        compare_view_streams, export_excel, mpxj_read, sha256)

CHANGESET_DIR = ROOT / "mcp" / "changesets"
EXPORT_DIR = ROOT / "exports"
LOG_DIR = ROOT / "logs"
APPROVAL_REQUIRED_RISK = {"medium", "high"}

# op 名称 -> (是否需要确认, 风险等级, 说明)
OP_SPECS: dict[str, dict] = {
    "set_actual_start": {"risk": "medium", "desc": "填写实际开始（不改变计划日期）"},
    "set_actual_finish": {"risk": "medium", "desc": "填写实际完成"},
    "set_percent_complete": {"risk": "medium", "desc": "填写完成率（Project 可能重算）"},
    "set_remaining_duration": {"risk": "medium", "desc": "填写剩余工期"},
    "set_task_fields": {"risk": "medium", "desc": "通用字段写入（日期字段自动带默认时分）"},
    "append_task_note": {"risk": "low", "desc": "追加备注，不覆盖原文"},
    "shift_task_dates": {"risk": "high", "desc": "平移计划日期（会触发下游排程变化）"},
    "set_predecessors": {"risk": "high", "desc": "写入任务的前置依赖（整串文本，行号引用）"},
    "create_milestone": {"risk": "high", "desc": "新建里程碑并插到指定位置"},
}


@dataclass
class ChangeSet:
    id: str
    title: str
    source_file: str
    source_sha256: str
    requested_by: str
    created_at: str
    ops: list[dict]
    preview: list[dict] = field(default_factory=list)
    risk_level: str = "low"
    status: str = "preview"          # preview -> approved -> applied -> verified / failed
    save_mode: str = "in_place"      # in_place（保留视图层）| new_revision（另存，会丢视图图形数据）
    approved_by: str | None = None
    approved_at: str | None = None
    output_file: str | None = None
    backup_file: str | None = None
    log_file: str | None = None
    verification: dict | None = None
    output_sha256: str | None = None
    errors: list[str] = field(default_factory=list)

    # ---------------------------------------------------------------- 序列化
    def to_dict(self) -> dict:
        return asdict(self)

    def save(self) -> Path:
        CHANGESET_DIR.mkdir(parents=True, exist_ok=True)
        path = CHANGESET_DIR / f"{self.id}.json"
        path.write_text(json.dumps(self.to_dict(), ensure_ascii=False, indent=2), encoding="utf-8")
        return path

    @staticmethod
    def load(change_set_id: str) -> "ChangeSet":
        path = CHANGESET_DIR / f"{change_set_id}.json"
        if not path.exists():
            raise FileNotFoundError(f"变更集不存在: {change_set_id}")
        return ChangeSet(**json.loads(path.read_text(encoding="utf-8")))

    def human_summary(self) -> str:
        lines = [
            f"Change Set {self.id}",
            f"标题      : {self.title}",
            f"源文件    : {self.source_file}",
            f"源文件哈希: {self.source_sha256[:16]}",
            f"风险等级  : {self.risk_level}",
            f"状态      : {self.status}",
            "",
            "改动明细（Before -> Proposed）:",
        ]
        for item in self.preview:
            lines.append(f"  · UID={item.get('unique_id')} {item.get('wbs','')} {item.get('name','')}")
            for k, v in item.get("changes", {}).items():
                lines.append(f"      {k}: {v.get('before')}  ->  {v.get('proposed')}")
            if item.get("append_note"):
                lines.append(f"      Notes 追加: {item['append_note']}")
        lines.append("")
        lines.append("预期影响:")
        lines.append("  · Project 将重新计算（派生字段与摘要上卷可能变化）")
        if self.save_mode == "in_place":
            lines.append("  · 保存方式：就地保存（FileSave）—— 已实测**完整保留**视图层"
                         "（甘特图上的文本/图形对象、OLE 链接、表格列定义）")
            lines.append("  · 就地保存会覆盖源文件，执行前自动备份到 backups/")
        else:
            lines.append("  · 保存方式：另存为新文件（FileSaveAs）—— **已知缺陷**："
                         "会丢失甘特图上的文本/图形对象与 OLE 链接")
        return "\n".join(lines)


# ---------------------------------------------------------------- 构建
def _task_index(snapshot: dict) -> dict[int, dict]:
    return {int(t["unique_id"]): t for t in snapshot["tasks"]}


FIELD_KEYS = {
    "Name": "name", "Start": "start", "Finish": "finish", "Duration": "duration",
    "ActualStart": "actual_start", "ActualFinish": "actual_finish",
    "PercentComplete": "percent_complete", "RemainingDuration": "remaining_duration",
    "Notes": "notes", "Milestone": "milestone", "Manual": "manual",
    "ResourceNames": "resource_names", "Predecessors": "predecessors",
    "ConstraintType": "constraint_type", "ConstraintDate": "constraint_date",
}


def build(title: str, source_file: str, ops: list[dict], requested_by: str = "agent") -> ChangeSet:
    """读取当前状态 → 生成预览（不写入任何东西）。"""
    source = Path(source_file)
    if not source.exists():
        raise FileNotFoundError(f"源文件不存在: {source}")

    initial_hash = sha256(str(source))
    snapshot = com_snapshot(str(source))
    if sha256(str(source)) != initial_hash:
        raise RuntimeError('预览读取期间源文件发生变化，请重新预览')
    index = _task_index(snapshot)

    cs = ChangeSet(
        id=f"CS-{time.strftime('%Y%m%d-%H%M%S')}-{uuid.uuid4().hex[:4]}",
        title=title,
        source_file=str(source),
        source_sha256=initial_hash,
        requested_by=requested_by,
        created_at=time.strftime("%Y-%m-%d %H:%M:%S"),
        ops=ops,
    )

    risks = []
    for op in ops:
        kind = op["op"]
        spec = OP_SPECS.get(kind)
        if not spec:
            cs.errors.append(f"不支持的 op: {kind}")
            continue
        risks.append(spec["risk"])
        uid = int(op["unique_id"]) if op.get("unique_id") is not None else None
        task = index.get(uid) if uid is not None else None
        if op.get("op") == "create_milestone":
            cs.preview.append({
                "op": "create_milestone",
                "new_task": op.get("name"), "milestone_date": op.get("date"),
                "insert_before_row": op.get("insert_before_row"),
                "outline_level": op.get("outline_level"),
                "changes": {"Milestone": {"before": None, "proposed": op.get("date")}},
            })
            continue
        if task is None and uid is not None:
            cs.errors.append(f"找不到 UniqueID={uid} 的任务")
            continue
        if task is not None and 'expected_notes' in op and (task.get('notes') or '') != op['expected_notes']:
            cs.errors.append(f'UID={uid} 的备注在读取后已变化，请重新预览')
            continue
        if op.get("op") == "set_predecessors":
            cs.preview.append({
                "op": "set_predecessors",
                "row_id": op.get("row_id"), "name": f"行 {op.get('row_id')}",
                "changes": {"Predecessors": {"before": None, "proposed": op.get("predecessors")}},
            })
            continue
        changes: dict[str, dict] = {}
        if kind == "shift_task_dates":
            changes["ShiftDays"] = {"before": task.get("start"), "proposed": f"+{op.get('days')} 天"}
        elif kind == "set_task_fields" or kind.startswith("set_"):
            fields = op.get("fields") or {}
            if kind == "set_actual_start":
                fields = {"ActualStart": op["date"]}
            elif kind == "set_actual_finish":
                fields = {"ActualFinish": op["date"]}
            elif kind == "set_percent_complete":
                fields = {"PercentComplete": op["percent"]}
            elif kind == "set_remaining_duration":
                fields = {"RemainingDuration": op["duration"]}
            for field_name, value in fields.items():
                if field_name not in FIELD_KEYS:
                    cs.errors.append(f"尚未实现可靠核验的字段: {field_name}")
                    continue
                key = FIELD_KEYS[field_name]
                changes[field_name] = {"before": task.get(key), "proposed": value}
        item = {
            "unique_id": uid,
            "wbs": task.get("wbs"),
            "name": task.get("name"),
            "outline_level": task.get("outline_level"),
            "changes": changes,
        }
        note = op.get("append_note") or op.get("note")
        if note:
            item["append_note"] = note
            item["changes"]["Notes"] = {"before": task.get("notes"), "proposed": f"{task.get('notes') or ''}\n{note}".strip()}
        cs.preview.append(item)

    cs.risk_level = "high" if "high" in risks else ("medium" if "medium" in risks else "low")
    cs.save()
    return cs


# ---------------------------------------------------------------- 应用
def apply(change_set_id: str, approved_by: str, note: str = "", keep_visible: bool = False,
          allow_in_place: bool = False) -> ChangeSet:
    """用户确认后执行写入。

    allow_in_place：就地保存会覆盖源文件（已自动备份）。这是**唯一能保留视图层
    （甘特图上的文本框/图形对象、OLE 链接、表格列定义）**的保存方式，
    因此对真实工程文件应当传 true；否则会退回另存为新文件并丢失这些内容。
    """
    cs = ChangeSet.load(change_set_id)
    if cs.errors:
        cs.status = "failed"
        cs.save()
        raise RuntimeError(f"变更集包含错误，已拒绝执行: {cs.errors}")
    if cs.risk_level in APPROVAL_REQUIRED_RISK and not approved_by:
        raise PermissionError("该变更集需要用户显式批准（approved_by 不能为空）")
    if not Path(cs.source_file).exists():
        raise FileNotFoundError(cs.source_file)
    if sha256(cs.source_file) != cs.source_sha256:
        raise RuntimeError("源文件在预览之后已被修改，请重新生成变更集（防止在错误状态上写入）")

    if cs.status not in ("preview", "approved"):
        raise RuntimeError(f"变更集状态 {cs.status} 不允许重复执行，请重新预览")
    # 2) 输出目标
    #    in_place      ：就地保存（FileSave）—— 实测唯一能保留视图图形对象与 OLE 链接的方式
    #    new_revision  ：另存为新文件（FileSaveAs）—— 会丢失甘特图上的文本框/图形与链接，仅作兼容保留
    EXPORT_DIR.mkdir(parents=True, exist_ok=True)
    in_place = cs.save_mode == "in_place"
    if in_place and not allow_in_place:
        raise PermissionError(
            "就地保存会覆盖源文件（已自动备份）。需显式传入 allow_in_place=true 才执行。"
            "注意：改用另存为新文件（new_revision）会丢失甘特图上的图形/文本对象与 OLE 链接。")
    cs.backup_file = backup_file(cs.source_file)
    out = Path(cs.source_file) if in_place else EXPORT_DIR / f"{Path(cs.source_file).stem}_MCP_{cs.id}.mpp"

    # 3) 组装 COM 写入 spec
    spec_ops = []
    for op in cs.ops:
        kind = op["op"]
        if kind in ("set_actual_start",):
            spec_ops.append({"op": "set_fields", "unique_id": op["unique_id"],
                             "fields": {"ActualStart": op["date"]}, "append_note": op.get("note")})
        elif kind == "set_actual_finish":
            spec_ops.append({"op": "set_fields", "unique_id": op["unique_id"],
                             "fields": {"ActualFinish": op["date"]}, "append_note": op.get("note")})
        elif kind == "set_percent_complete":
            spec_ops.append({"op": "set_fields", "unique_id": op["unique_id"],
                             "fields": {"PercentComplete": op["percent"]}, "append_note": op.get("note")})
        elif kind == "set_remaining_duration":
            spec_ops.append({"op": "set_fields", "unique_id": op["unique_id"],
                             "fields": {"RemainingDuration": op["duration"]}, "append_note": op.get("note")})
        elif kind == "set_task_fields":
            spec_ops.append({"op": "set_fields", "unique_id": op["unique_id"],
                             "fields": op.get("fields", {}), "append_note": op.get("note")})
        elif kind == "append_task_note":
            spec_ops.append({"op": "append_note", "unique_id": op["unique_id"],
                             "append_note": op.get("note") or op.get("append_note")})
        elif kind == "shift_task_dates":
            spec_ops.append({"op": "shift_dates", "unique_id": op["unique_id"],
                             "days": op.get("days", 0), "append_note": op.get("note")})
        elif kind == "set_predecessors":
            spec_ops.append({"op": "set_predecessors", "row_id": op["row_id"],
                             "predecessors": op["predecessors"]})
        elif kind == "create_milestone":
            spec_ops.append({"op": "create_milestone", "name": op["name"], "date": op["date"],
                             "insert_before_row": op.get("insert_before_row"),
                             "outline_level": op.get("outline_level")})
        else:
            raise RuntimeError(f"该 op 尚未在 COM 通道实现: {kind}")

    LOG_DIR.mkdir(parents=True, exist_ok=True)
    log_file = LOG_DIR / f"mcp_apply_{cs.id}.json"
    spec = {
        "source": cs.source_file,
        "expected_source_sha256": cs.source_sha256,
        "save_in_place": in_place,
        "log": str(log_file),
        "ops": spec_ops,
    }
    if not in_place:
        spec["save_as"] = str(out)

    cs.status = "approved"
    cs.approved_by = approved_by
    cs.approved_at = time.strftime("%Y-%m-%d %H:%M:%S")
    cs.save()

    # 4) 执行写入：就地保存（保留视图层）或另存为新文件
    cs.output_file = str(out)
    cs.log_file = str(log_file)
    # Persist the recovery metadata before verification.  A verification error
    # must never hide a completed COM save or leave the Change Set untraceable.
    cs.save()

    try:
        result = com_apply(spec, keep_visible=keep_visible)
    except Exception as exc:
        cs.status = "failed"
        cs.errors.append(f"写入未完成或结果不确定，请核查文件和日志: {exc}")
        cs.save()
        raise
    cs.output_sha256 = sha256(str(out))
    cs.save()

    # 5) 重开输出文件复核（COM 双读）
    try:
        cs.verification = verify(cs)
        cs.status = "verified" if cs.verification.get("ok") else "applied_with_warnings"
    except Exception as exc:
        cs.errors.append(f"写入已完成，但验证异常: {type(exc).__name__}: {exc}")
        cs.verification = {"ok": False, "reason": f"验证异常: {type(exc).__name__}: {exc}"}
        cs.status = "applied_with_warnings"
    finally:
        cs.save()
        _write_run_log(cs, note, result)
    return cs


def verify(cs: ChangeSet) -> dict:
    """重开输出文件，逐项核对预期改动是否落盘。"""
    if not cs.output_file or not Path(cs.output_file).exists():
        return {"ok": False, "reason": "输出文件不存在"}
    after = com_snapshot(cs.output_file)
    index = _task_index(after)
    checks = []
    for item in cs.preview:
        operation = item.get("op")
        if operation == "set_predecessors":
            row_id = item.get("row_id")
            task = next((t for t in after["tasks"] if t.get("id") == row_id), None)
            actual = task.get("predecessors") if task else None
            expected = item["changes"]["Predecessors"]["proposed"]
            ok = task is not None and str(actual) == str(expected)
            checks.append({"row_id": row_id, "field": "Predecessors", "ok": ok,
                           "expected": expected, "actual": actual})
            continue
        if operation == "create_milestone":
            name = item.get("new_task")
            expected_date = str(item.get("milestone_date"))[:10]
            matches = [t for t in after["tasks"] if t.get("name") == name and t.get("milestone")]
            task = next((t for t in matches if expected_date in str(t.get("start"))), None)
            ok = task is not None
            checks.append({"unique_id": task.get("unique_id") if task else None,
                           "field": "Milestone", "ok": ok, "expected_name": name,
                           "expected_date": expected_date,
                           "actual_start": task.get("start") if task else None})
            continue
        uid = item.get("unique_id")
        if uid is None:
            checks.append({"field": "-", "ok": False, "reason": "预览项缺少验证标识"})
            continue
        task = index.get(uid)
        if not task:
            checks.append({"unique_id": uid, "field": "-", "ok": False, "reason": "任务丢失"})
            continue
        for field_name, change in item["changes"].items():
            key = FIELD_KEYS.get(field_name)
            if field_name == "Notes":
                actual = task.get("notes") or ""
                expected = str(change["proposed"])
                ok = (expected.strip().replace('\r\n', '\n').replace('\r', '\n') ==
                      actual.strip().replace('\r\n', '\n').replace('\r', '\n'))
                checks.append({"unique_id": uid, "field": "Notes", "ok": ok,
                               "expected_prefix": expected.strip()[:30], "actual_prefix": actual.strip()[:30]})
                continue
            actual = task.get(key) if key else None
            proposed = change["proposed"]
            if field_name == "PercentComplete":
                # Project 会重算完成率，允许 ±3 个百分点
                try:
                    ok = actual is not None and abs(float(actual) - float(proposed)) <= 3
                except (TypeError, ValueError):
                    ok = False
            elif field_name in {"Start", "Finish", "ActualStart", "ActualFinish", "ConstraintDate"}:
                expected_date = str(proposed).replace("T", " ")
                actual_date = str(actual).replace("T", " ")
                length = 16 if len(expected_date) > 10 else 10
                ok = actual is not None and actual_date[:length] == expected_date[:length]
            else:
                ok = str(actual) == str(proposed)
            checks.append({"unique_id": uid, "field": field_name, "ok": bool(ok),
                           "expected": proposed, "actual": actual})
    ok = all(c["ok"] for c in checks) and not cs.errors

    # 视图层保留性核对：拿备份（写入前）与结果文件比内部视图/表格流
    view_layer = {"checked": False}
    if cs.backup_file and Path(cs.backup_file).exists() and cs.output_file:
        view_layer = compare_view_streams(cs.backup_file, cs.output_file)
        view_layer["checked"] = True
        view_layer["mode"] = cs.save_mode
        if view_layer.get("preserved") is False:
            ok = False
            view_layer["warning"] = (
                "视图层未被完整保留（视图/表格内部流发生变化）。"
                "若为 new_revision（另存为新文件）模式，属已知缺陷：会丢失甘特图上的图形/文本对象与 OLE 链接。")

    return {"ok": ok, "checks": checks, "checked_at": time.strftime("%Y-%m-%d %H:%M:%S"),
            "source_sha256_before": cs.source_sha256,
            "source_sha256_after": sha256(cs.source_file),
            "backup_available": bool(cs.backup_file and Path(cs.backup_file).exists()),
            "view_layer": view_layer}


def _write_run_log(cs: ChangeSet, note: str, result: dict) -> None:
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    path = LOG_DIR / f"mcp_changeset_{cs.id}.jsonl"
    view = ((cs.verification or {}).get("view_layer") or {})
    with open(path, "a", encoding="utf-8") as fh:
        fh.write(json.dumps({
            "change_set": cs.id, "title": cs.title, "risk": cs.risk_level,
            "approved_by": cs.approved_by, "note": note,
            "save_mode": cs.save_mode,
            "source": cs.source_file, "source_sha256_before": cs.source_sha256,
            "output": cs.output_file, "backup": cs.backup_file,
            "verification_ok": (cs.verification or {}).get("ok"),
            "view_layer_preserved": view.get("preserved"),
            "com_exit_code": result.get("exit_code"),
        }, ensure_ascii=False) + "\n")


def rollback(change_set_id: str) -> dict:
    """回滚：用备份还原源文件，并清理本次产生的输出文件（若有）。"""
    cs = ChangeSet.load(change_set_id)
    if cs.status not in ("verified", "applied_with_warnings"):
        raise RuntimeError(f"状态 {cs.status} 不支持自动回滚")
    if not cs.output_sha256:
        raise RuntimeError("历史变更集没有写入后哈希，无法安全自动回滚，请人工核对备份")
    if not cs.output_file or sha256(cs.output_file) != cs.output_sha256:
        raise RuntimeError("输出文件在写入后又发生变化，拒绝回滚覆盖后续工作")
    removed = []
    restored = None
    if cs.save_mode == "in_place":
        # 就地保存模式下源文件已被覆盖 —— 从备份还原才是真正的回滚
        if cs.backup_file and Path(cs.backup_file).exists():
            import shutil
            if sha256(cs.backup_file) != cs.source_sha256:
                raise RuntimeError("备份哈希与写入前不一致，拒绝回滚")
            with exclusive_project():
                shutil.copy2(cs.backup_file, cs.source_file)
            restored = cs.source_file
    elif cs.output_file and Path(cs.output_file).exists():
        with exclusive_project():
            Path(cs.output_file).unlink()
        removed.append(cs.output_file)
    cs.status = "rolled_back"
    cs.save()
    intact = restored is not None and sha256(cs.source_file) == cs.source_sha256
    return {"change_set": cs.id, "removed": removed, "restored_from_backup": restored,
            "source_matches_pre_write_hash": intact,
            "backup": cs.backup_file}
