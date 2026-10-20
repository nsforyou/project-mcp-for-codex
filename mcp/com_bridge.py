# -*- coding: utf-8 -*-
r"""COM 桥接层：MCP Server 通过它调用已实测的 PowerShell COM 脚本。

设计要点
    - 既有 MPP 写入通过 com_modify.ps1，默认备份后就地保存以保留视图层
    - 读取：MPXJ（快，不启动 Project）或 COM（可读视图层）
    - 子进程一律不经过管道捕获之外的额外包装，输出统一为 JSON 文件 + stdout
"""
from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
import tempfile
import time
from com_runtime import exclusive_project
from dataclasses import dataclass
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SCRIPTS = ROOT / "scripts"
WORK = ROOT / "mcp" / "work"
WORK.mkdir(parents=True, exist_ok=True)

POWERSHELL = os.environ.get("MPXJ_POWERSHELL", "powershell.exe")
MPXJ_LIB = os.environ.get("MPXJ_LIB", str(Path(os.environ.get('LOCALAPPDATA', str(Path.home()))) / 'ProjectMCP' / 'deps'))
COM_TIMEOUT = int(os.environ.get("PROJECT_COM_TIMEOUT", "900"))


class BridgeError(RuntimeError):
    pass


def _run_ps(script: str, args: list[str], timeout: int = COM_TIMEOUT) -> tuple[int, str]:
    cmd = [POWERSHELL, "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(SCRIPTS / script), *args]
    # Codex can prepend its bundled PowerShell 7 modules to PSModulePath.  The
    # scripts are deliberately run by Windows PowerShell 5.1, whose built-in
    # modules (including Microsoft.PowerShell.Utility/Get-FileHash) must win.
    env = dict(os.environ)
    system_root = Path(env.get("SystemRoot", r"C:\Windows"))
    ps_modules = [
        Path(env.get("ProgramFiles", r"C:\Program Files")) / "WindowsPowerShell" / "Modules",
        system_root / "System32" / "WindowsPowerShell" / "v1.0" / "Modules",
    ]
    env["PSModulePath"] = ";".join(str(path) for path in ps_modules if path.exists())
    with exclusive_project():
        proc = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8",
                              errors="replace", timeout=timeout, env=env, cwd=str(ROOT),
                              creationflags=subprocess.CREATE_NO_WINDOW)
    out = (proc.stdout or "") + (proc.stderr or "")
    return proc.returncode, out


def _run_py(args: list[str], timeout: int = 600) -> tuple[int, str]:
    env = dict(os.environ)
    env["MPXJ_LIB"] = MPXJ_LIB
    proc = subprocess.run([sys.executable, *args], capture_output=True, text=True,
                          encoding="utf-8", errors="replace", timeout=timeout, env=env, cwd=str(ROOT))
    return proc.returncode, (proc.stdout or "") + (proc.stderr or "")


# ---------------------------------------------------------------- 读取（COM）
def com_snapshot(mpp_path: str) -> dict:
    """只读采集 COM 指纹（ReadOnly=true 打开，pjDoNotSave 关闭，实测无副作用）。"""
    mpp = Path(mpp_path)
    if not mpp.exists():
        raise BridgeError(f"文件不存在: {mpp}")
    out = WORK / f"snapshot_{int(time.time()*1000)}.json"
    code, log = _run_ps("com_snapshot.ps1", ["-MppPath", str(mpp), "-Out", str(out)])
    if code != 0 or not out.exists():
        raise BridgeError(f"COM 只读采集失败(code={code}): {log[-800:]}")
    data = json.loads(out.read_text(encoding="utf-8"))
    out.unlink(missing_ok=True)
    return data


# ---------------------------------------------------------------- 读取（MPXJ）
def mpxj_read(mpp_path: str) -> dict:
    """MPXJ 结构化读取（不启动 Project，快）。"""
    mpp = Path(mpp_path)
    if not mpp.exists():
        raise BridgeError(f"文件不存在: {mpp}")
    out = WORK / f"mpxj_{int(time.time()*1000)}.json"
    code, log = _run_py(["scripts/mpp_probe.py", str(mpp), "--json", str(out), "--no-markdown"])
    if code != 0 or not out.exists():
        raise BridgeError(f"MPXJ 读取失败(code={code}): {log[-800:]}")
    data = json.loads(out.read_text(encoding="utf-8"))
    out.unlink(missing_ok=True)
    return data


# ---------------------------------------------------------------- 写入（COM）
def com_apply(spec: dict, keep_visible: bool = False) -> dict:
    """执行写入：spec 见 scripts/com_modify.ps1 说明。

    save_in_place=true 时必须传 -NoWorkingCopy：否则 com_modify 会在临时副本上做就地保存，
    改动落在副本上、目标文件纹丝不动（这类"空操作"必须靠验证环节发现）。
    spec 里已经带了源文件的绝对路径，调用方（change_set.apply）负责先备份。

    返回操作日志（含 before/after 快照、源文件 hash、保存结果）。
    """
    spec_path = WORK / f"spec_{int(time.time()*1000)}.json"
    spec_path.write_text(json.dumps(spec, ensure_ascii=False, indent=2), encoding="utf-8")
    args = ["-OpsJson", str(spec_path)]
    if keep_visible:
        args.append("-KeepVisible")
    if spec.get("save_in_place"):
        args.append("-NoWorkingCopy")
    code, log = _run_ps("com_modify.ps1", args)
    log_path = Path(spec.get("log", ""))
    result: dict = {"exit_code": code, "stdout": log[-4000:]}
    if log_path.exists():
        result["log"] = json.loads(log_path.read_text(encoding="utf-8"))
    spec_path.unlink(missing_ok=True)
    if code != 0:
        raise BridgeError(f"COM 写入失败(code={code}): {log[-1200:]}")
    return result


def com_recalculate_and_verify(mpp_path: str) -> dict:
    """重开文件采集指纹，用于写入后的复核。"""
    return com_snapshot(mpp_path)


def create_plan(source: str, output: str, *, source_sha256: str = "", preview: bool = True) -> dict:
    """Preview or create a fresh MPP. Existing MPP editing stays in Change Sets."""
    from plan_builder import inspect_xml
    info = inspect_xml(source)
    destination = Path(output).resolve()
    if destination.suffix.lower() != ".mpp" or destination.exists():
        raise BridgeError("输出必须是尚不存在的 .mpp 文件")
    digest = sha256(source)
    if preview:
        return {**info, "output": str(destination), "source_sha256": digest, "preview": True}
    if source_sha256 != digest:
        raise BridgeError("请先预览并传入匹配的 source_sha256；源文件可能已经改变")
    cmd = [sys.executable, str(ROOT / "mcp" / "plan_builder.py"), str(Path(source).resolve()), str(destination)]
    env = {**os.environ, "MPXJ_LIB": MPXJ_LIB, "PYTHONIOENCODING": "utf-8", "PROJECT_COM_PARENT": str(os.getpid())}
    with exclusive_project():
        try:
            proc = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace",
                                  timeout=COM_TIMEOUT, env=env, cwd=str(ROOT),
                                  creationflags=subprocess.CREATE_NO_WINDOW)
        except subprocess.TimeoutExpired as exc:
            raise BridgeError("创建超时；未报告成功。请检查后台 Project 进程和输出文件后重试。") from exc
    try:
        result = json.loads(proc.stdout.strip().splitlines()[-1])
    except (ValueError, IndexError) as exc:
        raise BridgeError(f"创建失败，exit={proc.returncode}: {proc.stderr[-2000:]}") from exc
    result["source_sha256"] = digest
    result["verified"] = proc.returncode == 0 and result.get("verification", {}).get("ok") is True
    log_path = destination.with_suffix(".creation.json")
    log_path.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    result["log"] = str(log_path)
    if not result["verified"]:
        raise BridgeError(f"文件已生成但核验失败，不可交付。详见 {log_path}")
    return result


# ---------------------------------------------------------------- 导出 / 分析
def export_excel(json_path: str, out_path: str) -> str:
    code, log = _run_py(["scripts/export_excel.py", json_path, out_path])
    if code != 0:
        raise BridgeError(f"Excel 导出失败: {log[-800:]}")
    return out_path


def diff_plans(before_json: str, after_json: str) -> str:
    code, log = _run_py(["scripts/diff_plans.py", before_json, after_json])
    return log


def backup_file(path: str, backup_dir: str | None = None) -> str:
    src = Path(path)
    dest_dir = Path(backup_dir) if backup_dir else ROOT / "backups"
    dest_dir.mkdir(parents=True, exist_ok=True)
    stamp = time.strftime("%Y%m%d_%H%M%S")
    dest = dest_dir / f"{src.stem}_MCP备份_{stamp}{src.suffix}"
    shutil.copy2(src, dest)
    return str(dest)


def sha256(path: str) -> str:
    import hashlib
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


# ---------------------------------------------------------------- 视图层取证
# CV_iew 是视图/图形/文本对象的载体：**必须逐字节保留**，任何变化都视为丢失。
# CTable / CEdl 在写入业务数据（备注等）时会正常增长，因此只把"缩小"视为可疑。
CRITICAL_STREAMS = ("214/CV_iew/Var2Data", "214/CV_iew/VarMeta", "214/CV_iew/FixedData")
GROWABLE_STREAMS = ("214/CTable/Var2Data", "214/CTable/VarMeta", "214/CTable/FixedData",
                    "214/CEdl/Var2Data", "214/CEdl/VarMeta")
VIEW_STREAMS = CRITICAL_STREAMS + GROWABLE_STREAMS


def _stream_hashes(path: str) -> dict[str, tuple[int, str]]:
    """读取 MPP 内部流的 (字节数, sha256)。

    注意：MPP 存储名带前导空格（如 '   214/CV_iew/Var2Data'），必须用原始名读取。
    """
    import hashlib as _hashlib
    try:
        import olefile
    except ImportError:  # olefile 缺失时降级为"无法取证"，不阻塞主流程
        return {}
    out: dict[str, tuple[int, str]] = {}
    try:
        ole = olefile.OleFileIO(str(path))
    except Exception:
        return {}
    try:
        for entry in ole.listdir(streams=True, storages=False):
            raw = "/".join(entry)
            key = raw.strip()
            if key not in VIEW_STREAMS:
                continue
            try:
                data = ole.openstream(raw).read()
            except Exception:
                continue
            out[key] = (len(data), _hashlib.sha256(data).hexdigest())
    finally:
        ole.close()
    return out


def compare_view_streams(before_path: str, after_path: str) -> dict:
    """对比两文件的视图/表格/链接存储，判断视图层是否受损。

    判定规则（用就地保存与另存为两种模式的实测结果校准）：
      - CV_iew（视图/图形/文本对象载体）：**必须逐字节一致**，任何变化都算丢失
      - CTable / CEdl：写入业务数据（备注等）时会正常增长，只有**缩小**才可疑
      - 另存为（SaveAs）的典型特征是 CV_iew 主体从 1.8MB 级骤降到 170KB 级
    """
    a = _stream_hashes(before_path)
    b = _stream_hashes(after_path)
    if not a or not b:
        return {"preserved": None, "reason": "无法读取流（olefile 缺失或文件非复合文档）"}

    details = []
    critical_lost = []
    suspicious_shrink = []
    for key in sorted(set(a) | set(b)):
        ha, hb = a.get(key), b.get(key)
        if not ha or not hb:
            critical_lost.append(key)
            details.append({"stream": key, "before_bytes": ha[0] if ha else None,
                            "after_bytes": hb[0] if hb else None, "identical": False,
                            "verdict": "流缺失"})
            continue
        identical = ha[1] == hb[1]
        verdict = "一致"
        if not identical:
            if key in CRITICAL_STREAMS:
                verdict = "丢失风险（视图载体被改写）"
                critical_lost.append(key)
            elif hb[0] < ha[0]:
                verdict = "可疑缩小"
                suspicious_shrink.append(key)
            else:
                verdict = f"正常增长 +{hb[0] - ha[0]} 字节"
        details.append({"stream": key, "before_bytes": ha[0], "after_bytes": hb[0],
                        "identical": identical, "verdict": verdict})

    preserved = not critical_lost and not suspicious_shrink
    result = {"preserved": preserved, "streams": details,
              "critical_lost": critical_lost, "suspicious_shrink": suspicious_shrink,
              "rule": "CV_iew 必须逐字节一致；CTable/CEdl 允许正常增长，缩小即可疑"}
    if not preserved:
        result["warning"] = (
            "视图层疑似受损：" + ("；".join(critical_lost) if critical_lost else "") +
            ("；".join(suspicious_shrink) if suspicious_shrink else "") +
            "。若为另存为新文件（new_revision）模式，属已知缺陷：会丢失甘特图上的图形/文本对象与 OLE 链接。")
    return result
