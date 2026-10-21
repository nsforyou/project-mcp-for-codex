"""Cross-process serialization and fail-closed protection for Project COM."""
from contextlib import contextmanager
from pathlib import Path
import msvcrt
import os
import psutil

LOCK = Path(os.environ.get("PROJECT_COM_LOCK", str(Path(__file__).resolve().parent / "work" / "project-com.lock")))


def project_processes():
    return [p.info for p in psutil.process_iter(["pid", "name"])
            if (p.info["name"] or "").lower() == "winproj.exe"]


@contextmanager
def project_lock():
    LOCK.parent.mkdir(parents=True, exist_ok=True)
    with LOCK.open("a+b") as handle:
        if handle.tell() == 0:
            handle.write(b"0")
            handle.flush()
        handle.seek(0)
        try:
            msvcrt.locking(handle.fileno(), msvcrt.LK_NBLCK, 1)
        except OSError as exc:
            raise RuntimeError("另一项 Project COM 操作正在运行，请等待完成") from exc
        try:
            yield
        finally:
            handle.seek(0)
            msvcrt.locking(handle.fileno(), msvcrt.LK_UNLCK, 1)


@contextmanager
def exclusive_project():
    with project_lock():
        running = project_processes()
        if running:
            raise RuntimeError(
                "检测到 Microsoft Project 正在运行。请保存并关闭 Project 后执行离线 CLI/MCP；"
                f"在线协作请使用 project_live_* 工具。本次未启动 COM、未修改文件。进程: {running}")
        yield
