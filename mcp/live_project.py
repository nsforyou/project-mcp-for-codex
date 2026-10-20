"""Opt-in live Project commands. Borrow the running application; never own it.

No FileOpen, FileClose, Quit, Visible, DisplayAlerts, selection or keyboard
operations are used here. Work is bounded in a separate worker process.
"""
from __future__ import annotations

import argparse
import gc
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import time
import uuid

from com_runtime import project_lock, project_processes
from note_fields import merge_fields

ROOT = Path(__file__).resolve().parent.parent
LIVE_DIR = ROOT / 'mcp' / 'changesets' / 'live'
FIELDS = {'Name', 'Notes', 'Start', 'Finish', 'Duration', 'PercentComplete', 'Predecessors', 'ResourceNames'}
READ_FIELDS = ('UniqueID', 'ID', 'Name', 'OutlineLevel', 'Summary', 'Start', 'Finish', 'Duration',
               'PercentComplete', 'Predecessors', 'ResourceNames', 'Notes', 'Manual')


def canonical(path):
    return os.path.normcase(os.path.realpath(str(path)))


def digest(value):
    return hashlib.sha256(json.dumps(value, sort_keys=True, ensure_ascii=False, default=str).encode('utf-8')).hexdigest()


def file_hash(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def _gui_info(thread_id):
    import ctypes
    from ctypes import wintypes
    class GUIINFO(ctypes.Structure):
        _fields_ = [('cbSize', wintypes.DWORD), ('flags', wintypes.DWORD),
                    ('hwndActive', wintypes.HWND), ('hwndFocus', wintypes.HWND),
                    ('hwndCapture', wintypes.HWND), ('hwndMenuOwner', wintypes.HWND),
                    ('hwndMoveSize', wintypes.HWND), ('hwndCaret', wintypes.HWND),
                    ('rcCaret', wintypes.RECT)]
    info = GUIINFO()
    info.cbSize = ctypes.sizeof(info)
    user32 = ctypes.WinDLL('user32', use_last_error=True)
    user32.GetGUIThreadInfo.argtypes = [wintypes.DWORD, ctypes.POINTER(GUIINFO)]
    user32.GetGUIThreadInfo.restype = wintypes.BOOL
    if not user32.GetGUIThreadInfo(thread_id, ctypes.byref(info)):
        raise ctypes.WinError(ctypes.get_last_error())
    return {'flags': info.flags, 'hwndFocus': info.hwndFocus, 'hwndCaret': info.hwndCaret}


def _ui_idle(pid):
    """Conservatively reject disabled/modal windows and a visible edit caret."""
    import win32gui
    import win32process
    windows = []
    win32gui.EnumWindows(lambda hwnd, _: windows.append(hwnd) if
        win32process.GetWindowThreadProcessId(hwnd)[1] == pid and win32gui.IsWindowVisible(hwnd)
        and not win32gui.GetWindow(hwnd, 4) else None, None)
    for hwnd in windows:
        if not win32gui.IsWindowEnabled(hwnd):
            raise RuntimeError('Project 有模态对话框，请先关闭对话框')
    foreground = win32gui.GetForegroundWindow()
    if foreground and win32process.GetWindowThreadProcessId(foreground)[1] == pid:
        info = _gui_info(win32process.GetWindowThreadProcessId(foreground)[0])
        if info.get('flags', 0) & 0x1e:
            raise RuntimeError('Project 正在使用菜单或移动窗口，请稍后执行')
        focus = info.get('hwndFocus')
        if info.get('hwndCaret') or (focus and 'edit' in win32gui.GetClassName(focus).lower()):
            raise RuntimeError('Project 正在编辑单元格/备注，请先按 Enter 完成编辑，再执行')


def attach(target, expected_pid=None):
    import win32com.client
    processes = project_processes()
    if len(processes) != 1:
        raise RuntimeError('在线协作要求恰好一个 Project 进程；不会自动启动或关闭 Project')
    pid = processes[0]['pid']
    if expected_pid is not None and pid != expected_pid:
        raise RuntimeError('Project 进程已经重启，请重新预览')
    _ui_idle(pid)
    app = win32com.client.GetActiveObject('MSProject.Application')
    project = app.ActiveProject
    if project is None or canonical(project.FullName) != canonical(target):
        raise RuntimeError('当前活动项目与指定文件不一致，请切换到目标项目后重试')
    return app, project, pid


def snapshot(project, pid):
    if project.Tasks.Count > 500:
        raise ValueError('在线预览当前限 500 个任务，请使用离线流程处理更大计划')
    tasks = []
    for task in project.Tasks:
        if task is None:
            continue
        row = {}
        for name in READ_FIELDS:
            value = getattr(task, name)
            if hasattr(value, 'strftime'):
                value = value.strftime('%Y-%m-%d %H:%M')
            row[name] = value
        tasks.append(row)
    resources = [{'UniqueID': r.UniqueID, 'Name': r.Name, 'MaxUnits': r.MaxUnits} for r in project.Resources if r is not None]
    data = {'file': canonical(project.FullName), 'pid': pid, 'saved': bool(project.Saved),
            'tasks': tasks, 'resources': resources}
    data['fingerprint'] = digest(data)
    return data


def validate_ops(ops, snap):
    if not isinstance(ops, list) or not 1 <= len(ops) <= 20:
        raise ValueError('一次在线修改须为 1~20 项；大量重建请使用离线流程')
    index = {t['UniqueID']: t for t in snap['tasks']}
    proposed = []
    for op in ops:
        uid = int(op['unique_id'])
        if uid not in index:
            raise ValueError(f'找不到任务 UID={uid}')
        fields = dict(op.get('fields') or {})
        if op.get('note_fields') is not None:
            if 'Notes' in fields:
                raise ValueError('不能同时传 Notes 和 note_fields')
            fields['Notes'] = merge_fields(index[uid]['Notes'] or '', op['note_fields'])
        if not fields or set(fields) - FIELDS:
            raise ValueError(f'不支持的在线字段: {set(fields) - FIELDS}')
        if index[uid]['Summary'] and set(fields) - {'Name', 'Notes'}:
            raise ValueError('摘要任务仅允许在线修改名称和备注')
        if 'PercentComplete' in fields and not 0 <= float(fields['PercentComplete']) <= 100:
            raise ValueError('完成率须在 0~100 之间')
        if 'Duration' in fields and not isinstance(fields['Duration'], (int, float)):
            raise ValueError('在线 Duration 须填写工作分钟数')
        for key in ('Start', 'Finish'):
            if key in fields:
                from datetime import datetime
                fields[key] = datetime.fromisoformat(str(fields[key])).strftime('%Y-%m-%d %H:%M')
        if any(value is None or isinstance(value, (dict, list)) for value in fields.values()):
            raise ValueError('字段值必须为文字、数字或布尔值')
        proposed.append({'unique_id': uid, 'name': index[uid]['Name'], 'fields': fields,
                         'before': {key: index[uid][key] for key in fields}})
    pairs = [(op['unique_id'], key) for op in proposed for key in op['fields']]
    if len(pairs) != len(set(pairs)):
        raise ValueError('同一字段存在重复修改，请合并后预览')
    return proposed


def load_change(change_id):
    if not isinstance(change_id, str) or not change_id.startswith('LIVE-') or any(c not in 'LIVE-0123456789abcdef' for c in change_id):
        raise ValueError('非法在线变更 ID')
    return LIVE_DIR / f'{change_id}.json'


def worker(action, args):
    import pythoncom
    pythoncom.CoInitialize()
    app = project = None
    try:
        if action == 'apply':
            record_path = load_change(args['change_id'])
            record = json.loads(record_path.read_text(encoding='utf-8'))
            if record['status'] != 'preview' or not str(args.get('approved_by', '')).strip():
                raise ValueError('需要尚未执行的在线预览和明确的 approved_by')
            target = record['snapshot']['file']
            app, project, pid = attach(target, record['snapshot']['pid'])
        else:
            target = args['file']
            app, project, pid = attach(target)
        snap = snapshot(project, pid)
        if action == 'read':
            return snap
        if action == 'preview':
            if not snap['saved']:
                raise RuntimeError('请先在 Project 中保存当前改动（无需关闭），再生成可执行预览')
            if not Path(target).is_file():
                raise ValueError('请先将项目保存为本地 MPP')
            proposed = validate_ops(args['ops'], snap)
            change_id = 'LIVE-' + uuid.uuid4().hex
            record = {'id': change_id, 'title': args.get('title') or '在线计划调整',
                      'status': 'preview', 'created_at': time.time(), 'snapshot': snap,
                      'disk_hash': file_hash(target), 'proposed': proposed}
            LIVE_DIR.mkdir(parents=True, exist_ok=True)
            load_change(change_id).write_text(json.dumps(record, ensure_ascii=False, indent=2), encoding='utf-8')
            return {'change_id': change_id, 'file': snap['file'], 'preview': proposed,
                    'note': '尚未修改。执行后保留窗口，默认不保存，由用户查看后 Ctrl+S 或撤销。'}
        if snap['fingerprint'] != record['snapshot']['fingerprint'] or file_hash(target) != record['disk_hash']:
            raise RuntimeError('计划在预览后发生变化，请重新预览；本次未写入')
        backup = ROOT / 'backups' / f'{Path(target).stem}_{record["id"]}.mpp'
        backup.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(target, backup)
        if file_hash(backup) != record['disk_hash']:
            raise RuntimeError('备份与预览不一致，拒绝写入')
        record.update(status='applying', backup=str(backup), approved_by=args['approved_by'])
        record_path.write_text(json.dumps(record, ensure_ascii=False, indent=2), encoding='utf-8')
        opened_undo = False
        try:
            _ui_idle(pid)
            if canonical(app.ActiveProject.FullName) != canonical(target):
                raise RuntimeError('活动项目已切换，停止执行')
            app.OpenUndoTransaction('MCP 在线调整 ' + record['id'])
            opened_undo = True
            for op in record['proposed']:
                if canonical(app.ActiveProject.FullName) != canonical(target):
                    raise RuntimeError('执行期间活动项目发生切换')
                task = project.Tasks.UniqueID(op['unique_id'])
                for key, value in op['fields'].items():
                    setattr(task, key, value)
            app.CloseUndoTransaction()
            opened_undo = False
            after = snapshot(project, pid)
            index = {t['UniqueID']: t for t in after['tasks']}
            checks = []
            for op in record['proposed']:
                for key, expected in op['fields'].items():
                    actual = index[op['unique_id']][key]
                    normalize = lambda v: str(v).replace('\r\n', '\n').replace('\r', '\n').strip()
                    matches = (float(expected) == float(actual) if key in ('Duration', 'PercentComplete')
                               else normalize(expected) == normalize(actual))
                    checks.append({'unique_id': op['unique_id'], 'field': key,
                                   'expected': expected, 'actual': actual,
                                   'ok': matches})
            requested = {(op['unique_id'], key) for op in record['proposed'] for key in op['fields']}
            prior = {t['UniqueID']: t for t in snap['tasks']}
            other_changes = [
                {'unique_id': uid, 'name': row['Name'], 'field': key,
                 'before': prior[uid][key], 'after': value}
                for uid, row in index.items() if uid in prior
                for key, value in row.items()
                if (uid, key) not in requested and prior[uid].get(key) != value
            ]
            record.update(status='memory_verified' if all(c['ok'] for c in checks) else 'needs_review',
                          checks=checks, after=after, other_changes=other_changes, saved=False)
            return {'change_id': record['id'], 'status': record['status'], 'checks': checks,
                    'backup': str(backup), 'project_left_open': True, 'saved': False,
                    'other_changes': other_changes,
                    'note': '改动已显示在 Project。请查看后 Ctrl+S 保存；可使用撤销列表中的 MCP 在线调整撤销。'}
        except Exception as exc:
            record.update(status='needs_review', error=str(exc))
            raise
        finally:
            if opened_undo:
                try:
                    app.CloseUndoTransaction()
                except Exception:
                    pass
            record_path.write_text(json.dumps(record, ensure_ascii=False, indent=2), encoding='utf-8')
    finally:
        # Borrowed application: release Python references only. No Quit/Close.
        project = app = None
        gc.collect()
        pythoncom.CoUninitialize()


def run_live(action, args):
    if action not in ('read', 'preview', 'apply'):
        raise ValueError('非法在线命令')
    with project_lock():
        proc = subprocess.Popen([sys.executable, str(Path(__file__).resolve()), '--worker', action],
            stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
            text=True, encoding='utf-8', creationflags=subprocess.CREATE_NO_WINDOW)
        try:
            stdout, stderr = proc.communicate(json.dumps(args, ensure_ascii=False), timeout=20)
        except subprocess.TimeoutExpired:
            proc.kill()
            proc.communicate()
            if action == 'apply':
                path = load_change(args['change_id'])
                if path.exists():
                    record = json.loads(path.read_text(encoding='utf-8'))
                    record.update(status='needs_review', error='工作进程超时；可能部分写入，请检查 Project 和撤销记录')
                    path.write_text(json.dumps(record, ensure_ascii=False, indent=2), encoding='utf-8')
            raise RuntimeError('Project 在线命令超时，已停止工作进程；不会自动重试或关闭 Project。写入结果须人工核对。')
    if proc.returncode:
        raise RuntimeError(stderr.strip() or stdout.strip() or '在线命令失败')
    return json.loads(stdout)


if __name__ == '__main__':
    for stream in (sys.stdin, sys.stdout, sys.stderr):
        stream.reconfigure(encoding='utf-8')
    parser = argparse.ArgumentParser()
    parser.add_argument('--worker', choices=('read', 'preview', 'apply'), required=True)
    command = parser.parse_args()
    try:
        print(json.dumps(worker(command.worker, json.load(sys.stdin)), ensure_ascii=False))
    except Exception as exc:
        print(f'{type(exc).__name__}: {exc}', file=sys.stderr)
        raise SystemExit(1)
