"""Test live commands against a test-owned visible Project, never user files."""
import json
from pathlib import Path
import shutil
import sys
from datetime import datetime

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from com_runtime import exclusive_project
from live_project import run_live, file_hash
from com_bridge import create_plan


def main():
    import pythoncom
    import win32com.client
    root = Path(__file__).resolve().parents[2]
    target = root / 'mcp' / 'work' / f'在线协作测试_{datetime.now():%Y%m%d_%H%M%S}.mpp'
    source = target.with_suffix('.seed.mpp')
    fixture = Path(__file__).with_name('basic_plan.xml')
    plan_preview = create_plan(str(fixture), str(source))
    create_plan(str(fixture), str(source), preview=False, source_sha256=plan_preview['source_sha256'])
    app = None
    pythoncom.CoInitialize()
    try:
        with exclusive_project():
            shutil.copy2(source, target)
            app = win32com.client.DispatchEx('MSProject.Application')
            app.FileOpenEx(str(target), False)
            app.Visible = True
            app.FileSave()
        stale = run_live('preview', {'file': str(target), 'ops': [
            {'unique_id': 1, 'fields': {'Name': 'A_stale'}}]})
        app.ActiveProject.Tasks.UniqueID(2).Name = 'B_manual'
        try:
            run_live('apply', {'change_id': stale['change_id'], 'approved_by': '自动测试'})
            raise AssertionError('stale preview should have failed')
        except RuntimeError as exc:
            assert '发生变化' in str(exc), exc
        assert app.ActiveProject.Tasks.UniqueID(2).Name == 'B_manual'
        app.ActiveProject.Tasks.UniqueID(2).Name = 'B'
        app.FileSave()
        before = run_live('read', {'file': str(target)})
        disk_before = file_hash(target)
        preview = run_live('preview', {'file': str(target), 'ops': [
            {'unique_id': 1, 'fields': {'Name': 'A_test', 'ResourceNames': 'BB'},
             'note_fields': {'工程量': 120, '单位': '米', '当前状态': '接线中'}},
            {'unique_id': 2, 'fields': {'Start': '2026-11-04T08:00', 'Finish': '2026-11-06T17:00',
                                      'PercentComplete': 25}},
        ]})
        assert file_hash(target) == disk_before
        result = run_live('apply', {'change_id': preview['change_id'],
                                   'approved_by': '用户授权在线协作开发；仅写入自动创建的测试副本'})
        assert result['status'] == 'memory_verified', result
        after = run_live('read', {'file': str(target)})
        assert after['pid'] == before['pid']
        assert not after['saved']
        assert file_hash(target) == disk_before, 'live apply must not auto-save'
        assert app.Projects.Count == 1 and app.Visible
        # One native Undo restores the entire confirmed batch.
        app.Undo(1)
        restored = run_live('read', {'file': str(target)})
        assert restored['tasks'] == before['tasks'], 'Undo did not restore task data'
        report = {'ok': True, 'target': str(target), 'change_id': preview['change_id'],
                  'checks': result['checks'], 'same_process': True, 'window_kept_open': True,
                  'disk_unchanged_until_user_save': True, 'single_undo_restored_tasks': True}
        report['stale_preview_rejected_without_overwriting_user_edit'] = True
        target.with_suffix('.live-test.json').write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
        print(json.dumps(report, ensure_ascii=False))
    finally:
        # Only the test harness owns this fresh instance, never the live code.
        if app is not None:
            app.FileCloseEx(0)
            app.Quit(0)
        app = None
        pythoncom.CoUninitialize()


if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    main()
