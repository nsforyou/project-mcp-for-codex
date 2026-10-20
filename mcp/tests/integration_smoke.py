"""Actual Project round trip on a new disposable plan only."""
import json
import sys
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from com_bridge import create_plan, sha256
import change_set as CS
from note_fields import merge_fields
from project_mcp_server import t_get_note_fields


def main():
    source = Path(__file__).with_name('basic_plan.xml')
    output = source.parents[1] / 'work' / f'cli_integration_{datetime.now():%Y%m%d_%H%M%S}.mpp'
    preview = create_plan(str(source), str(output))
    created = create_plan(str(source), str(output), preview=False, source_sha256=preview['source_sha256'])
    assert created['verified'], created
    before = sha256(str(output))
    cs = CS.build('CLI 创建/修改/回滚集成验证', str(output), [
        {'op': 'set_task_fields', 'unique_id': 2, 'fields': {'Name': 'B_test'}},
        {'op': 'set_task_fields', 'unique_id': 2, 'fields': {'Notes': merge_fields('保留的现场说明。' * 70,
            {'工程量': 120, '单位': '米', '当前状态': '接线中', '已完成量': 35})}},
    ], requested_by='integration-test')
    assert not cs.errors, cs.errors
    applied = CS.apply(cs.id, approved_by='用户已授权完善 MCP；仅测试新建样本', allow_in_place=True)
    assert all(c['ok'] for c in applied.verification['checks']), applied.to_dict()
    view_preserved = applied.verification['view_layer'].get('preserved')
    assert applied.status == ('verified' if view_preserved is not False else 'applied_with_warnings')
    note_fields = t_get_note_fields({'file': str(output), 'unique_id': 2})
    assert note_fields['fields']['工程量'] == '120', note_fields
    assert note_fields['fields']['当前状态'] == '接线中', note_fields
    restored = CS.rollback(cs.id)
    assert sha256(str(output)) == before, restored
    report = {'ok': True, 'created': str(output), 'change_set': cs.id,
              'create_checks': len(created['verification']['checks']),
              'write_checks': len(applied.verification['checks']),
              'view_layer_preserved': view_preserved, 'rollback_hash_matches': True,
              'note_fields_round_trip': True}
    path = output.with_suffix('.test-report.json')
    path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps(report, ensure_ascii=False))


if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    main()
