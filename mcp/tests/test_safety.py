import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from com_runtime import exclusive_project
from plan_builder import inspect_xml
from com_bridge import create_plan
from note_fields import read_fields, merge_fields
from project_mcp_server import t_get_note_fields

FIXTURE = Path(__file__).with_name("basic_plan.xml")


class SafetyTests(unittest.TestCase):
    def test_batch_note_fields_filters_state(self):
        snap = {'tasks': [
            {'unique_id': 1, 'name': 'A', 'notes': '工程量：120\r当前状态：接线中'},
            {'unique_id': 2, 'name': 'B', 'notes': '当前状态：未开始'},
        ]}
        with patch('project_mcp_server.active_or', return_value='sample.mpp'), patch('project_mcp_server.com_snapshot', return_value=snap):
            result = t_get_note_fields({'field_name': '当前状态', 'equals': '接线中'})
            self.assertEqual(result['count'], 1)
            self.assertEqual(result['tasks'][0]['fields']['工程量'], '120')

    def test_note_fields_preserve_prose_and_migrate_fields(self):
        notes = '保留的现场说明\n工程量：120\n单位：米\n当前状态：未开始'
        merged = merge_fields(notes, {'当前状态': '接线中', '已完成量': 35})
        self.assertIn('保留的现场说明', merged)
        self.assertNotIn('当前状态：未开始', merged)
        self.assertEqual(read_fields(merged)['fields']['当前状态'], '接线中')
        self.assertEqual(read_fields(merged)['fields']['已完成量'], '35')
        self.assertEqual(read_fields(merged)['fields'], read_fields(merge_fields(merged, {}))['fields'])
        self.assertEqual(read_fields(merged.replace('\r\n', '\r'))['fields']['工程量'], '120')

    def test_note_fields_reject_ambiguous_or_invalid_input(self):
        with self.assertRaises(ValueError):
            merge_fields('工程量：10\n工程量：20', {'工程量': 30})
        with self.assertRaises(ValueError):
            merge_fields('', {'工程量': -1})
        with self.assertRaises(ValueError):
            merge_fields('', {'当前状态': '一行\n另一行'})

    def test_running_project_is_never_automated(self):
        with patch('com_runtime.project_processes', return_value=[{'pid': 123, 'name': 'WINPROJ.EXE'}]):
            with self.assertRaisesRegex(RuntimeError, '正在运行'):
                with exclusive_project():
                    self.fail('must not enter')

    def test_preview_and_hash_guard(self):
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / 'new.mpp'
            result = create_plan(str(FIXTURE), str(output))
            self.assertEqual(result['tasks'], 2)
            self.assertFalse(output.exists())
            with self.assertRaisesRegex(Exception, 'source_sha256'):
                create_plan(str(FIXTURE), str(output), preview=False, source_sha256='wrong')

    def test_existing_output_is_protected(self):
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / 'existing.mpp'
            output.touch()
            with self.assertRaisesRegex(Exception, '尚不存在'):
                create_plan(str(FIXTURE), str(output))

    def test_cycle_rejected_before_com(self):
        text = FIXTURE.read_text(encoding='utf-8').replace('<Name>A</Name>', '<Name>A</Name><PredecessorLink><PredecessorUID>20</PredecessorUID><Type>1</Type></PredecessorLink>')
        with tempfile.TemporaryDirectory() as directory:
            xml = Path(directory) / 'cycle.xml'
            xml.write_text(text, encoding='utf-8')
            with self.assertRaisesRegex(ValueError, '成环'):
                inspect_xml(xml)


if __name__ == '__main__':
    unittest.main()
