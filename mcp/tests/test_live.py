import sys
import unittest
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from live_project import validate_ops, load_change, run_live, digest, _ui_idle, _gui_info


class LiveTests(unittest.TestCase):
    def setUp(self):
        self.snap = {'tasks': [{'UniqueID': 12, 'Name': '接线', 'Summary': False,
                               'Notes': '现场说明', 'PercentComplete': 0}], 'saved': True}

    def test_note_fields_are_preview_only(self):
        result = validate_ops([{'unique_id': 12, 'note_fields': {'工程量': 120, '当前状态': '施工中'}}], self.snap)
        self.assertEqual(self.snap['tasks'][0]['Notes'], '现场说明')
        self.assertIn('工程量：120', result[0]['fields']['Notes'])

    def test_repeated_fields_rejected(self):
        with self.assertRaises(ValueError):
            validate_ops([{'unique_id': 12, 'fields': {'Name': 'A'}},
                          {'unique_id': 12, 'fields': {'Name': 'B'}}], self.snap)

    def test_unsupported_operation_rejected(self):
        with self.assertRaises(ValueError):
            validate_ops([{'unique_id': 12, 'fields': {'OutlineLevel': 2}}], self.snap)

    def test_bad_change_id_cannot_escape_directory(self):
        with self.assertRaises(ValueError):
            load_change('LIVE-../../elsewhere')

    def test_fingerprint_detects_user_edits(self):
        original = digest(self.snap)
        self.snap['tasks'][0]['Name'] = '人工修改'
        self.assertNotEqual(original, digest(self.snap))

    def test_timeout_does_not_quit_application(self):
        import subprocess
        with patch('live_project.project_lock'), patch('live_project.subprocess.Popen') as popen:
            popen.return_value.communicate.side_effect = [subprocess.TimeoutExpired('worker', 20), ('', '')]
            with self.assertRaisesRegex(RuntimeError, '超时'):
                run_live('read', {'file': 'sample.mpp'})
            popen.return_value.kill.assert_called_once()

    def test_caret_blocks_live_com_before_attach(self):
        with patch('win32gui.EnumWindows'), patch('win32gui.GetForegroundWindow', return_value=100), \
             patch('win32process.GetWindowThreadProcessId', return_value=(7, 123)), \
             patch('live_project._gui_info', return_value={'flags': 0, 'hwndFocus': 0, 'hwndCaret': 100}):
            with self.assertRaisesRegex(RuntimeError, '编辑'):
                _ui_idle(123)

    def test_native_gui_info_available(self):
        import win32api
        info = _gui_info(win32api.GetCurrentThreadId())
        self.assertIn('hwndCaret', info)


if __name__ == '__main__':
    unittest.main()
