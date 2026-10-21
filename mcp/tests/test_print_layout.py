import json
from pathlib import Path
import sys
import tempfile
import subprocess
import unittest
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import print_layout as printing
import project_mcp_server as server


class PrintTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.source = self.root/'sample.mpp'
        self.source.write_bytes(b'synthetic placeholder; preview does not invoke Project')
        self.out = self.root/'new-output'

    def tearDown(self):
        self.temp.cleanup()

    def test_preview_is_read_only_and_deterministic(self):
        before = self.source.read_bytes()
        first = printing.preview(self.source, self.out)
        self.assertEqual(first, printing.preview(self.source, self.out))
        self.assertFalse(self.out.exists())
        self.assertEqual(self.source.read_bytes(), before)

    def test_hash_binds_source_settings_and_destination(self):
        first = printing.preview(self.source, self.out)['preview_sha256']
        for output, profile in ((self.out, {'paper':'A4'}), (self.root/'other', {})):
            self.assertNotEqual(first, printing.preview(self.source, output, profile)['preview_sha256'])
        self.source.write_bytes(b'changed')
        self.assertNotEqual(first, printing.preview(self.source, self.out)['preview_sha256'])

    def test_wrong_hash_fails_before_lock_or_com(self):
        with patch.object(printing, 'exclusive_project') as guard:
            with self.assertRaisesRegex(ValueError, 'preview'):
                printing.export(self.source, self.out, preview_sha256='wrong')
            guard.assert_not_called()
        self.assertFalse(self.out.exists())

    def test_existing_output_rejected(self):
        self.out.mkdir()
        (self.out/'keep').write_text('keep')
        with self.assertRaises(FileExistsError):
            printing.preview(self.source, self.out)
        self.assertEqual((self.out/'keep').read_text(), 'keep')

    def test_invalid_profiles(self):
        for profile in ({'font_size':5}, {'font_size':True}, {'inside_progress':'yes'}, {'paper':'A0'},
                        {'start':'2026-11-01'}, {'start':'2026-12-01','finish':'2026-11-01'},
                        {'header':'&[Page]'}, {'unknown':1}, {'columns':[]},
                        {'columns':[{'field':'name','width':1}]},
                        {'columns':[{'field':'name','width':20},{'field':'name','width':20}]},
                        {'columns':[{'field':'id','width':5}]}):
            with self.subTest(profile=profile), self.assertRaises((ValueError, TypeError)):
                printing.normalize(profile)

    def test_no_shared_mutable_defaults(self):
        a = printing.normalize()
        a['columns'][0]['width'] = 77
        self.assertEqual(printing.normalize()['columns'][0]['width'],5)

    def test_busy_project_does_not_start_worker(self):
        preview = printing.preview(self.source, self.out)
        with patch('com_runtime.project_processes', return_value=[{'pid':1}]), \
             patch('com_runtime.LOCK',self.root/'guard.lock'), \
             patch.object(printing,'_fitz'), patch.object(printing.subprocess,'run') as worker:
            with self.assertRaisesRegex(RuntimeError,'正在运行'):
                printing.export(self.source,self.out,preview_sha256=preview['preview_sha256'])
            worker.assert_not_called()
        self.assertFalse(self.out.exists())

    def test_tools_registered_and_protocol_errors_stay_local(self):
        self.assertEqual(len(server.TOOLS),30)
        request = {'jsonrpc':'2.0','id':1,'method':'tools/call','params':{
            'name':'project_export_pdf','arguments':{'file':str(self.source),'output_dir':str(self.out),'preview_sha256':'bad'}}}
        self.assertTrue(server.handle(request)['result']['isError'])
        self.assertEqual(server.handle({'id':2,'method':'ping'})['result'],{})

    def test_stdio_initialization_and_tool_list(self):
        messages = [{'jsonrpc':'2.0','id':1,'method':'initialize'},
                    {'jsonrpc':'2.0','id':2,'method':'tools/list'}]
        result = subprocess.run([sys.executable,str(Path(server.__file__))],
                                input=''.join(json.dumps(m)+'\n' for m in messages),
                                capture_output=True,text=True,encoding='utf-8',timeout=15)
        self.assertEqual(result.returncode,0,result.stderr)
        responses = [json.loads(line) for line in result.stdout.splitlines()]
        self.assertEqual(responses[0]['result']['serverInfo']['version'],'0.4.0')
        tools = responses[1]['result']['tools']
        self.assertEqual(len(tools),30)
        self.assertEqual(len({t['name'] for t in tools}),30)
        schema = next(t for t in tools if t['name']=='project_export_pdf')['inputSchema']
        self.assertIn('inside_progress',schema['properties']['profile']['properties'])

    def test_pdf_checks_reject_wrong_paper_and_small_font(self):
        try:
            import fitz
        except ImportError:
            self.skipTest('optional PDF dependency not installed')
        pdf = self.root/'test.pdf'
        with fitz.open() as document:
            page = document.new_page(width=842,height=595)
            page.insert_text((30,30),'Synthetic PDF',fontsize=6)
            document.save(pdf)
        result = printing.check_pdf(pdf,'A3')
        self.assertFalse(result['basic_checks_passed'])
        self.assertFalse(result['pages'][0]['paper_matches'])


if __name__=='__main__':
    unittest.main()
