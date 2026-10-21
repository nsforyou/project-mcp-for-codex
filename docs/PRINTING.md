# 甘特图排版 / Gantt print layout

Version 0.4.0. Install optional PDF dependencies into the server's Python environment:

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements-print.txt
```

## Preview and export / 预览与执行

Save the following as `print-preview.json`, replacing the paths. 输出目录必须尚不存在。

```json
{
  "file": "C:\\ProjectMCP\\plans\\demo.mpp",
  "output_dir": "C:\\ProjectMCP\\exports\\demo-print-01",
  "profile": {
    "paper": "A3",
    "orientation": "landscape",
    "font_size": 10,
    "timescale": "days",
    "inside_progress": true,
    "header": "Construction schedule",
    "footer": "Review before use",
    "columns": [
      {"field": "name", "width": 34, "title": "Task name"},
      {"field": "start", "width": 13},
      {"field": "finish", "width": 13},
      {"field": "percent_complete", "width": 7}
    ]
  }
}
```

```powershell
$preview = .\.venv\Scripts\python.exe mcp\project_cli.py call project_preview_print_layout --args-file print-preview.json | ConvertFrom-Json
$execute = Get-Content -Raw print-preview.json | ConvertFrom-Json
$execute | Add-Member -NotePropertyName preview_sha256 -NotePropertyValue $preview.preview_sha256
$execute | ConvertTo-Json -Depth 10 | Set-Content -Encoding UTF8 print-export.json
.\.venv\Scripts\python.exe mcp\project_cli.py call project_export_pdf --args-file print-export.json
```

Preview does not open Project or write output. Execution requires the exact same source, destination and normalized profile, bound by `preview_sha256`. If any of these change, preview again. A malformed or stale hash fails before COM. The output directory is created exclusively, and failures retain partial output for diagnosis.

预览不连接 Project、不写文件。执行指纹绑定源文件内容、参数和目标路径；任何一项变化后都需重新预览。执行只打开源文件的副本，修改副本的专用打印表和视图，使用原生 PDF 导出，再在 PDF 副本中补充页码。副本就地保存并重新打开，比较任务、资源、分配；源文件另作 SHA256 核验。不通过核验则不报告成功。

## Profile fields / 参数

| Field | Accepted values / 说明 |
|---|---|
| `paper` | `A3` (default), `A4` |
| `orientation` | `landscape` (default), `portrait` |
| `font_size` | Integer 9–16 points, default 10 |
| `font_name` | Default empty: keep existing font. Otherwise use an installed font's display name accepted by Project; localized font names may be required. 默认保留原字体。 |
| `timescale` | `days` (week/day tiers), `weeks` (month/week tiers) |
| `inside_progress` | Default `true`: bind Percent Complete inside standard normal/manual task bars. `false`: leave existing bar text unchanged. 不改完成率。 |
| `header`, `footer` | Plain single-line text, up to 200 characters. No `&` field/format codes. PDF page numbering is added separately. |
| `start`, `finish` | Specify both or neither; ISO `YYYY-MM-DD`. Omit to use Project's full project range. 日期窗口可能排除任务。 |
| `columns` | 1–7 unique columns; `name` required. Each object has `field`, integer `width` 4–80 and optional `title`. |

Column fields: `id`, `name`, `start`, `finish`, `duration`, `percent_complete`, `resource_names`. Column width uses Project's native width units; row height and wrapping are set in the copied table. The exported view expands tasks and clears filters/grouping, so it is a dedicated full task Gantt layout, not an exact reproduction of the current visible view.

## Results and review / 成果与验收

- `layout.mpp`: copied source with print layout saved in place; source stays untouched.
- `native.pdf`: native Project export, retained for comparison.
- `reopened.pdf`: native re-export after reopening the saved copy, checked for the same page structure and text.
- `print.pdf`: numbered delivery PDF.
- `verification.json`: normalized settings, task/resource/assignment consistency, source hash check and basic PDF checks.

`project_check_pdf_layout` accepts `{ "file": ".../print.pdf", "paper": "A3", "orientation": "landscape" }`. `paper` is optional. It checks page size, presence of text, minimum font and literal field codes. This is not a proof of task completeness, clipping, correct custom shapes or readability.

The successful export status is `exported_needs_visual_review`. Review every page. Short bars may not fit labels; milestones remain zero-length symbols. Large plans can have substantial final-page whitespace. Custom bar styles or languages other than English/Simplified Chinese require separate adaptation. Do not infer that every drawing, OLE object, inserted subproject or custom view has been validated.

基础核验不替代目视验收。日期范围、分页、字体、长名称、甘特条标签和里程碑都需检查。此版本没有自动反复调版、物理打印机控制或打印完成回执。

## Concurrent MCP installations / 多实例

The export shares the offline/live COM lock and refuses to start while any `WINPROJ.EXE` is running. It never attaches to or closes the user's existing Project window. Multiple checkouts must set the same absolute `PROJECT_COM_LOCK` in their server environments; the default remains `mcp/work/project-com.lock` within each checkout. Example for coordinated testing:

```powershell
$env:PROJECT_COM_LOCK = 'C:\ProjectMCP\shared\project-com.lock'
```

All instances must use the same value and a version supporting this setting. A timeout may leave the test-owned Project process or partial output; inspect it before retrying. No existing MCP process is restarted by these tools. Reconnect the client deliberately when ready to load the new tool list.

## Tests / 测试

```powershell
python -m unittest discover -s mcp/tests -p 'test_*.py'
python mcp/tests/print_integration.py
# Additional A4 portrait/week-timescale profile:
python mcp/tests/print_integration.py --extended
```

Native tests require installed desktop Project and optional PDF dependencies, and must run when Project is idle. They use the synthetic A/B fixture by default; `--file` is for another disposable synthetic MPP only. Never pass a business plan to integration tests. Rendered pages and test reports stay in ignored `mcp/work/` folders, and are not published to GitHub.
