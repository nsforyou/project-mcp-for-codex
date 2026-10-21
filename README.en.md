# Project MCP for codex

Authors: **JhonPrice**, ChatGPT-6.1 Sol

English · [中文](README.md)

[Noncommercial license](LICENSE) · [Issues](https://github.com/nsforyou/project-mcp-for-codex/issues)

Version **0.4.0**, **30 MCP tools**. Read and edit Microsoft Project plans on Windows through Codex, or use the same handlers through a CLI. The repository includes only synthetic tasks A/B and resources AA/BB.

## Capabilities

- Read tasks, dependencies, resources, calendars, baselines and schedule summaries.
- Read unsaved task data in an already-open Project window.
- Preview approved live edits, apply them in a native undo group, and leave the window open for review/save.
- Back up and apply offline task changes, verify saved data, and restore the source through rollback.
- Preserve prose while maintaining structured note fields for quantities and status.
- Create a new MPP from basic Project XML and export task lists to Excel.
- Preview Gantt print settings, export an MPP copy plus A3/A4 PDF, and run basic PDF checks.
- Show the current Percent Complete field inside standard task bars without changing progress data.

The tool does not make construction decisions or provide a complete resource-leveling algorithm. Changing dates does not guarantee correct downstream scheduling when dependencies are absent or tasks are manually scheduled.

## Requirements and installation

Windows, desktop Microsoft Project with COM support (validated on Professional 2019 / version 16.0), Windows Python (validated with 64-bit Python 3.13), and Java 21 for MPXJ. Run in the same desktop user session as Project. Web Project/Planner and a standalone WSL, Linux or macOS environment cannot replace Project COM.

```powershell
git clone https://github.com/nsforyou/project-mcp-for-codex.git C:\ProjectMCP\project-mcp-for-codex
Set-Location C:\ProjectMCP\project-mcp-for-codex
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
$env:MPXJ_LIB = 'C:\ProjectMCP\deps'
.\.venv\Scripts\python.exe -m pip install --target $env:MPXJ_LIB -r requirements-mpxj.txt
# Optional: native PDF export and checks
.\.venv\Scripts\python.exe -m pip install -r requirements-print.txt
```

Use an ASCII-only MPXJ dependency path. Configure Java 21 on PATH or JAVA_HOME. Open Project manually once to complete first-run/activation prompts. The server does not call the OpenAI API or require an API key. Tool results can be passed to the model by the host; local file access does not mean offline model inference.

## Connect an MCP client

Add this to your existing Codex configuration, adjusting paths:

```toml
[mcp_servers.microsoft_project]
command = 'C:\ProjectMCP\project-mcp-for-codex\.venv\Scripts\python.exe'
args = ['C:\ProjectMCP\project-mcp-for-codex\mcp\project_mcp_server.py']

[mcp_servers.microsoft_project.env]
MPXJ_LIB = 'C:\ProjectMCP\deps'
```

See the [official MCP configuration documentation](https://developers.openai.com/codex/mcp/) and the [Chinese guide](README.md) for the full installation and collaboration walkthrough. Do not register duplicate instances unintentionally. Reconnect the client after an update to load the new tool list; updating files does not restart an existing server.

## Preview and edit

Keep Project open when using `project_live_read`, `project_live_preview` and `project_live_apply`. Explicitly provide the intended MPP path. Save first and finish cell editing before previewing a write. A live edit stays in memory for the user to review and save, and uses a native undo group.

Offline COM operations require Project to be closed, and refuse to attach to an existing window. The default offline save is in place after backup; this preserves source view streams better than SaveAs. Most ordinary disk reads use MPXJ and do not see unsaved edits.

CLI tool inventory and invocation:

```powershell
.\.venv\Scripts\python.exe mcp\project_cli.py tools
.\.venv\Scripts\python.exe mcp\project_cli.py call project_get_project_info --args-file info.json
```

`info.json` contains `{ "file": "C:\\ProjectMCP\\plans\\demo.mpp" }`. The CLI and MCP use exactly the same handlers. Use task UniqueID as the stable identifier. Review a proposed change before supplying its approved execution parameters.

## Gantt formatting and printable PDF

Use `project_preview_print_layout` to validate settings without COM or writes. Then call `project_export_pdf` with the matching `preview_sha256`, identical settings and a new output directory. The hash binds source contents, destination and settings. Execution copies the source, applies layout to that copy, saves it in place, reopens it and compares task/resource/assignment data. It verifies that the source hash is unchanged.

Profiles support A3/A4, landscape/portrait, columns and widths, wrapping, 9–16 point text, day/week timescales, paired start/finish dates, plain-text headers/footers and inside-bar progress. Default font family is inherited. `inside_progress: false` leaves existing bar text unchanged. Standard English and Simplified Chinese normal/manual task styles are supported; custom styles require validation.

Outputs are `layout.mpp`, `native.pdf`, numbered `print.pdf` and `verification.json`. `project_check_pdf_layout` checks paper size, text presence, minimum font and unexpanded field codes. Success is `exported_needs_visual_review`: inspect every page for clipping, missing tasks, labels and readability. No physical print job is sent. See [parameters and complete CLI example](docs/PRINTING.md).

Multiple installations must use a shared absolute `PROJECT_COM_LOCK` to coordinate access to the same Project process. Existing servers keep their current lock setting until deliberately reconnected. Failed or timed-out exports retain partial output and do not report success.

## All 30 tools

| Category | Tools |
|---|---|
| Live collaboration (3) | `project_live_read`, `project_live_preview`, `project_live_apply` |
| File context (2) | `project_open`, `project_close` |
| Project/schedule (3) | `project_get_project_info`, `project_get_schedule_summary`, `project_get_critical_path` |
| Tasks/dependencies (4) | `project_list_tasks`, `project_find_tasks`, `project_get_task`, `project_get_dependencies` |
| Resources/views (4) | `project_list_resources`, `project_list_calendars`, `project_get_baselines`, `project_get_view_layer` |
| Offline changes (5) | `project_preview_changes`, `project_apply_changes`, `project_verify_changes`, `project_rollback_changes`, `project_list_change_sets` |
| Notes/progress (4) | `project_append_task_note`, `project_set_actual_progress`, `project_get_task_note_fields`, `project_preview_task_note_fields` |
| Create/Excel (2) | `project_create_from_xml`, `project_export_excel` |
| Print/PDF (3) | `project_preview_print_layout`, `project_export_pdf`, `project_check_pdf_layout` |

`project_close` clears the server's file context; it does not close a user's Project window. Complete parameter schemas are returned by the `tools` CLI command.

## Tests, recovery and updates

```powershell
python -m unittest discover -s mcp/tests -p 'test_*.py'
# Opt-in native integration tests: disposable synthetic plans only, Project idle
python mcp/tests/integration_smoke.py
python mcp/tests/print_integration.py
```

Test reports and generated plans remain in ignored local output folders. Never publish business MPPs, logs, credentials or private paths in an issue. Keep backups and change-set records until accepted results are confirmed. A successful basic PDF check does not certify engineering correctness.

For updates, use a clean checkout and fast-forward pull, update changed dependency files, then reconnect the client when safe. Do not interrupt another project's running MCP session just to reload tools.

Licensed under [PolyForm Noncommercial 1.0.0](LICENSE); retain [NOTICE](NOTICE) when distributing. Commercial use beyond the permitted purposes requires separate written authorization. Third-party libraries and Microsoft products retain their own licenses.

## Support

[Buy Me a Coffee — support OpenAI / subscribe to Codex](https://chatgpt.com/codex/pricing/) · [OpenAI official website](https://openai.com/)

If you would like to support the tools behind this project, consider an official plan that suits your needs. This link goes to OpenAI's subscription information, not a personal donation account. Subscription is optional; current access and terms are listed on the official page.

## Engineering Philosophy | 工程理念

**中文**

我们相信，人工智能不应该取代工程师的判断，而应该扩展工程师的能力。

工程师负责定义目标、建立约束、识别风险并验证结果；AI 负责处理繁琐、重复且容易出错的工作。

本项目尝试将自然语言、AI Agent 与 Microsoft Project 结合，让工程计划从编制、调整、校验到排版交付，逐步形成可追溯、可验证、可自动化的工作流程。

**AI 不替工程师承担责任，但可以帮助工程师完成更多工作。**

**English**

We believe artificial intelligence should not replace engineering judgment, but extend engineering capabilities.

Engineers define objectives, establish constraints, identify risks, and validate results. AI handles repetitive, time-consuming, and error-prone tasks.

This project explores the integration of natural language, AI agents, and Microsoft Project to build a traceable, verifiable, and increasingly automated workflow for project planning, scheduling, validation, formatting, and delivery.

**AI does not replace engineering accountability. It empowers engineers to accomplish more.**

---

*Developed through human–AI collaboration.*

*Author: JhonPrice*

> Built by an engineer, developed with AI, verified in the real world.
>
> 由工程师提出需求，与人工智能协作开发，并在真实工程实践中验证。
