"""Preview, copy-only native Gantt PDF export and PDF inspection.

Optional dependencies are imported only when exporting/checking a PDF.
No physical print jobs or changes to a source MPP are supported.
"""
from __future__ import annotations

import gc
import hashlib
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
import tempfile
from collections import Counter
from datetime import date, datetime, timedelta

from com_runtime import exclusive_project, project_processes

ROOT = Path(__file__).resolve().parent.parent
FIELDS = {"id": "pjTaskID", "name": "pjTaskName", "start": "pjTaskStart",
          "finish": "pjTaskFinish", "percent_complete": "pjTaskPercentComplete",
          "duration": "pjTaskDuration", "resource_names": "pjTaskResourceNames"}
DEFAULTS = {"paper": "A3", "orientation": "landscape", "font_size": 10,
            "font_name": "", "timescale": "days", "inside_progress": True,
            "header": "", "footer": "", "start": None, "finish": None,
            "columns": [{"field": "id", "width": 5}, {"field": "name", "width": 34},
                        {"field": "start", "width": 13}, {"field": "finish", "width": 13},
                        {"field": "percent_complete", "width": 7}]}

PROFILE_SCHEMA = {"type": "object", "additionalProperties": False, "properties": {
    "paper": {"type":"string", "enum":["A3","A4"], "default":"A3"},
    "orientation": {"type":"string", "enum":["landscape","portrait"], "default":"landscape"},
    "font_size": {"type":"integer", "minimum":9, "maximum":16, "default":10},
    "font_name": {"type":"string", "description":"Installed font display name; empty preserves the current font"},
    "timescale": {"type":"string", "enum":["days","weeks"], "default":"days"},
    "inside_progress": {"type":"boolean", "default":True},
    "header": {"type":"string", "maxLength":200}, "footer": {"type":"string", "maxLength":200},
    "start": {"type":["string","null"], "format":"date"}, "finish": {"type":["string","null"], "format":"date"},
    "columns": {"type":"array", "minItems":1, "maxItems":7, "items": {
        "type":"object", "additionalProperties":False, "required":["field","width"], "properties":{
            "field":{"type":"string", "enum":list(FIELDS)}, "width":{"type":"integer", "minimum":4, "maximum":80},
            "title":{"type":"string", "maxLength":80}}}}}}


def sha256(path):
    digest = hashlib.sha256()
    with Path(path).open("rb") as stream:
        for block in iter(lambda: stream.read(1024*1024), b""):
            digest.update(block)
    return digest.hexdigest()


def normalize(profile=None):
    if profile is not None and not isinstance(profile, dict):
        raise ValueError("profile must be an object")
    profile = profile or {}
    if set(profile)-set(DEFAULTS):
        raise ValueError(f"Unsupported layout fields: {sorted(set(profile)-set(DEFAULTS))}")
    result = json.loads(json.dumps(DEFAULTS))
    result.update(profile)
    for key, choices in (("paper", ("A3", "A4")), ("orientation", ("landscape", "portrait")),
                         ("timescale", ("days", "weeks"))):
        if result[key] not in choices:
            raise ValueError(f"Invalid {key}")
    if type(result["font_size"]) is not int or not 9 <= result["font_size"] <= 16:
        raise ValueError("font_size must be an integer between 9 and 16")
    if type(result["inside_progress"]) is not bool:
        raise ValueError("inside_progress must be boolean")
    for key in ("header", "footer", "font_name"):
        text = result[key]
        if not isinstance(text, str) or len(text)>200 or any(c in text for c in ('&', '\n', '\r', '\x00')):
            raise ValueError(f"{key} must be plain single-line text without format codes")
    for key in ("start", "finish"):
        if result[key] is not None:
            if not isinstance(result[key],str) or not re.fullmatch(r"\d{4}-\d{2}-\d{2}",result[key]):
                raise ValueError("Dates must use YYYY-MM-DD")
            date.fromisoformat(result[key])
    if (result["start"] is None) != (result["finish"] is None):
        raise ValueError("Specify both start and finish, or neither")
    if result["start"] and result["finish"] < result["start"]:
        raise ValueError("finish precedes start")
    columns = result["columns"]
    if not isinstance(columns, list) or not 1 <= len(columns) <= 7:
        raise ValueError("columns must contain 1 to 7 fields")
    seen = set()
    for col in columns:
        if not isinstance(col, dict) or set(col)-{"field", "width", "title"}:
            raise ValueError("Invalid column object")
        field, width = col.get("field"), col.get("width")
        if field not in FIELDS or field in seen:
            raise ValueError("Unknown or duplicate column field")
        seen.add(field)
        if type(width) is not int or not 4 <= width <= 80:
            raise ValueError("Column width must be an integer between 4 and 80")
        if "title" in col and (not isinstance(col["title"], str) or len(col["title"])>80 or '\n' in col["title"]):
            raise ValueError("Invalid column title")
    if "name" not in seen:
        raise ValueError("The name column is required")
    return result


def preview(file, output_dir, profile=None):
    source = Path(file).resolve(strict=True)
    if not source.is_file() or source.suffix.lower() != ".mpp":
        raise ValueError("Source must be an existing .mpp file")
    destination = Path(output_dir).resolve()
    if destination.exists():
        raise FileExistsError("output_dir must not exist; overwriting is not allowed")
    settings = normalize(profile)
    result = {"source": str(source), "output_dir": str(destination),
              "source_sha256": sha256(source), "profile": settings}
    result["preview_sha256"] = hashlib.sha256(json.dumps(result, sort_keys=True, ensure_ascii=False).encode("utf-8")).hexdigest()
    result.update(status="preview", source_modified=False,
                  outputs=["layout.mpp", "native.pdf", "reopened.pdf", "print.pdf", "verification.json"],
                  limitations=["Project must be closed; saved data only", "No physical printer jobs",
                               "Date window can exclude tasks; short bars may not fit their labels",
                               "New dedicated Gantt table; visual review remains required"])
    return result


def _fitz():
    try:
        import fitz
    except ImportError as exc:
        raise RuntimeError("PDF tools require: pip install -r requirements-print.txt") from exc
    return fitz


def check_pdf(file, paper=None, orientation="landscape"):
    fitz = _fitz()
    with fitz.open(file) as document:
        if document.is_encrypted:
            raise ValueError("Encrypted PDF is unsupported")
        pages, checks = [], []
        dims = {"A3": (420, 297), "A4": (297, 210)}
        if paper is not None and paper not in dims:
            raise ValueError("Invalid paper")
        if orientation not in ("landscape", "portrait"):
            raise ValueError("Invalid orientation")
        for page in document:
            text = page.get_text()
            sizes = [s["size"] for b in page.get_text("dict")["blocks"] if "lines" in b
                     for line in b["lines"] for s in line["spans"] if s["text"].strip()]
            width, height = page.rect.width*25.4/72, page.rect.height*25.4/72
            info = {"page": page.number+1, "width_mm": round(width, 2), "height_mm": round(height, 2),
                    "text_characters": len(text), "min_font_pt": round(min(sizes),2) if sizes else None}
            page_ok = bool(text.strip()) and "&[" not in text and bool(sizes) and min(sizes)>=8.5
            if paper:
                expected = dims[paper] if orientation == "landscape" else dims[paper][::-1]
                info["paper_matches"] = abs(width-expected[0])<1 and abs(height-expected[1])<1
                page_ok = page_ok and info["paper_matches"]
            info["basic_checks_passed"] = page_ok
            pages.append(info)
            checks.append(page_ok)
    return {"file": str(Path(file).resolve()), "page_count": len(pages), "pages": pages,
            "basic_checks_passed": bool(pages) and all(checks), "visual_review_required": True,
            "scope": "Text presence, paper size, minimum font and literal field codes only; not a clipping or completeness proof"}


def _snapshot(project):
    return {"tasks": [{key: str(getattr(t, key)) for key in
                        ("UniqueID", "Name", "Start", "Finish", "Duration", "PercentComplete", "Predecessors",
                         "ResourceNames", "OutlineLevel", "Milestone", "Manual", "Notes")}
                       for t in project.Tasks if t is not None],
            "resources": [{key: str(getattr(r,key)) for key in ("UniqueID", "Name", "MaxUnits")}
                          for r in project.Resources if r is not None],
            "assignments": [{key: str(getattr(a,key)) for key in ("TaskUniqueID", "ResourceUniqueID", "Units", "Work")}
                            for t in project.Tasks if t is not None for a in t.Assignments]}


def _worker(request):
    import pythoncom
    import win32com.client
    # Parent holds shared lock; never attach an existing instance.
    if project_processes():
        raise RuntimeError("Project is already running; no document opened")
    plan = preview(request["source"], request["output_dir"], request["profile"])
    if plan["preview_sha256"] != request["preview_sha256"]:
        raise ValueError("Source/profile changed since preview")
    _fitz()
    output, source = Path(plan["output_dir"]), Path(plan["source"])
    output.mkdir(parents=True, exist_ok=False)
    working = output/"layout.mpp"
    shutil.copy2(source, working)
    if sha256(working) != plan["source_sha256"]:
        raise ValueError("Source changed while copying")
    report = {"status": "failed", "source": str(source), "output_dir": str(output), "profile": plan["profile"]}
    app = project = view = None
    owned = False
    pythoncom.CoInitialize()
    try:
        app = win32com.client.DispatchEx("MSProject.Application")
        if app.Projects.Count:
            raise RuntimeError("Unexpected open document; refusing to touch it")
        owned = True
        app = win32com.client.gencache.EnsureDispatch(app)
        c = win32com.client.constants
        app.Visible, app.DisplayAlerts = False, False
        app.AskToUpdateLinks = False
        app.FileOpenEx(Name=str(working))
        project = app.ActiveProject
        if project.Subprojects.Count:
            raise ValueError("Linked/inserted subprojects are not supported by this print exporter")
        before = _snapshot(project)
        report["task_count"] = len(before["tasks"])
        report["table_percentages"] = dict(Counter(t["PercentComplete"]+"%" for t in before["tasks"]))
        report["inside_percentages"] = dict(Counter(t["PercentComplete"]+"%" for t in before["tasks"] if t["Milestone"]=="False"))
        profile = plan["profile"]
        views = [v.Name for v in project.ViewsSingle if v.Screen == c.pjGantt]
        if not views:
            raise RuntimeError("No Gantt view available in this document")
        app.ViewApply(Name=views[0], SinglePane=True)
        app.FilterClear()
        app.GroupClear()
        app.OutlineShowAllTasks()
        table = "MCP Print " + plan["preview_sha256"][:12]
        for index, column in enumerate(profile["columns"]):
            field = app.FieldConstantToFieldName(getattr(c, FIELDS[column["field"]]))
            if app.TableEditEx(Name=table, TaskTable=True, Create=index==0, NewFieldName=field,
                               Title=column.get("title", "ID" if column["field"]=="id" else field), Width=column["width"], ColumnPosition=-1,
                               RowHeight=2, WrapText=True, DateFormat=c.pjDate_mm_dd,
                               ShowAddNewColumn=False) is False:
                raise RuntimeError("Cannot configure print table")
        app.TableApply(table)
        app.TextStylesEx(Item=c.pjAll, Size=profile["font_size"])
        if profile["font_name"]:
            app.TextStylesEx(Item=c.pjAll, Font=profile["font_name"])
        app.TimescaleEdit(MajorUnits=c.pjTimescaleWeeks if profile["timescale"]=="days" else c.pjTimescaleMonths,
                          MinorUnits=c.pjTimescaleDays if profile["timescale"]=="days" else c.pjTimescaleWeeks,
                          MajorCount=1, MinorCount=1)
        if profile["inside_progress"]:
            field = app.FieldConstantToFieldName(c.pjTaskPercentComplete)
            # Project localizes style names. Known names are tried; PDF checks below
            # must confirm this is visible, rather than trusting the COM return value.
            changed = []
            for alternatives in (("Task", "任务"), ("Manually Scheduled", "手动任务")):
                for style in alternatives:
                    try:
                        if app.GanttBarStyleEdit(Item=style, InsideText=field) is not False:
                            changed.append(style)
                            break
                    except pythoncom.com_error:
                        continue
            if not changed:
                raise RuntimeError("No supported task bar styles found for progress labels")
            report["progress_styles"] = changed
        app.FilePageSetupView(AllSheetColumns=True, PrintNotes=False, PrintBlankPages=False, BestPageFitTimescale=True)
        app.FilePageSetupLegend(LegendOn=c.pjNoLegend)
        for alignment in (0, 1, 2):
            app.FilePageSetupHeader(Alignment=alignment, Text=profile["header"] if alignment==1 else "")
            app.FilePageSetupFooter(Alignment=alignment, Text=profile["footer"] if alignment==1 else "")
        if app.FilePageSetupPage(Portrait=profile["orientation"]=="portrait", PagesWide=1, PagesTall=99,
                                  PaperSize=c.pjPaperA3 if profile["paper"]=="A3" else c.pjPaperA4) is False:
            raise RuntimeError("Page setup failed")
        options = {"Filename": str(output/"native.pdf"), "FileType": 0,
                   "IncludeDocumentProperties": False, "IncludeDocumentMarkup": False}
        if profile["start"]:
            options.update(FromDate=datetime.combine(date.fromisoformat(profile["start"]), datetime.min.time()),
                           ToDate=datetime.combine(date.fromisoformat(profile["finish"]), datetime.max.time()))
        if app.DocumentExport(**options) is False:
            raise RuntimeError("Native PDF export failed")
        if _snapshot(project) != before:
            raise RuntimeError("Task/resource data changed during layout")
        app.FileSave()  # Save the copied MPP in place: avoids SaveAs view loss.
        app.FileCloseEx(0)
        project = None
        app.FileOpenEx(Name=str(working), ReadOnly=True)
        project = app.ActiveProject
        report["saved_data_unchanged"] = _snapshot(project)==before
        if not report["saved_data_unchanged"]:
            raise RuntimeError("Saved copy failed data verification")
        reopened_options = dict(options, Filename=str(output/"reopened.pdf"))
        if app.DocumentExport(**reopened_options) is False:
            raise RuntimeError("Saved layout PDF re-export failed")
        fitz = _fitz()
        with fitz.open(output/"native.pdf") as first, fitz.open(output/"reopened.pdf") as reopened:
            signature = lambda doc: [(p.rect.width,p.rect.height,p.get_text()) for p in doc]
            report["saved_layout_text_unchanged"] = signature(first)==signature(reopened)
        if not report["saved_layout_text_unchanged"]:
            raise RuntimeError("Saved layout text/page structure changed after reopening")
        app.FileCloseEx(0)
        project = None
        with fitz.open(output/"native.pdf") as pdf:
            for i, page in enumerate(pdf):
                page.insert_text((page.rect.width-85, page.rect.height-25), f"{i+1} / {len(pdf)}", fontsize=9)
            pdf.save(output/"print.pdf")
        if profile["inside_progress"] and not profile["start"]:
            with fitz.open(output/"print.pdf") as pdf:
                actual = Counter(w[4] for page in pdf for w in page.get_text("words") if re.fullmatch(r"\d+%",w[4]))
            expected = Counter(report["inside_percentages"])
            if any(col["field"]=="percent_complete" for col in profile["columns"]):
                expected += Counter(report["table_percentages"])
            report["progress_text_matches_expected"] = actual==expected
            if actual!=expected:
                report["warnings"] = ["Percentage text count differs from expected; inspect short bars and custom styles. This is a heuristic, not a layout proof."]
        report["pdf"] = check_pdf(output/"print.pdf", profile["paper"], profile["orientation"])
        report["source_unchanged"] = sha256(source)==plan["source_sha256"]
        report["status"] = "exported_needs_visual_review" if report["pdf"]["basic_checks_passed"] and report["source_unchanged"] else "verification_failed"
        report["pdf_path"], report["mpp_copy"] = str(output/"print.pdf"), str(working)
        return report
    finally:
        if app is not None and owned:
            try:
                app.FileCloseEx(0)
            except Exception:
                pass
            try:
                app.Quit(0)
            except Exception as exc:
                report["cleanup_error"] = str(exc)
                report["status"] = "cleanup_failed"
        app = project = view = None
        gc.collect()
        pythoncom.CoUninitialize()
        report["source_unchanged"] = sha256(source)==plan["source_sha256"]
        if not report["source_unchanged"]:
            report["status"] = "verification_failed"
        (output/"verification.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")


def export(file, output_dir, profile=None, preview_sha256=""):
    request = preview(file, output_dir, profile)
    if not preview_sha256 or preview_sha256 != request["preview_sha256"]:
        raise ValueError("Run project_preview_print_layout first; provide its matching preview_sha256")
    _fitz()  # Fail before starting Project if optional dependency is missing.
    with exclusive_project():
        with tempfile.TemporaryDirectory(prefix="project-print-") as temp:
            path = Path(temp)/"request.json"
            path.write_text(json.dumps(request, ensure_ascii=False), encoding="utf-8")
            env = dict(os.environ, PROJECT_PRINT_PARENT=str(os.getpid()))
            try:
                worker = subprocess.run([sys.executable, str(Path(__file__).resolve()), "--worker", str(path)],
                                        capture_output=True, text=True, encoding="utf-8", errors="replace",
                                        timeout=int(os.environ.get("PROJECT_COM_TIMEOUT", "900")), env=env,
                                        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
            except subprocess.TimeoutExpired as exc:
                raise RuntimeError("PDF export timed out; inspect the test-owned Project process and output before retrying") from exc
    if worker.returncode:
        raise RuntimeError(f"PDF export failed; partial output retained: {worker.stderr[-2000:]}")
    result = json.loads(worker.stdout.strip().splitlines()[-1])
    if result["status"] != "exported_needs_visual_review" or not result.get("source_unchanged"):
        raise RuntimeError(f"PDF verification failed; inspect {output_dir}")
    return result


if __name__ == "__main__":
    if len(sys.argv)!=3 or sys.argv[1]!="--worker":
        raise SystemExit("Use project_cli.py or the MCP print tools")
    if os.environ.get("PROJECT_PRINT_PARENT") != str(os.getppid()):
        raise SystemExit("Worker must be started by the locked export entry point")
    request = json.loads(Path(sys.argv[2]).read_text(encoding="utf-8"))
    print(json.dumps(_worker(request), ensure_ascii=False))
