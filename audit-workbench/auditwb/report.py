"""Writes the outputs of a run to Excel:

  WorkingPaper_<company>_<timestamp>.xlsx
      Cover, Summary, Data Quality, Parameters and one evidence sheet per test with
      the rule, parameters, population, the exact SQL executed and every exception row.
  AnnexureII_<company>_<timestamp>.xlsx
      Company-wise Data-Driven Audit Plan in the OO-53 Annexure-II layout, pre-filled
      from the results. Columns that need the audit team's judgement (recommendation,
      reason codes) are left for the team unless set in the profile.
"""

from __future__ import annotations

import datetime as dt
import re
from pathlib import Path

from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from .engine import EXCEPTIONS, NIL, NOT_EXECUTABLE, Run, TestResult
from .ingest import sha256_of

EXCEL_MAX_ROWS = 1_000_000
LAKH = 100_000

BOLD = Font(bold=True)
TITLE = Font(bold=True, size=14)
HEADER_FILL = PatternFill("solid", fgColor="DDE4EE")
FLAG_FILL = PatternFill("solid", fgColor="FCE4D6")
OK_FILL = PatternFill("solid", fgColor="E2EFDA")
GREY_FILL = PatternFill("solid", fgColor="EDEDED")
WRAP = Alignment(wrap_text=True, vertical="top")
THIN = Side(style="thin", color="999999")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

DISCLAIMER = (
    "Results are exceptions / red flags identified by deterministic rules on ERP data. They are not audit "
    "observations. Audit observations shall be developed only after verification and substantiation of the "
    "exceptions with reference to supporting records during field audit, and no approval, authorisation or "
    "sanction shall be treated as verified on the basis of remote analysis alone (OO-53 para 10)."
)

STATUS_FILL = {EXCEPTIONS: FLAG_FILL, NIL: OK_FILL, NOT_EXECUTABLE: GREY_FILL}


def lakh(value):
    return None if value is None else round(value / LAKH, 2)


def _safe_name(text: str) -> str:
    return re.sub(r"[^A-Za-z0-9]+", "_", text).strip("_")[:40] or "company"


def _cell(ws, value, font=None, fill=None, fmt=None, wrap=False):
    c = WriteOnlyCell(ws, value=value)
    if font:
        c.font = font
    if fill:
        c.fill = fill
    if fmt:
        c.number_format = fmt
    if wrap:
        c.alignment = WRAP
    return c


def _excel_value(v):
    if isinstance(v, (list, tuple)):
        return ", ".join(str(x) for x in v)
    if isinstance(v, float) and v != v:  # NaN
        return None
    if isinstance(v, (dt.date, dt.datetime, int, float, str, bool)) or v is None:
        return v
    return str(v)


def _kv(ws, key, value, wrap=True):
    ws.append([_cell(ws, key, font=BOLD), _cell(ws, _excel_value(value), wrap=wrap)])


def _topic_label(run: Run, topic_id: str) -> str:
    t = run.library.topics[topic_id]
    return f"{t.id} {t.title} ({t.annex_ref})"


def unit_summary(result: TestResult, top: int = 3) -> str:
    if not result.units:
        return ""
    total = sum(v for _, v in result.units) or 0
    parts = []
    for unit, value in result.units[:top]:
        share = f", {round(100 * value / total)}% of value" if total > 0 else ""
        parts.append(f"{unit} ({lakh(value)} lakh{share})")
    return "; ".join(parts)


# --------------------------------------------------------------------------- working paper

def _cover(wb, run: Run):
    ws = wb.create_sheet("Cover")
    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 110
    p = run.profile
    ws.append([_cell(ws, "Audit Analytics Workbench: Working Paper", font=TITLE)])
    ws.append([])
    for k, v in [
        ("Company", p.company), ("Unit", p.unit), ("Tier", p.tier), ("Lead Audit Team", p.lead_team),
        ("Audit period", f"{p.period_from:%d.%m.%Y} to {p.period_to:%d.%m.%Y}"),
        ("Cut-off date", f"{p.cutoff_date:%d.%m.%Y}"),
        ("Run started", f"{run.started:%d.%m.%Y %H:%M:%S}"),
        ("Run finished", f"{run.finished:%d.%m.%Y %H:%M:%S}"),
        ("Run by (Windows user)", run.run_by), ("Machine", run.machine),
        ("Workbench version", run.tool_version),
        ("Catalogue", run.library.catalogue_version),
        ("Run profile", str(p.path)), ("Run profile SHA-256", sha256_of(p.path)),
        ("Tests run", len(run.results)),
        ("With exceptions", sum(r.status == EXCEPTIONS for r in run.results)),
        ("Nil exceptions", sum(r.status == NIL for r in run.results)),
        ("Not executable / error", sum(not r.executed for r in run.results)),
    ]:
        _kv(ws, k, v)
    ws.append([])
    _kv(ws, "Nature of results", DISCLAIMER)
    _kv(ws, "Reproducibility",
        "Each test sheet records the exact SQL executed and the parameter values used. Re-running the same "
        "workbench version with the same profile on files with the same SHA-256 hashes (Data Quality sheet) "
        "reproduces these results.")
    _kv(ws, "Status legend",
        "EXCEPTIONS: rule found exceptions. NIL: rule ran on the data and found none (analytical assurance, "
        "subject to the completeness of the extract). NOT EXECUTABLE: required data, fields or criteria were "
        "not available; this is NOT a nil result. ERROR: the test failed; see the note.")


def _summary(wb, run: Run):
    ws = wb.create_sheet("Summary")
    widths = [16, 48, 60, 16, 14, 16, 12, 18, 16, 60, 70]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    head = ["Test ID", "Test", "Annexure-I topic(s)", "Status", "Population (no.)",
            "Population value (Rs. lakh)", "Exceptions", "Counted as", "Value involved (Rs. lakh)",
            "Units with most exceptions", "Note"]
    ws.append([_cell(ws, h, font=BOLD, fill=HEADER_FILL, wrap=True) for h in head])
    for r in run.results:
        ws.append([
            _cell(ws, r.test.id), _cell(ws, r.test.title, wrap=True),
            _cell(ws, "; ".join(_topic_label(run, t) for t in r.test.topics), wrap=True),
            _cell(ws, r.status, fill=STATUS_FILL.get(r.status)),
            _cell(ws, r.population_n), _cell(ws, lakh(r.population_value), fmt="#,##0.00"),
            _cell(ws, r.n_exceptions if r.executed else None), _cell(ws, r.test.exception_unit),
            _cell(ws, lakh(r.value) if r.executed else None, fmt="#,##0.00"),
            _cell(ws, unit_summary(r), wrap=True), _cell(ws, r.message, wrap=True),
        ])


def _data_quality(wb, run: Run):
    ws = wb.create_sheet("Data Quality")
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 120
    ws.append([_cell(ws, "Data Quality and Completeness", font=TITLE)])
    ws.append([_cell(ws, "Reconcile the control totals below with a known figure (trial balance, report footer, "
                         "annual accounts) before relying on NIL results.", wrap=True)])
    for st in run.datasets.values():
        src = run.profile.datasets.get(st.id)
        if not src and not st.load and not st.error:
            continue
        ws.append([])
        ws.append([_cell(ws, f"{st.id}: {st.title}", font=BOLD, fill=HEADER_FILL),
                   _cell(ws, None, fill=HEADER_FILL)])
        if src and src.sap_report:
            _kv(ws, "SAP report / transaction", src.sap_report)
        if src and src.extracted_on:
            _kv(ws, "Extracted on", src.extracted_on)
        if st.error:
            _kv(ws, "NOT LOADED", st.error)
            continue
        load = st.load
        for f in load.files:
            _kv(ws, "File", f.path)
            _kv(ws, "SHA-256", f.sha256)
            if f.sheet:
                _kv(ws, "Sheet", f.sheet)
            _kv(ws, "Header found on row", f.header_row)
            _kv(ws, "Lines read / rows loaded", f"{f.rows_read} / {f.rows_loaded}")
            if f.dropped:
                _kv(ws, "Lines dropped", "; ".join(f"{k}: {v}" for k, v in f.dropped.items()))
            _kv(ws, "Mapping (standard field <- column)",
                "; ".join(f"{k} <- {v}" for k, v in f.mapping.items()))
            if f.unmapped_columns:
                _kv(ws, "Columns not used", "; ".join(f.unmapped_columns))
        unmapped = [n for n in load.dataset.fields if n not in load.mapped_fields]
        if unmapped:
            _kv(ws, "Standard fields not available", ", ".join(unmapped))
        for fname, (total, negatives) in st.control_totals.items():
            _kv(ws, f"Control total: {fname}", f"{total:,.2f} ({negatives} negative values)" if total is not None else "no values")
        for fname, (lo, hi, nulls) in st.date_ranges.items():
            _kv(ws, f"Date range: {fname}", f"{lo} to {hi} ({nulls} empty)")
        for fname, counter in load.parse_failures.items():
            examples = ", ".join(f"'{k}'" for k, _ in counter.most_common(5))
            _kv(ws, f"Unreadable values: {fname}", f"{sum(counter.values())} values set to empty, e.g. {examples}")
        for w in load.warnings:
            _kv(ws, "WARNING", w)


def _parameters(wb, run: Run):
    ws = wb.create_sheet("Parameters")
    for i, w in enumerate([16, 26, 50, 40, 40, 10], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    ws.append([_cell(ws, h, font=BOLD, fill=HEADER_FILL) for h in
               ["Test ID", "Parameter", "Meaning", "Value used", "Library default", "Changed"]])
    for r in run.results:
        for name, p in r.test.params.items():
            used = r.params.get(name, p.default)
            ws.append([_cell(ws, r.test.id), _cell(ws, name), _cell(ws, p.label, wrap=True),
                       _cell(ws, _excel_value(used)), _cell(ws, _excel_value(p.default)),
                       _cell(ws, "YES" if used != p.default else "", fill=FLAG_FILL if used != p.default else None)])


def _test_sheet(wb, run: Run, r: TestResult):
    ws = wb.create_sheet(r.test.id[:31])
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 26
    for i in range(3, 30):
        ws.column_dimensions[get_column_letter(i)].width = 18
    t = r.test
    ws.append([_cell(ws, f"{t.id}: {t.title}", font=TITLE)])
    _kv(ws, "Annexure-I topic(s)", "; ".join(_topic_label(run, x) for x in t.topics))
    _kv(ws, "Status", r.status)
    _kv(ws, "Rule", t.rule)
    _kv(ws, "Reading the result", t.interpretation)
    if any(run.library.topics[x].approval for x in t.topics):
        _kv(ws, "Approvals", "Approvals, authorisations and sanctions are not visible remotely and are NOT verified "
                             "by this test (OO-53 para 10). Verify them in the field.")
    for name, p in t.params.items():
        _kv(ws, f"Parameter: {p.label}", r.params.get(name))
    _kv(ws, "Audit period / cut-off",
        f"{run.profile.period_from:%d.%m.%Y} to {run.profile.period_to:%d.%m.%Y} / {run.profile.cutoff_date:%d.%m.%Y}")
    _kv(ws, "Datasets used", ", ".join(t.datasets))
    _kv(ws, "Population tested", f"{r.population_n} (value Rs. {lakh(r.population_value)} lakh)"
        if r.population_n is not None else "")
    _kv(ws, "Exceptions", f"{r.n_exceptions} {t.exception_unit}; value involved Rs. {lakh(r.value)} lakh")
    if r.units:
        _kv(ws, "Concentration", unit_summary(r, top=5))
    _kv(ws, "Questions for the field", "\n".join(f"- {q}" for q in t.field_questions))
    _kv(ws, "Records required for substantiation", t.records_required)
    _kv(ws, "SQL executed", t.sql.strip())
    ws.append([])
    if not r.rows:
        ws.append([_cell(ws, "No exceptions." if r.status == NIL else r.message, font=BOLD)])
        return
    ws.append([_cell(ws, c, font=BOLD, fill=HEADER_FILL) for c in r.columns])
    money_cols = {i for i, c in enumerate(r.columns)
                  if c == "exception_value" or re.search(r"value|amount|price|cost|budget|actual|balance", c)}
    for n, row in enumerate(r.rows):
        if n >= EXCEL_MAX_ROWS:
            ws.append([_cell(ws, f"Truncated: {len(r.rows) - EXCEL_MAX_ROWS} further rows not written "
                                 f"(Excel row limit). Narrow the parameters or split the period.", font=BOLD)])
            break
        ws.append([_cell(ws, _excel_value(v), fmt="#,##0.00" if i in money_cols and isinstance(v, float) else
                         ("dd.mm.yyyy" if isinstance(v, dt.date) else None))
                   for i, v in enumerate(row)])


def write_working_paper(run: Run, out_dir: Path) -> Path:
    wb = Workbook(write_only=True)
    _cover(wb, run)
    _summary(wb, run)
    _data_quality(wb, run)
    _parameters(wb, run)
    for r in run.results:
        if r.executed:
            _test_sheet(wb, run, r)
    path = out_dir / f"WorkingPaper_{_safe_name(run.profile.company)}_{run.started:%Y%m%d_%H%M%S}.xlsx"
    wb.save(path)
    return path


# --------------------------------------------------------------------------- Annexure-II

ANNEX_HEADERS = [
    "Sl.", "Business Process", "Audit Topic analysed (Priority I/II)", "Reason for Selection (Code)",
    "Data Source and Mode of Analysis (Remote / On-site)", "No. of Exceptions identified in remote analysis",
    "Value involved (Rs. in lakh)", "Recommended for Field Verification (Y/N)", "Reason for selection in detail",
    "Units/Departments proposed for Field Verification", "Records required for Substantiation",
    "For DA Cell use: Approved / Modified",
]
ANNEX_WIDTHS = [5, 20, 34, 12, 30, 22, 14, 14, 44, 30, 40, 16]
REASON_CODES = [("M", "High materiality"), ("C", "High compliance risk"),
                ("P", "Previous audit observation / inspection report"), ("D", "Data readily available"),
                ("F", "Fraud indicator"), ("S", "Company-specific strategic significance"),
                ("A", "Outcome of Accounts audit"), ("O", "Others (to be specified)")]


def annexure_rows(run: Run) -> list[dict]:
    lib, profile = run.library, run.profile
    topic_ids = profile.topics or list(dict.fromkeys(t for r in run.results for t in r.test.topics))
    rows = []
    for tid in topic_ids:
        topic = lib.topics[tid]
        results = [r for r in run.results if tid in r.test.topics]
        executed = [r for r in results if r.executed]
        plan = profile.topic_plans.get(tid)
        priority = {"I": "Priority-I", "II": "Priority-II"}.get(topic.priority, "Company-specific")

        if executed:
            sources = []
            for r in executed:
                for ds in r.test.datasets:
                    src = profile.datasets.get(ds)
                    label = (src.sap_report if src and src.sap_report else
                             lib.datasets[ds].sap_sources[0] if lib.datasets[ds].sap_sources else ds)
                    if label not in sources:
                        sources.append(label)
            data_source = "; ".join(sources) + "; Remote"
            if len(executed) == 1:
                r = executed[0]
                exceptions = "Nil" if r.status == NIL else f"{r.n_exceptions} {r.test.exception_unit}"
                value = lakh(r.value) if r.status == EXCEPTIONS else None
            else:
                exceptions = "; ".join(("Nil" if r.status == NIL else f"{r.n_exceptions} {r.test.exception_unit}")
                                       + f" ({r.test.id})" for r in executed)
                value = "; ".join(f"{lakh(r.value) if r.status == EXCEPTIONS else '-'} ({r.test.id})" for r in executed)
            total_exceptions = sum(r.n_exceptions for r in executed)
        else:
            data_source = "On-site"
            exceptions = "To be analysed on-site"
            value = None
            total_exceptions = None

        detail = []
        if plan and plan.remarks:
            detail.append(plan.remarks)
        elif executed and total_exceptions == 0:
            detail.append("No exceptions: analytical assurance on the process (OO-53 para 7(a)), subject to "
                          "completeness of the extract.")
        elif executed:
            conc = "; ".join(unit_summary(r) for r in executed if r.units)
            if conc:
                detail.append("Exceptions concentrated in " + conc)
        else:
            reasons = [f"{r.test.id}: {r.message}" for r in results if r.status == NOT_EXECUTABLE]
            detail.append(topic.note or ("No library test yet; analysis on-site." if not results else ""))
            detail.extend(reasons)
        if topic.approval:
            detail.append("Approvals / sanctions not verifiable remotely (OO-53 para 10).")

        units = plan.units if plan and plan.units else "; ".join(
            u for r in executed for u, _ in r.units[:3])
        records = "; ".join(dict.fromkeys(r.test.records_required for r in results if r.test.records_required))
        rows.append({
            "topic": topic,
            "bp": lib.business_processes.get(topic.bp, ""),
            "title": f"{topic.title} ({priority}) [{topic.id}]",
            "reason_codes": plan.reason_codes if plan else "",
            "data_source": data_source,
            "exceptions": exceptions,
            "value": value,
            "recommend": plan.recommend if plan else "",
            "detail": "; ".join(d.rstrip(".") for d in detail if d) + ".",
            "units": units,
            "records": records,
            "executed": bool(executed),
            "total_exceptions": total_exceptions,
        })
    return rows


def annexure_checks(run: Run, rows: list[dict]) -> list[str]:
    notes = []
    bps = sorted({r["topic"].bp for r in rows})
    if len(rows) != 16:
        notes.append(f"{len(rows)} topics listed; OO-53 para 6 expects sixteen (four topics in each of four "
                     f"business processes), unless fewer are justified.")
    if len(bps) < 4 and not run.profile.fewer_bp_reason:
        notes.append(f"{len(bps)} business process(es) taken up; record the reasons where fewer than four "
                     f"(set annexure_ii.fewer_business_processes_reason in the profile).")
    recommended = [r for r in rows if str(r["recommend"]).upper() == "Y"]
    if recommended:
        if len(recommended) != 8:
            notes.append(f"{len(recommended)} topics recommended for field verification; ordinarily eight (para 7).")
        if len({r['topic'].bp for r in recommended}) < 2:
            notes.append("Recommended topics cover fewer than two business processes (para 7).")
        onsite = [r for r in recommended if not r["executed"]]
        if onsite:
            notes.append(f"{len(onsite)} recommended topic(s) need on-site analysis; keep these to the minimum in "
                         f"the first year and justify on assessed risk (para 7(b)).")
    else:
        notes.append("Column 'Recommended for Field Verification' is for the Lead Audit Team to fill.")
    return notes


def write_annexure_ii(run: Run, out_dir: Path) -> Path:
    from openpyxl.worksheet.page import PageMargins

    p = run.profile
    rows = annexure_rows(run)
    wb = Workbook()
    ws = wb.active
    ws.title = "Annexure-II"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_margins = PageMargins(left=0.3, right=0.3, top=0.5, bottom=0.5)
    for i, w in enumerate(ANNEX_WIDTHS, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws["A1"] = "Annexure-II"
    ws["A1"].font = BOLD
    ws["A2"] = "Company-wise Data-Driven Audit Plan"
    ws["A2"].font = TITLE
    ws["A3"] = (f"Company: {p.company}{' / ' + p.unit if p.unit else ''}    Tier: {p.tier}    "
                f"Lead Audit Team: {p.lead_team}    Period of Audit: {p.period_from:%d.%m.%Y} to {p.period_to:%d.%m.%Y}")
    bps = list(dict.fromkeys(r["bp"] for r in rows))
    ws["A4"] = (f"Business Processes taken up: {', '.join(bps)}"
                + (f"    (Where fewer than four, reasons: {p.fewer_bp_reason})" if len(bps) < 4 else ""))
    for cell in ("A3", "A4"):
        ws[cell].alignment = Alignment(wrap_text=False)

    header_row = 6
    for i, h in enumerate(ANNEX_HEADERS, 1):
        c = ws.cell(row=header_row, column=i, value=h)
        c.font, c.fill, c.alignment, c.border = BOLD, HEADER_FILL, Alignment(wrap_text=True, vertical="center"), BOX
    ws.freeze_panes = ws.cell(row=header_row + 1, column=1)

    for n, r in enumerate(rows, 1):
        values = [n, r["bp"], r["title"], r["reason_codes"], r["data_source"], r["exceptions"], r["value"],
                  r["recommend"], r["detail"], r["units"], r["records"], ""]
        for i, v in enumerate(values, 1):
            c = ws.cell(row=header_row + n, column=i, value=v)
            c.alignment, c.border = WRAP, BOX
            if i == 7 and isinstance(v, float):
                c.number_format = "#,##0.00"
            if i == 6 and r["total_exceptions"]:
                c.fill = FLAG_FILL
            elif i == 6 and r["total_exceptions"] == 0:
                c.fill = OK_FILL

    row = header_row + len(rows) + 2
    ws.cell(row=row, column=1, value="Checks against OO-53").font = BOLD
    for note in annexure_checks(run, rows):
        row += 1
        ws.cell(row=row, column=1, value=f"- {note}")
    row += 2
    ws.cell(row=row, column=1, value="Reason-for-Selection Codes").font = BOLD
    for code, text in REASON_CODES:
        row += 1
        ws.cell(row=row, column=1, value=code)
        ws.cell(row=row, column=2, value=text)
    row += 2
    ws.cell(row=row, column=1, value=DISCLAIMER).alignment = Alignment(wrap_text=False)

    path = out_dir / f"AnnexureII_{_safe_name(p.company)}_{run.started:%Y%m%d_%H%M%S}.xlsx"
    wb.save(path)
    return path
