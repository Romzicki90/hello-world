"""The run profile: one TOML file per company audit, filled in by the audit team.

It says which company and period is being audited, where each SAP extract is, how
columns map (only where automatic mapping fails), the SAP number / date settings, the
topics selected for Annexure-II and any parameter changes (thresholds, criteria lists).
See profiles/TEMPLATE.toml.
"""

from __future__ import annotations

import datetime as dt
import tomllib
from dataclasses import dataclass, field
from pathlib import Path


@dataclass
class DatasetSource:
    files: list[Path]
    sheet: str | None = None
    columns: dict[str, str] = field(default_factory=dict)
    extracted_on: dt.date | None = None
    sap_report: str = ""


@dataclass
class TopicPlan:
    reason_codes: str = ""
    units: str = ""
    remarks: str = ""
    recommend: str = ""


@dataclass
class Profile:
    path: Path
    company: str
    unit: str
    tier: str
    lead_team: str
    period_from: dt.date
    period_to: dt.date
    cutoff_date: dt.date
    decimal_notation: str
    date_format: str
    datasets: dict[str, DatasetSource]
    topics: list[str]
    topic_plans: dict[str, TopicPlan]
    tests: list[str] | None
    params: dict[str, dict]
    fewer_bp_reason: str = ""


def _date(value, name) -> dt.date:
    if isinstance(value, dt.date):
        return value
    try:
        return dt.date.fromisoformat(str(value))
    except ValueError:
        raise ValueError(f"profile: {name} must be a date written as YYYY-MM-DD") from None


def load_profile(path: Path) -> Profile:
    path = Path(path)
    raw = tomllib.loads(path.read_text(encoding="utf-8"))
    base = path.parent
    audit = raw.get("audit", {})
    sap = raw.get("sap", {})
    period_from = _date(audit.get("period_from"), "audit.period_from")
    period_to = _date(audit.get("period_to"), "audit.period_to")

    datasets = {}
    for ds_id, spec in raw.get("datasets", {}).items():
        files = spec.get("files") or ([spec["file"]] if "file" in spec else [])
        datasets[ds_id] = DatasetSource(
            files=[(base / f).resolve() for f in files],
            sheet=spec.get("sheet") or None,
            columns=spec.get("columns", {}),
            extracted_on=_date(spec["extracted_on"], f"datasets.{ds_id}.extracted_on") if spec.get("extracted_on") else None,
            sap_report=spec.get("sap_report", ""),
        )

    annex = raw.get("annexure_ii", {})
    plans = {tid: TopicPlan(reason_codes=p.get("reason_codes", ""), units=p.get("units", ""),
                            remarks=p.get("remarks", ""), recommend=p.get("recommend", ""))
             for tid, p in annex.get("topic", {}).items()}

    return Profile(
        path=path,
        company=audit.get("company", "(company not set)"),
        unit=audit.get("unit", ""),
        tier=audit.get("tier", "A"),
        lead_team=audit.get("lead_team", ""),
        period_from=period_from,
        period_to=period_to,
        cutoff_date=_date(audit.get("cutoff_date", period_to), "audit.cutoff_date"),
        decimal_notation=sap.get("decimal_notation", "1,234,567.89"),
        date_format=sap.get("date_format", "DD.MM.YYYY"),
        datasets=datasets,
        topics=annex.get("topics", []),
        topic_plans=plans,
        tests=raw.get("run", {}).get("tests"),
        params=raw.get("params", {}),
        fewer_bp_reason=annex.get("fewer_business_processes_reason", ""),
    )
