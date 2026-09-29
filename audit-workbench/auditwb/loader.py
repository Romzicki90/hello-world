"""Loads the standard datasets, the Annexure-I catalogue and the analytic test library."""

from __future__ import annotations

import re
import tomllib
from dataclasses import dataclass, field
from pathlib import Path

LIBRARY_DIR = Path(__file__).parent / "library"

FIELD_TYPES = {"id", "text", "date", "amount", "qty"}
GLOBAL_PARAMS = {"period_from", "period_to", "cutoff_date"}


def normalise_header(text) -> str:
    """Header comparison key: lower case, letters and digits only."""
    return re.sub(r"[^a-z0-9]", "", str(text).lower())


@dataclass
class Field:
    name: str
    type: str
    label: str
    synonyms: list[str]
    key: bool = False
    sign_by: str | None = None
    strip_prefixes: list[str] = field(default_factory=list)
    keep_zeros: bool = False

    @property
    def sql_type(self) -> str:
        return {"date": "DATE", "amount": "DOUBLE", "qty": "DOUBLE"}.get(self.type, "VARCHAR")


@dataclass
class Dataset:
    id: str
    title: str
    sap_sources: list[str]
    description: str
    fields: dict[str, Field]
    unique: list[str] = field(default_factory=list)       # key that must not repeat (master data)
    no_overlap: dict | None = None                       # {key=[...], from=..., to=...} validity periods
    control_total: str | None = None                     # field reconciled to the SAP report total
    criteria: bool = False                               # audit-supplied criteria, not an SAP extract


@dataclass
class Topic:
    id: str
    bp: int
    priority: str
    serial: int
    title: str
    mode: str
    approval: bool = False
    note: str = ""
    withdrawn: bool = False

    @property
    def annex_ref(self) -> str:
        if self.priority == "CS":
            return f"Company-specific {self.id}"
        return f"BP{self.bp} Priority-{self.priority} Sl.{self.serial}"


@dataclass
class Param:
    name: str
    label: str
    default: object = None
    required: bool = False
    help: str = ""


@dataclass
class Test:
    id: str
    title: str
    topics: list[str]
    exception_unit: str
    requires: list[str]
    requires_any: list[list[str]]
    rule: str
    interpretation: str
    field_questions: list[str]
    records_required: str
    population_sql: str
    sql: str
    params: dict[str, Param]
    source_file: str
    uses: list[str] = field(default_factory=list)   # optional datasets joined if loaded
    coverage_sql: str = ""                          # rows of (metric, n, of_n)

    @property
    def datasets(self) -> list[str]:
        names = [r.split(".")[0] for r in self.requires]
        names += [r.split(".")[0] for group in self.requires_any for r in group]
        names += self.uses
        return list(dict.fromkeys(names))


@dataclass
class Library:
    datasets: dict[str, Dataset]
    catalogue_version: str
    business_processes: dict[int, str]
    topics: dict[str, Topic]
    tests: dict[str, Test]

    def tests_for_topic(self, topic_id: str) -> list[Test]:
        return [t for t in self.tests.values() if topic_id in t.topics]


def load_datasets(path: Path) -> dict[str, Dataset]:
    raw = tomllib.loads(path.read_text(encoding="utf-8"))
    datasets = {}
    for ds_id, spec in raw.items():
        fields = {}
        for name, f in spec["fields"].items():
            if f["type"] not in FIELD_TYPES:
                raise ValueError(f"{ds_id}.{name}: unknown field type {f['type']!r}")
            fields[name] = Field(
                name=name,
                type=f["type"],
                label=f.get("label", name),
                synonyms=f.get("synonyms", []),
                key=f.get("key", False),
                sign_by=f.get("sign_by"),
                strip_prefixes=f.get("strip_prefixes", []),
                keep_zeros=f.get("keep_zeros", False),
            )
        datasets[ds_id] = Dataset(
            id=ds_id,
            title=spec["title"],
            sap_sources=spec.get("sap_sources", []),
            description=spec.get("description", ""),
            fields=fields,
            unique=spec.get("unique", []),
            no_overlap=spec.get("no_overlap"),
            control_total=spec.get("control_total"),
            criteria=spec.get("criteria", False),
        )
        ds = datasets[ds_id]
        for name in ds.unique + ([ds.control_total] if ds.control_total else []):
            if name not in fields:
                raise ValueError(f"{ds_id}: unknown field {name!r} in unique / control_total")
        if ds.no_overlap:
            for name in ds.no_overlap["key"] + [ds.no_overlap["from"], ds.no_overlap["to"]]:
                if name not in fields:
                    raise ValueError(f"{ds_id}: unknown field {name!r} in no_overlap")
    return datasets


def load_catalogue(path: Path):
    raw = tomllib.loads(path.read_text(encoding="utf-8"))
    bps = {int(k): v for k, v in raw["business_processes"].items()}
    topics = {}
    for t in raw["topics"]:
        if t["id"] in topics:
            raise ValueError(f"Duplicate topic id {t['id']}")
        topics[t["id"]] = Topic(**t)
    return raw["catalogue_version"], bps, topics


def load_tests(directory: Path) -> dict[str, Test]:
    tests = {}
    for path in sorted(directory.glob("*.toml")):
        raw = tomllib.loads(path.read_text(encoding="utf-8"))
        for t in raw.get("test", []):
            if t["id"] in tests:
                raise ValueError(f"Duplicate test id {t['id']} in {path.name}")
            params = {
                name: Param(name=name, label=p.get("label", name), default=p.get("default"),
                            required=p.get("required", False), help=p.get("help", ""))
                for name, p in t.get("params", {}).items()
            }
            tests[t["id"]] = Test(
                id=t["id"],
                title=t["title"],
                topics=t["topics"],
                exception_unit=t.get("exception_unit", "exceptions"),
                requires=t.get("requires", []),
                requires_any=t.get("requires_any", []),
                rule=" ".join(t["rule"].split()),
                interpretation=" ".join(t.get("interpretation", "").split()),
                field_questions=t.get("field_questions", []),
                records_required=t.get("records_required", ""),
                population_sql=t.get("population_sql", ""),
                sql=t["sql"],
                params=params,
                source_file=path.name,
                uses=t.get("uses", []),
                coverage_sql=t.get("coverage_sql", ""),
            )
    return tests


def sql_parameters(sql: str) -> set[str]:
    return set(re.findall(r"\$([A-Za-z_]\w*)", sql))


def validate(lib: Library) -> list[str]:
    """Consistency checks on the library itself; returns a list of problems."""
    problems = []
    for t in lib.tests.values():
        for topic in t.topics:
            if topic not in lib.topics:
                problems.append(f"{t.id}: unknown topic {topic}")
        for ref in t.requires + [r for g in t.requires_any for r in g]:
            ds, _, fld = ref.partition(".")
            if ds not in lib.datasets or fld not in lib.datasets[ds].fields:
                problems.append(f"{t.id}: requires unknown field {ref}")
        for ds in t.uses:
            if ds not in lib.datasets:
                problems.append(f"{t.id}: uses unknown dataset {ds}")
        for sql in (t.sql, t.population_sql, t.coverage_sql):
            for name in sql_parameters(sql) - GLOBAL_PARAMS:
                if name not in t.params:
                    problems.append(f"{t.id}: SQL uses ${name} but no such parameter is declared")
    return problems


def load_library(directory: Path = LIBRARY_DIR) -> Library:
    version, bps, topics = load_catalogue(directory / "catalogue.toml")
    lib = Library(
        datasets=load_datasets(directory / "datasets.toml"),
        catalogue_version=version,
        business_processes=bps,
        topics=topics,
        tests=load_tests(directory / "tests"),
    )
    problems = validate(lib)
    if problems:
        raise ValueError("Test library is inconsistent:\n  " + "\n  ".join(problems))
    return lib
