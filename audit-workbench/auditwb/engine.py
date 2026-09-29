"""Runs the selected analytic tests on the loaded SAP extracts.

Everything happens on this machine: extracts are parsed in Python, written to a
temporary folder that is deleted afterwards, and queried with an in-memory DuckDB
database with extension auto-install switched off (no network access is attempted).
"""

from __future__ import annotations

import csv
import datetime as dt
import getpass
import platform
import tempfile
from collections import defaultdict
from dataclasses import dataclass, field
from pathlib import Path

import duckdb

from . import __version__
from .ingest import DatasetLoad, IngestError, load_dataset
from .loader import Library, Test, sql_parameters
from .profile import Profile

REQUIRED_RESULT_COLUMNS = ("exception_id", "exception_value", "reason")

EXCEPTIONS = "EXCEPTIONS"
NIL = "NIL"
NOT_EXECUTABLE = "NOT EXECUTABLE"
ERROR = "ERROR"


@dataclass
class TestResult:
    test: Test
    status: str
    message: str = ""
    params: dict = field(default_factory=dict)
    columns: list[str] = field(default_factory=list)
    rows: list[tuple] = field(default_factory=list)
    n_exceptions: int = 0
    value: float | None = None
    population_n: int | None = None
    population_value: float | None = None
    units: list[tuple[str, float]] = field(default_factory=list)

    @property
    def executed(self) -> bool:
        return self.status in (EXCEPTIONS, NIL)


@dataclass
class DatasetStatus:
    id: str
    title: str
    load: DatasetLoad | None = None
    error: str = ""
    control_totals: dict = field(default_factory=dict)
    date_ranges: dict = field(default_factory=dict)


@dataclass
class Run:
    profile: Profile
    library: Library
    datasets: dict[str, DatasetStatus]
    results: list[TestResult]
    started: dt.datetime
    finished: dt.datetime
    tool_version: str = __version__
    run_by: str = ""
    machine: str = ""

    def result(self, test_id: str) -> TestResult:
        return next(r for r in self.results if r.test.id == test_id)


def _connect() -> duckdb.DuckDBPyConnection:
    return duckdb.connect(":memory:", config={
        "autoinstall_known_extensions": False,
        "autoload_known_extensions": False,
    })


def _create_table(con, ds_id: str, dataset, rows: list[dict], tmp: Path):
    cols = list(dataset.fields.values())
    col_defs = ", ".join(f'"{f.name}" {f.sql_type}' for f in cols)
    con.execute(f'CREATE TABLE "{ds_id}" ({col_defs})')
    if not rows:
        return
    path = tmp / f"{ds_id}.csv"
    with open(path, "w", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh)
        w.writerow([f.name for f in cols])
        for r in rows:
            w.writerow(["" if r.get(f.name) is None else
                        (r[f.name].isoformat() if isinstance(r[f.name], dt.date) else r[f.name])
                        for f in cols])
    struct = ", ".join(f"'{f.name}': '{f.sql_type}'" for f in cols)
    literal_path = str(path).replace("'", "''")
    con.execute(f"""INSERT INTO "{ds_id}" SELECT * FROM read_csv('{literal_path}', header = true,
                    auto_detect = false, columns = {{{struct}}}, nullstr = '', quote = '"', escape = '"')""")
    path.unlink()


def _profile_dataset(con, ds_id: str, dataset, mapped: set[str]) -> tuple[dict, dict]:
    totals, ranges = {}, {}
    for f in dataset.fields.values():
        if f.name not in mapped:
            continue
        if f.type in ("amount", "qty"):
            s, neg = con.execute(
                f'SELECT sum("{f.name}"), count(*) FILTER (WHERE "{f.name}" < 0) FROM "{ds_id}"').fetchone()
            totals[f.name] = (s, neg)
        elif f.type == "date":
            lo, hi, nulls = con.execute(
                f'SELECT min("{f.name}"), max("{f.name}"), count(*) FILTER (WHERE "{f.name}" IS NULL) FROM "{ds_id}"').fetchone()
            ranges[f.name] = (lo, hi, nulls)
    return totals, ranges


def resolve_params(test: Test, profile: Profile) -> dict:
    values = {name: p.default for name, p in test.params.items()}
    common = profile.params.get("common", {})
    values.update({k: v for k, v in common.items() if k in test.params})
    values.update(profile.params.get(test.id, {}))
    return values


def check_executable(test: Test, datasets: dict[str, DatasetStatus], params: dict) -> list[str]:
    missing = []

    def available(ref: str) -> bool:
        ds, _, fld = ref.partition(".")
        st = datasets.get(ds)
        return bool(st and st.load and fld in st.load.mapped_fields)

    not_loaded = sorted({r.split(".")[0] for r in test.requires
                         if not (datasets.get(r.split(".")[0]) and datasets[r.split(".")[0]].load)})
    if not_loaded:
        missing.append("dataset(s) not loaded: " + ", ".join(not_loaded))
    missing_fields = [r for r in test.requires if r.split(".")[0] not in not_loaded and not available(r)]
    if missing_fields:
        missing.append("field(s) not in extract: " + ", ".join(missing_fields))
    for group in test.requires_any:
        if not any(available(r) for r in group):
            missing.append("needs one of: " + " / ".join(group))
    for name, p in test.params.items():
        v = params.get(name)
        if p.required and (v is None or v == [] or v == ""):
            missing.append(f"parameter '{name}' ({p.label}) must be set in the profile")
    return missing


def _run_test(con, test: Test, profile: Profile, datasets) -> TestResult:
    params = resolve_params(test, profile)
    missing = check_executable(test, datasets, params)
    if missing:
        return TestResult(test=test, status=NOT_EXECUTABLE, message="; ".join(missing), params=params)

    bind = dict(params, period_from=profile.period_from, period_to=profile.period_to,
                cutoff_date=profile.cutoff_date)
    try:
        if test.population_sql.strip():
            names = sql_parameters(test.population_sql)
            pop = con.execute(test.population_sql, {k: bind[k] for k in names}).fetchone()
            population_n, population_value = (pop[0], pop[1]) if pop else (None, None)
        else:
            population_n = population_value = None
        names = sql_parameters(test.sql)
        cur = con.execute(test.sql, {k: bind[k] for k in names})
        columns = [d[0] for d in cur.description]
        rows = cur.fetchall()
    except duckdb.Error as exc:
        return TestResult(test=test, status=ERROR, message=str(exc).splitlines()[0], params=params)

    absent = [c for c in REQUIRED_RESULT_COLUMNS if c not in columns]
    if absent:
        return TestResult(test=test, status=ERROR, params=params,
                          message=f"test SQL does not return required column(s): {', '.join(absent)}")

    i_id, i_val = columns.index("exception_id"), columns.index("exception_value")
    i_unit = columns.index("audit_unit") if "audit_unit" in columns else None
    ids = {r[i_id] for r in rows}
    vals = [r[i_val] for r in rows if r[i_val] is not None]
    by_unit = defaultdict(float)
    if i_unit is not None:
        for r in rows:
            if r[i_unit]:
                by_unit[r[i_unit]] += float(r[i_val] or 0)
    return TestResult(
        test=test,
        status=EXCEPTIONS if rows else NIL,
        params=params,
        columns=columns,
        rows=rows,
        n_exceptions=len(ids),
        value=float(sum(vals)) if vals else (0.0 if not rows else None),
        population_n=population_n,
        population_value=float(population_value) if population_value is not None else None,
        units=sorted(by_unit.items(), key=lambda kv: -kv[1]),
    )


def select_tests(profile: Profile, lib: Library) -> list[Test]:
    if profile.tests:
        unknown = [t for t in profile.tests if t not in lib.tests]
        if unknown:
            raise ValueError(f"profile [run] tests: unknown test id(s) {', '.join(unknown)}")
        return [lib.tests[t] for t in profile.tests]
    if profile.topics:
        unknown = [t for t in profile.topics if t not in lib.topics]
        if unknown:
            raise ValueError(f"profile annexure_ii topics: unknown topic id(s) {', '.join(unknown)}")
        chosen = [t for t in lib.tests.values() if set(t.topics) & set(profile.topics)]
        return chosen
    return list(lib.tests.values())


def run(profile: Profile, lib: Library, progress=print) -> Run:
    started = dt.datetime.now()
    tests = select_tests(profile, lib)
    needed = {ds for t in tests for ds in t.datasets} | set(profile.datasets)
    con = _connect()
    statuses: dict[str, DatasetStatus] = {}
    with tempfile.TemporaryDirectory(prefix="auditwb_") as tmpdir:
        tmp = Path(tmpdir)
        for ds_id, dataset in lib.datasets.items():
            st = DatasetStatus(id=ds_id, title=dataset.title)
            statuses[ds_id] = st
            src = profile.datasets.get(ds_id)
            rows = []
            if src and ds_id in needed:
                missing_files = [str(f) for f in src.files if not f.exists()]
                if not src.files:
                    st.error = "no file given in the profile"
                elif missing_files:
                    st.error = "file not found: " + ", ".join(missing_files)
                else:
                    progress(f"Loading {ds_id} from {', '.join(f.name for f in src.files)}")
                    try:
                        st.load = load_dataset(dataset, src.files, sheet=src.sheet, explicit=src.columns,
                                               decimal_notation=profile.decimal_notation,
                                               date_format=profile.date_format)
                        rows = st.load.rows
                    except IngestError as exc:
                        st.error = str(exc)
            _create_table(con, ds_id, dataset, rows, tmp)
            if st.load:
                st.control_totals, st.date_ranges = _profile_dataset(con, ds_id, dataset, st.load.mapped_fields)
                st.load.rows = []  # data now lives only in the in-memory database

        results = []
        for test in tests:
            progress(f"Running {test.id}: {test.title}")
            results.append(_run_test(con, test, profile, statuses))
    con.close()
    try:
        user = getpass.getuser()
    except Exception:
        user = ""
    return Run(profile=profile, library=lib, datasets=statuses, results=results, started=started,
               finished=dt.datetime.now(), run_by=user, machine=platform.node())
