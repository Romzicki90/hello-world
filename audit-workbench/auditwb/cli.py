"""Command line entry point.

    python -m auditwb topics                    Annexure-I topics and the tests available for each
    python -m auditwb checklist PROFILE         which SAP extracts and columns the selected topics need
    python -m auditwb check PROFILE             load the extracts and report mapping / executability only
    python -m auditwb run PROFILE [--out DIR]   run the tests and write the working paper and Annexure-II
    python -m auditwb demo DIR                  generate synthetic SAP extracts and run them end to end
"""

from __future__ import annotations

import argparse
import shutil
import sys
from pathlib import Path

from . import __version__
from .engine import check_executable, resolve_params, run, select_tests
from .loader import load_library
from .profile import load_profile
from .report import write_annexure_ii, write_working_paper

TEMPLATE = Path(__file__).parent / "library" / "TEMPLATE_profile.toml"


def cmd_topics(args, lib):
    for bp, name in lib.business_processes.items():
        print(f"\nBP{bp} {name}")
        for t in lib.topics.values():
            if t.bp != bp:
                continue
            tests = ", ".join(x.id for x in lib.tests_for_topic(t.id)) or "-"
            flags = " [on-site]" if t.mode == "On-site" else ""
            flags += " [approval]" if t.approval else ""
            print(f"  {t.id}  P-{t.priority:<2} {t.title}{flags}\n         tests: {tests}")


def cmd_checklist(args, lib):
    profile = load_profile(Path(args.profile))
    tests = select_tests(profile, lib)
    needed: dict[str, set[str]] = {}
    for t in tests:
        for ref in t.requires + [r for g in t.requires_any for r in g]:
            ds, _, fld = ref.partition(".")
            needed.setdefault(ds, set()).add(fld)
    print(f"Extraction checklist for {profile.company}: {len(tests)} tests")
    for ds, fields in needed.items():
        d = lib.datasets[ds]
        print(f"\n{ds}: {d.title}\n  SAP: {'; '.join(d.sap_sources)}\n  {d.description}")
        for fname, f in d.fields.items():
            mark = "needed" if fname in fields else "useful"
            print(f"    [{mark}] {f.label}  (e.g. {', '.join(f.synonyms[:3])})")


def cmd_check(args, lib):
    from .engine import DatasetStatus
    from .ingest import IngestError, load_dataset

    profile = load_profile(Path(args.profile))
    tests = select_tests(profile, lib)
    statuses = {}
    for ds_id, src in profile.datasets.items():
        st = DatasetStatus(id=ds_id, title=lib.datasets[ds_id].title)
        statuses[ds_id] = st
        try:
            st.load = load_dataset(lib.datasets[ds_id], src.files, sheet=src.sheet, explicit=src.columns,
                                   decimal_notation=profile.decimal_notation, date_format=profile.date_format)
            print(f"{ds_id}: {sum(f.rows_loaded for f in st.load.files)} rows; mapped {len(st.load.mapped_fields)}"
                  f"/{len(lib.datasets[ds_id].fields)} fields")
            for f in st.load.files:
                for k, v in f.mapping.items():
                    print(f"    {k:<20} <- {v}")
            for w in st.load.warnings:
                print(f"    WARNING: {w}")
        except (IngestError, OSError) as exc:
            print(f"{ds_id}: NOT LOADED: {exc}")
    print()
    for t in tests:
        missing = check_executable(t, statuses, resolve_params(t, profile))
        print(f"{t.id:<18} {'ready' if not missing else 'NOT EXECUTABLE: ' + '; '.join(missing)}")


def cmd_run(args, lib):
    profile = load_profile(Path(args.profile))
    out = Path(args.out) if args.out else profile.path.parent / "output"
    out.mkdir(parents=True, exist_ok=True)
    result = run(profile, lib)
    wp = write_working_paper(result, out)
    ax = write_annexure_ii(result, out)
    print()
    for r in result.results:
        extra = f"{r.n_exceptions} {r.test.exception_unit}" if r.executed else r.message
        print(f"{r.test.id:<18} {r.status:<15} {extra}")
    print(f"\nWorking paper: {wp}\nAnnexure-II:   {ax}")


def cmd_demo(args, lib):
    from .synthetic import generate

    target = Path(args.dir)
    profile_path = generate(target)
    print(f"Synthetic SAP extracts written to {target}")
    args.profile, args.out = str(profile_path), None
    cmd_run(args, lib)


def cmd_init(args, lib):
    dest = Path(args.path)
    if dest.exists():
        sys.exit(f"{dest} already exists")
    shutil.copy(TEMPLATE, dest)
    print(f"Profile template written to {dest}")


def main(argv=None):
    ap = argparse.ArgumentParser(prog="auditwb", description=f"Audit Analytics Workbench {__version__} (offline)")
    sub = ap.add_subparsers(dest="cmd", required=True)
    sub.add_parser("topics", help="list Annexure-I topics and available tests")
    p = sub.add_parser("checklist", help="SAP extracts and columns needed for a profile")
    p.add_argument("profile")
    p = sub.add_parser("check", help="load extracts and show mapping / executability without running tests")
    p.add_argument("profile")
    p = sub.add_parser("run", help="run tests and write the working paper and Annexure-II")
    p.add_argument("profile")
    p.add_argument("--out", help="output folder (default: 'output' next to the profile)")
    p = sub.add_parser("demo", help="generate synthetic SAP extracts in DIR and run them")
    p.add_argument("dir")
    p = sub.add_parser("init", help="write a blank run profile")
    p.add_argument("path")
    sub.add_parser("gui", help="open the point-and-click window")
    args = ap.parse_args(argv)
    if args.cmd == "gui":
        from .gui import main as gui_main
        return gui_main()
    lib = load_library()
    {"topics": cmd_topics, "checklist": cmd_checklist, "check": cmd_check, "run": cmd_run,
     "demo": cmd_demo, "init": cmd_init}[args.cmd](args, lib)


if __name__ == "__main__":
    main()
