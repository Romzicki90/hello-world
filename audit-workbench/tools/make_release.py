"""Builds the ready-to-copy Windows folder (as a zip) for office machines without internet.

    python tools/make_release.py

Needs internet on the machine that builds the release (to fetch the Windows wheels),
never on the office machine. Output: dist/AuditWorkbench_v<version>.zip containing

    AuditWorkbench/
        START HERE.txt
        Install.bat
        Run Audit Workbench.bat
        auditwb/            the program and its test library
        wheels/             DuckDB and openpyxl for 64-bit Windows, Python 3.11 - 3.14
        docs/README.md      technical notes
"""

from __future__ import annotations

import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from auditwb import __version__  # noqa: E402

PYTHON_VERSIONS = ["3.11", "3.12", "3.13", "3.14"]
PACKAGES = ["duckdb>=1.0", "openpyxl>=3.1"]


def fetch_wheels(dest: Path, versions: list[str]):
    dest.mkdir(parents=True, exist_ok=True)
    for version in versions:
        subprocess.run([sys.executable, "-m", "pip", "download", *PACKAGES, "--only-binary=:all:",
                        "--platform", "win_amd64", "--python-version", version, "--implementation", "cp",
                        "-d", str(dest), "-q"], check=True)


def main():
    import argparse

    ap = argparse.ArgumentParser()
    ap.add_argument("--python", nargs="+", default=PYTHON_VERSIONS,
                    help="Windows Python versions to bundle wheels for (fewer = smaller zip)")
    versions = ap.parse_args().python
    build = ROOT / "build" / "AuditWorkbench"
    if build.exists():
        shutil.rmtree(build)
    build.mkdir(parents=True)

    shutil.copytree(ROOT / "auditwb", build / "auditwb",
                    ignore=shutil.ignore_patterns("__pycache__", "*.pyc"))
    for name in ("Install.bat", "Run Audit Workbench.bat", "START HERE.txt"):
        text = (ROOT / "windows" / name).read_text(encoding="utf-8")
        (build / name).write_bytes(text.replace("\r\n", "\n").replace("\n", "\r\n").encode("utf-8"))
    (build / "docs").mkdir()
    shutil.copy(ROOT / "README.md", build / "docs" / "README.md")

    cache = ROOT / "build" / "wheels"
    fetch_wheels(cache, versions)
    tags = {f"cp{v.replace('.', '')}" for v in versions}
    shutil.copytree(cache, build / "wheels", ignore=lambda _d, names: [
        n for n in names if n.startswith("duckdb-") and not any(f"-{t}-" in n for t in tags)])

    dist = ROOT / "dist"
    dist.mkdir(exist_ok=True)
    suffix = "" if versions == PYTHON_VERSIONS else "_py" + "-".join(v.replace(".", "") for v in versions)
    target = dist / f"AuditWorkbench_v{__version__}{suffix}.zip"
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as zf:
        for path in sorted(build.rglob("*")):
            zf.write(path, path.relative_to(build.parent))
    print(f"Release written to {target} ({target.stat().st_size / 1e6:.1f} MB)")


if __name__ == "__main__":
    main()
