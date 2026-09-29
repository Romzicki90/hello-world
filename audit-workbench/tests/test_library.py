"""Every analytic test must find exactly the exceptions planted in the synthetic data:
all planted exceptions, and none of the near-misses or background transactions."""

import pytest

from auditwb.engine import EXCEPTIONS, NOT_EXECUTABLE, run
from auditwb.loader import load_library
from auditwb.profile import load_profile
from auditwb.report import annexure_rows, write_annexure_ii, write_working_paper
from auditwb.synthetic import PLANTED, generate


@pytest.fixture(scope="module")
def demo_run(tmp_path_factory):
    target = tmp_path_factory.mktemp("demo")
    profile = load_profile(generate(target))
    return run(profile, load_library(), progress=lambda *_: None)


def test_every_library_test_has_planted_cases():
    assert set(PLANTED) == set(load_library().tests)


@pytest.mark.parametrize("test_id", sorted(PLANTED))
def test_finds_exactly_the_planted_exceptions(demo_run, test_id):
    result = demo_run.result(test_id)
    assert result.status == EXCEPTIONS, result.message
    column, expected = PLANTED[test_id]
    found = {str(row[result.columns.index(column)]) for row in result.rows}
    assert found == expected


def test_population_and_value_are_reported(demo_run):
    for r in demo_run.results:
        assert r.population_n, f"{r.test.id}: population not reported"
        assert r.value is not None and r.value > 0, f"{r.test.id}: value involved missing"


def test_split_po_cluster_value(demo_run):
    r = demo_run.result("PRC-SPLIT-01")
    assert r.n_exceptions == 1
    assert r.value == pytest.approx(1_050_000)


def test_ingest_drops_subtotal_and_total_lines(demo_run):
    load = demo_run.datasets["vendor_items"].load
    dropped = sum(sum(f.dropped.values()) for f in load.files)
    assert dropped >= 2
    assert "vendor" in load.mapped_fields and "amount" in load.mapped_fields


def test_missing_data_is_not_executable_not_nil(tmp_path):
    profile = load_profile(generate(tmp_path))
    del profile.datasets["grir"]
    profile.params["FIN-SENSGL-01"] = {"sensitive_gls": []}
    result = run(profile, load_library(), progress=lambda *_: None)
    assert result.result("PRC-GRIR-01").status == NOT_EXECUTABLE
    assert "grir" in result.result("PRC-GRIR-01").message
    assert result.result("FIN-SENSGL-01").status == NOT_EXECUTABLE


def test_settings_saved_from_window_reload_identically(tmp_path):
    from auditwb.profile import dump_toml, read_raw

    raw = read_raw(generate(tmp_path / "demo"))
    raw["params"]["PRC-SPLIT-01"] = {"threshold": 2000000}
    saved = tmp_path / "saved" / "audit_settings.toml"
    saved.parent.mkdir()
    saved.write_text(dump_toml(raw), encoding="utf-8")
    assert read_raw(saved) == raw
    result = run(load_profile(saved), load_library(), progress=lambda *_: None)
    assert result.result("PRC-SPLIT-01").n_exceptions == 0   # 10.5 lakh cluster is below 20 lakh


def test_outputs_are_written(demo_run, tmp_path):
    wp = write_working_paper(demo_run, tmp_path)
    ax = write_annexure_ii(demo_run, tmp_path)
    assert wp.exists() and ax.exists()
    rows = {r["topic"].id: r for r in annexure_rows(demo_run)}
    assert rows["T1A01"]["exceptions"] == "1 clusters"
    assert rows["T1A01"]["value"] == pytest.approx(10.5)
    assert rows["T9A01"]["exceptions"] == "To be analysed on-site"
