"""End-to-end checks on the synthetic demo tender (examples/demo-tender).

The private development suite ran against real, anonymised tender files that
cannot be published; these tests exercise the same pipeline on invented data.
"""

import json
import subprocess
import sys
from pathlib import Path

import openpyxl
import pytest

ROOT = Path(__file__).resolve().parents[1]
DEMO = ROOT / "examples" / "demo-tender"
sys.path.insert(0, str(ROOT))


@pytest.fixture(scope="session", autouse=True)
def demo_files():
    subprocess.run([sys.executable, str(ROOT / "examples" / "make_demo_tender.py")], check=True)
    return DEMO


@pytest.fixture(scope="session")
def run(tmp_path_factory):
    from boq.pipeline import fill_pipeline

    out = tmp_path_factory.mktemp("runs")
    quotes = sorted(str(p) for p in (DEMO / "quotes").glob("*.xlsx"))
    return fill_pipeline(str(DEMO / "boq-r0.xlsx"), quotes, markup_pct=12, area_sqft=25000,
                         out_dir=str(out), kb_dir=str(out / "kb"))


def _state(run):
    return json.loads((Path(run["run_dir"]) / "state.json").read_text())


class TestExtraction:
    def test_boq_items_and_columns(self):
        from boq.extract import extract_workbook, validate_extraction
        from boq.ingest import load_workbook

        p = extract_workbook(load_workbook(DEMO / "boq-r0.xlsx"))
        assert validate_extraction(p) == []
        boq = next(s for s in p["sheets"] if s["kind"] == "boq")
        assert len(boq["items"]) == 8

    def test_supply_install_split_is_summed(self):
        from boq.extract import extract_workbook
        from boq.ingest import load_workbook

        p = extract_workbook(load_workbook(DEMO / "quotes" / "vendor-a-floors-glass.xlsx"))
        s = p["sheets"][0]
        assert s["supply_install_split"] is True
        for it in s["items"]:
            assert it["total_rate"] == (it["supply_rate"] or 0) + (it["install_rate"] or 0)


class TestPricing:
    def test_everything_priced_nothing_silently_dropped(self, run):
        assert run["items_priced"] == 8
        assert run["review_queue"] == 0 and run["unmatched"] == 0

    def test_l1_is_lowest_total_rate(self, run):
        # glass partition: vendor A 4900+950=5850 vs vendor B 5620 -> B wins
        st = json.dumps(_state(run))
        assert "vendor-b-ceilings-partitions" in st
        assert abs(run["totals"]["cost_basis"] - 8_432_050.0) < 0.01

    def test_markup_round_convention_and_margin(self, run):
        t = run["totals"]
        assert t["subtotal_ex_gst"] == pytest.approx(9_443_896.0)
        # margin as % of offer for a 12% markup = 12/112
        assert t["margin_pct"] == pytest.approx(10.71, abs=0.01)
        assert t["rate_per_sqft_ex_gst"] == pytest.approx(377.76, abs=0.01)

    def test_client_copy_values_only_format_kept(self, run):
        cc = openpyxl.load_workbook(run["outputs"]["client_copy"]["file"])
        ws = cc["BOQ"]
        assert ws["E6"].value == pytest.approx(2049.6)       # ROUND(1830*1.12, 2)
        assert not str(ws["E6"].value).startswith("=")     # no formulas leak
        assert ws["A4"].fill.fgColor.rgb.endswith("1F3864")  # header styling untouched


class TestRevisionsAndTarget:
    def test_revision_diff(self):
        from boq.diff import diff_files

        d = diff_files(str(DEMO / "boq-r0.xlsx"), str(DEMO / "boq-r1.xlsx")).to_dict()
        assert d["counts"] == {"unchanged": 6, "qty_changed": 1, "added": 1, "removed": 1}

    def test_target_back_solve_warns_below_margin_floor(self, run):
        from boq.pipeline import load_priced_state
        from boq.price import back_solve_target

        state = load_priced_state(str(Path(run["run_dir"]) / "state.json"))
        res = back_solve_target(state, 10_500_000, margin_floor_pct=8.0)
        assert res["feasible"]
        assert abs(res["achieved_total"] - 10_500_000) < 50
        assert res["margin_floor_warnings"]
