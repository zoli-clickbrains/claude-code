import shutil
import subprocess

import pytest
from openpyxl import load_workbook

from icp.config import CRITERIA, GROUPS, SCALE_MAX, SEGMENTS, default_config
from icp.engine import criterion_points, score_client, score_clients
from icp.sample import SAMPLE_CLIENTS
from icp.workbook import FIRST, GATE_FAIL, GATE_OK, S_RANK, S_RES, ResCols, build_workbook

CFG = default_config()
BY_CODE = {c["kod"]: c for c in CRITERIA}
IDEAL_SEGMENT = SEGMENTS[0][0]


def _all(points, **extra):
    """Minden 1–10-es skálán ugyanaz a pont, plusz egyedi értékek."""
    client = {c["kod"]: points for c in CRITERIA if c["tipus"] == "skala"}
    return client | extra


def test_weights_sum_to_100_and_groups_are_complete():
    assert sum(c["suly"] for c in CRITERIA) == 100
    assert {c["csoport"] for c in CRITERIA} == {g for g, _ in GROUPS}
    assert sum(c["csoport"] == "Marketingezhetőség" for c in CRITERIA) == 17
    assert len({c["kod"] for c in CRITERIA}) == len(CRITERIA)


@pytest.mark.parametrize("value,expected", [(0, 1), (149_999, 1), (150_000, 3), (599_999, 5), (2_000_000, 10)])
def test_budget_bands(value, expected):
    assert criterion_points(BY_CODE["keret"], value, CFG) == expected


@pytest.mark.parametrize("days,expected", [(0, 10), (1, 8), (7, 8), (8, 6), (30, 4), (31, 2), (90, 1)])
def test_payment_delay_bands_are_descending(days, expected):
    assert criterion_points(BY_CODE["keses"], days, CFG) == expected


def test_scale_is_clamped_to_1_10():
    assert criterion_points(BY_CODE["emberi"], 14, CFG) == SCALE_MAX
    assert criterion_points(BY_CODE["emberi"], 0, CFG) == 1
    assert criterion_points(BY_CODE["emberi"], 6.5, CFG) == 7


def test_missing_and_unknown_values_get_missing_points():
    assert criterion_points(BY_CODE["roas"], "", CFG) == CFG["missing_points"]
    assert criterion_points(BY_CODE["szegmens"], "Ismeretlen", CFG) == CFG["missing_points"]
    assert criterion_points(BY_CODE["szegmens"], f" {IDEAL_SEGMENT.upper()} ", CFG) == 10


def test_perfect_and_worst_clients_hit_score_limits():
    best = _all(10, nev="X", szegmens=IDEAL_SEGMENT, arbevetel=5000, keret=20_000_000, roas=10, keses=0, hossz=36)
    r = score_client(best)
    assert (r.pontszam, r.szint, r.leggyengebb, r.ertekkapu) == (100.0, "A", "—", True)
    assert set(r.csoportok.values()) == {100.0}

    worst = _all(1, nev="Y", szegmens=SEGMENTS[-1][0], arbevetel=0, keret=0, roas=0, keses=90, hossz=0)
    r = score_client(worst)
    assert (r.pontszam, r.szint) == (0.0, "D")


def test_value_gate_caps_non_ideal_segments_at_second_tier():
    strong = _all(10, nev="Erős, de nem értékalapú", szegmens="Általános vállalkozás", arbevetel=5000,
                  keret=20_000_000, roas=10, keses=0, hossz=36)
    r = score_client(strong)
    assert r.pontszam >= 80 and not r.ertekkapu
    assert (r.szint, r.szorzo) == ("B", 1.05)

    r = score_client(strong | {"szegmens": IDEAL_SEGMENT, "felelosseg": 6})
    assert (r.ertekkapu, r.szint) == (False, "B")
    r = score_client(strong | {"szegmens": IDEAL_SEGMENT, "felelosseg": 7})
    assert (r.ertekkapu, r.szint) == (True, "A")


def test_fee_uses_largest_of_three_components():
    base = {"nev": "Z", "google": "igen", "meta": "nem", "orak": 10, "keret": 0}
    r = score_client(base)
    assert r.munkadij == 60_000 + 10 * 12_000
    assert r.dij_alapja == "Munkaigény × szintszorzó"
    assert r.javasolt_dij == 240_000  # 180 000 × 1,30 = 234 000 → felfelé kerekítve

    r = score_client({**base, "keret": 8_000_000})
    assert (r.javasolt_dij, r.dij_alapja) == (640_000, "Hirdetési keret arányában")

    r = score_client({"nev": "Kicsi", "google": "nem", "meta": "nem", "orak": 1})
    assert (r.javasolt_dij, r.dij_alapja) == (150_000, "Minimumdíj")


def test_ranking_orders_and_handles_ties():
    results = score_clients([{"nev": "B"}, {"nev": "A"}, {"nev": ""}])
    assert [r.nev for r in results] == ["B", "A"]
    assert [r.rang for r in results] == [1, 1]


def test_sample_covers_every_tier_and_the_gate():
    results = score_clients(SAMPLE_CLIENTS)
    assert {r.szint for r in results} == {"A", "B", "C", "D"}
    fast_fashion = next(r for r in results if "Gyorsdivat" in r.nev)
    assert not fast_fashion.ertekkapu and fast_fashion.szint != "A"


# ---- A táblázat képletei ugyanazt adják, mint a Python-motor ----

@pytest.fixture(scope="module")
def recalculated(tmp_path_factory):
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        pytest.skip("LibreOffice nincs telepítve")
    d = tmp_path_factory.mktemp("wb")
    src = d / "icp.xlsx"
    build_workbook(SAMPLE_CLIENTS).save(src)
    out = d / "out"
    subprocess.run([soffice, "--headless", f"-env:UserInstallation=file://{d}/lo", "--convert-to", "xlsx",
                    "--outdir", str(out), str(src)], check=True, capture_output=True, timeout=180)
    return load_workbook(out / "icp.xlsx", data_only=True)


def test_workbook_formulas_match_engine(recalculated):
    ws = recalculated[S_RES]
    rc = ResCols()
    f = rc.fixed
    expected = {r.nev: r for r in score_clients(SAMPLE_CLIENTS)}
    for i in range(len(SAMPLE_CLIENTS)):
        row = FIRST + i
        name = ws[f"A{row}"].value
        e = expected[name]
        for c in CRITERIA:
            assert ws[f"{rc.points[c['kod']]}{row}"].value == e.pontok[c["kod"]], (name, c["kod"])
        for g, _ in GROUPS:
            assert ws[f"{rc.groups[g]}{row}"].value == pytest.approx(e.csoportok[g]), (name, g)
        got = {k: ws[f"{f[k]}{row}"].value for k in f}
        assert got["score"] == pytest.approx(e.pontszam), name
        assert got["rank"] == e.rang
        assert got["gate"] == (GATE_OK if e.ertekkapu else GATE_FAIL), name
        assert (got["tier"], got["tier_name"]) == (e.szint, e.szint_nev), name
        assert got["mult"] == pytest.approx(e.szorzo), name
        assert (got["best"], got["worst"]) == (e.legerosebb, e.leggyengebb), name
        assert got["fee"] == e.javasolt_dij, name
        assert got["basis"] == e.dij_alapja, name
        assert got["missing"] == e.hianyzo_adatok
        assert (got["diff"] if got["diff"] != "" else None) == e.elteres
    assert ws[f"A{FIRST + len(SAMPLE_CLIENTS)}"].value in (None, "")


def test_ranking_sheet_is_sorted(recalculated):
    ws = recalculated[S_RANK]
    names = [ws[f"B{r}"].value for r in range(5, 5 + len(SAMPLE_CLIENTS))]
    assert names == [r.nev for r in score_clients(SAMPLE_CLIENTS)]
    assert ws[f"B{5 + len(SAMPLE_CLIENTS)}"].value in (None, "")
