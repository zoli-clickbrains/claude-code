import shutil
import subprocess

import pytest
from openpyxl import load_workbook

from icp.config import CRITERIA, GROUPS, SCALE_MAX, default_config
from icp.engine import ACCEPT, REJECT, WAITLIST, criterion_points, portfolio_summary, score_client, score_clients
from icp.sample import SAMPLE_CLIENTS
from icp.workbook import FIRST, NO, S_PORT, S_RANK, S_RES, YES, PortCells, ResCols, build_workbook

CFG = default_config()
BY_CODE = {c["kod"]: c for c in CRITERIA}
OPTIONS = {c["kod"]: [name for name, _ in c["opciok"]] for c in CRITERIA if c["tipus"] == "lista"}
IDEAL_SEGMENT = OPTIONS["szegmens"][0]
OPEN_TO_COMMISSION, NO_COMMISSION = OPTIONS["jutalek_nyitottsag"][1], OPTIONS["jutalek_nyitottsag"][3]
SYSTEM = {g for g in CFG["system_groups"]}


def _value_for(c, points):
    """Olyan bemeneti érték, amelyre a kritérium pontja a lehető legközelebb van `points`-hoz."""
    if c["tipus"] == "skala":
        return points
    pairs = c["savok"] if c["tipus"] == "sav" else c["opciok"]
    return min(pairs, key=lambda pair: abs(pair[1] - points))[0]


def _split(system_points, partner_points, name="X", **extra):
    """Minden szempont kitöltve: a rendszer- és a partnercsoportokban eltérő ponttal, ideális szegmensben."""
    client = {c["kod"]: _value_for(c, system_points if c["csoport"] in SYSTEM else partner_points)
              for c in CRITERIA}
    return client | {"nev": name, "szegmens": IDEAL_SEGMENT} | extra


def _client(points, name="X", **extra):
    return _split(points, points, name, **extra)


def test_weights_groups_and_merged_criteria():
    assert sum(c["suly"] for c in CRITERIA) == 100
    assert {c["csoport"] for c in CRITERIA} == {g for g, _ in GROUPS}
    assert sum(c["csoport"] == "Marketingezhetőség" for c in CRITERIA) == 15
    assert len({c["kod"] for c in CRITERIA}) == len(CRITERIA)
    assert BY_CODE["keret"]["suly"] < max(c["suly"] for c in CRITERIA)
    assert {"tapasztalat", "jutalek_nyitottsag", "elvarasok", "emberi", "szakmai"} <= set(BY_CODE)


@pytest.mark.parametrize("value,expected", [(0, 1), (149_999, 1), (150_000, 3), (599_999, 5), (2_000_000, 10)])
def test_budget_bands(value, expected):
    assert criterion_points(BY_CODE["keret"], value, CFG) == expected


@pytest.mark.parametrize("days,expected", [(0, 10), (1, 8), (7, 8), (8, 6), (30, 4), (31, 2), (90, 1)])
def test_payment_delay_bands_are_descending(days, expected):
    assert criterion_points(BY_CODE["keses"], days, CFG) == expected


def test_scale_and_list_points():
    assert criterion_points(BY_CODE["emberi"], 14, CFG) == SCALE_MAX
    assert criterion_points(BY_CODE["emberi"], 0, CFG) == 1
    assert criterion_points(BY_CODE["emberi"], 6.5, CFG) == 7
    assert criterion_points(BY_CODE["szegmens"], f" {IDEAL_SEGMENT.upper()} ", CFG) == 10
    assert criterion_points(BY_CODE["tapasztalat"], OPTIONS["tapasztalat"][0], CFG) == 10
    assert criterion_points(BY_CODE["roas"], "", CFG) == CFG["missing_points"]
    assert criterion_points(BY_CODE["szegmens"], "Ismeretlen", CFG) == CFG["missing_points"]


def test_perfect_client_is_stable_system():
    best = _client(10, arbevetel=5000, keret=20_000_000, roas=10, keses=0, hossz=36,
                   tapasztalat=OPTIONS["tapasztalat"][0], jutalek_nyitottsag=OPTIONS["jutalek_nyitottsag"][0])
    r = score_client(best)
    assert (r.pontszam, r.rendszer, r.statusz, r.leggyengebb) == (100.0, 100.0, 1, "—")


@pytest.mark.parametrize("system,expected", [(10, 1), (7, 2), (5, 3)])
def test_status_follows_system_score(system, expected):
    assert score_client(_split(system, 9)).statusz == expected


def test_status4_requires_strong_values_and_relationship():
    weak_system = _split(4, 10)
    r = score_client(weak_system)
    assert r.rendszer < 40 and r.statusz == 4
    assert score_client(_split(4, 6)).statusz == 5  # gyenge kapcsolat → kritikus
    assert score_client(_split(3, 10)).statusz == 5  # a rendszer 25 pont alatt → kritikus


def test_knockout_rejects_only_on_entered_data():
    assert score_client(_client(10, szegmens=OPTIONS["szegmens"][-1])).statusz == 5
    r = score_client(_client(10, felelosseg=3))
    assert (r.kizart, r.statusz, r.javasolt_dij) == (True, 5, None)
    assert not score_client(_client(10, felelosseg="")).kizart  # hiányzó adat nem zár ki


def test_only_one_status4_is_accepted():
    results = score_clients([_split(4, 10, "Első"), _split(4, 10, "Második"), _split(10, 10, "Stabil")])
    assert [(r.nev, r.statusz, r.dontes) for r in results] == [
        ("Stabil", 1, ACCEPT), ("Első", 4, ACCEPT), ("Második", 4, WAITLIST)]


def test_fee_uses_largest_of_three_components():
    base = _client(8, google="igen", meta="nem", orak=10, keret=0)
    r = score_client(base)
    assert r.statusz == 1
    assert r.munkadij == 60_000 + 10 * 12_000
    assert (r.javasolt_dij, r.dij_alapja) == (170_000, "Munkaigény × státuszszorzó")  # 180 000 × 0,90

    r = score_client({**base, "keret": 8_000_000})
    assert (r.javasolt_dij, r.dij_alapja) == (640_000, "Hirdetési keret arányában")

    r = score_client({**base, "orak": 1})
    assert (r.javasolt_dij, r.dij_alapja) == (150_000, "Minimumdíj")


def test_commission_model_needs_open_client_and_commission_status():
    base = _client(8, google="igen", meta="igen", orak=20, bevetel=2_000_000)
    r = score_client(base | {"jutalek_nyitottsag": OPEN_TO_COMMISSION})
    assert r.statusz == 1 and r.jutalekos_modell
    assert r.javasolt_dij == 320_000  # 350 000 × 0,90 = 315 000 → felfelé kerekítve
    assert r.jutalekos_fix_dij == 240_000 and r.jutalek_szazalek == 0.05 and r.varhato_jutalek == 100_000
    assert not score_client(base | {"jutalek_nyitottsag": NO_COMMISSION}).jutalekos_modell
    status3 = score_client(_split(5, 9, jutalek_nyitottsag=OPEN_TO_COMMISSION))
    assert status3.statusz == 3 and not status3.jutalekos_modell


def test_ranking_orders_and_handles_ties():
    results = score_clients([{"nev": "B"}, {"nev": "A"}, {"nev": ""}])
    assert [r.nev for r in results] == ["B", "A"]
    assert [r.rang for r in results] == [1, 1]
    assert all(r.dontes == REJECT for r in results)


def test_sample_portfolio_meets_targets():
    results = score_clients(SAMPLE_CLIENTS)
    summary = portfolio_summary(results)
    assert {r.statusz for r in results} == {1, 2, 3, 4, 5}
    assert summary["varolista"] == 1
    assert summary["top2_rendben"] and summary["s3_rendben"] and summary["s4_rendben"]


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


def _opt(value):
    return None if value == "" else value


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
        assert got["system"] == pytest.approx(e.rendszer), name
        assert got["rank"] == e.rang
        assert got["knockout"] == (YES if e.kizart else NO), name
        assert (got["status"], got["status_name"], got["fee_model"]) == (e.statusz, e.statusz_nev, e.dijmodell)
        assert got["decision"] == e.dontes, name
        assert _opt(got["mult"]) == e.szorzo, name
        assert (got["best"], got["worst"]) == (e.legerosebb, e.leggyengebb), name
        assert _opt(got["fee"]) == e.javasolt_dij, name
        assert got["basis"] == e.dij_alapja, name
        assert _opt(got["diff"]) == e.elteres, name
        assert _opt(got["comm_model"]) == (None if e.statusz == 5 else YES if e.jutalekos_modell else NO), name
        assert _opt(got["comm_fee"]) == e.jutalekos_fix_dij, name
        assert _opt(got["comm_pct"]) == e.jutalek_szazalek, name
        assert _opt(got["comm_expected"]) == e.varhato_jutalek, name
        assert got["missing"] == e.hianyzo_adatok
    assert ws[f"A{FIRST + len(SAMPLE_CLIENTS)}"].value in (None, "")


def test_ranking_sheet_is_sorted(recalculated):
    ws = recalculated[S_RANK]
    names = [ws[f"B{r}"].value for r in range(5, 5 + len(SAMPLE_CLIENTS))]
    assert names == [r.nev for r in score_clients(SAMPLE_CLIENTS)]
    assert ws[f"B{5 + len(SAMPLE_CLIENTS)}"].value in (None, "")


def test_portfolio_sheet_matches_engine(recalculated):
    ws = recalculated[S_PORT]
    summary = portfolio_summary(score_clients(SAMPLE_CLIENTS))
    pc = PortCells
    assert [ws[f"C{pc.count_row0 + i}"].value for i in range(5)] == [summary["darab"][k] for k in range(1, 6)]
    assert ws[pc.accepted].value == summary["elfogadott"]
    assert ws[pc.waitlist].value == summary["varolista"]
    assert ws[f"D{pc.top2}"].value == pytest.approx(summary["top2_arany"])
    assert ws[f"D{pc.s3}"].value == pytest.approx(summary["s3_arany"])
    assert [ws[f"E{r}"].value for r in (pc.top2, pc.s3, pc.s4)] == ["Rendben"] * 3
