"""ICP-pontozó és díjjavaslat-motor.

Ugyanazt a logikát valósítja meg, mint a táblázat képletei (a tesztek ezt
LibreOffice-szal ellenőrzik), így később a belső rendszerbe közvetlenül
beépíthető.
"""

import math
from dataclasses import dataclass

from .config import GROUPS, default_config


@dataclass
class Result:
    nev: str
    pontok: dict  # kritériumkód -> 1–10 pont
    csoportok: dict  # csoportnév -> 0–100
    pontszam: float  # 0–100
    ertekkapu: bool
    szint: str
    szint_nev: str
    legerosebb: str
    leggyengebb: str
    munkadij: float
    szorzo: float
    keret_minimum: float
    javasolt_dij: int
    dij_alapja: str
    jelenlegi_dij: float | None
    elteres: float | None
    elteres_szazalek: float | None
    jutalek: float | None
    hianyzo_adatok: int
    rang: int = 0


def _blank(value):
    return value is None or (isinstance(value, str) and value.strip() == "")


def _round1(value):
    # Excel-féle kerekítés (.5 felfelé); a Python round() banki kerekítést használna.
    return math.floor(value * 10 + 0.5 + 1e-9) / 10


def _band_points(value, bands):
    value = max(float(value), bands[0][0])
    points = bands[0][1]
    for lower, pts in bands:
        if value >= lower:
            points = pts
    return points


def criterion_points(criterion, value, config):
    """Egy kritérium pontja (1–10); hiányzó adatnál a beállított pont."""
    missing = config["missing_points"]
    if _blank(value):
        return missing
    kind = criterion["tipus"]
    if kind == "sav":
        return _band_points(value, criterion["savok"])
    if kind == "skala":
        return min(config["scale_max"], max(1, math.floor(float(value) + 0.5)))
    if kind == "lista":
        lookup = {name.casefold(): pts for name, pts in config["segments"]}
        return lookup.get(str(value).strip().casefold(), missing)
    raise ValueError(f"Ismeretlen kritériumtípus: {kind}")


def _weighted(criteria, points, top):
    """Súlyozott pontszám 0–100 között."""
    # Ugyanabban a műveleti sorrendben, mint a táblázat képlete, hogy a
    # kerekítési határesetek (pl. x,x5) is egyezzenek.
    total = sum(c["suly"] for c in criteria)
    return 100 * sum(c["suly"] * (points[c["kod"]] - 1) for c in criteria) / ((top - 1) * total)


def _yes(value):
    return not _blank(value) and str(value).strip().casefold() == "igen"


def _num(value):
    return 0.0 if _blank(value) else float(value)


def _ceiling(value, step):
    # A táblázat is 2 tizedesre kerekít a felfelé kerekítés előtt, hogy a
    # lebegőpontos zaj (pl. 340000 × 1,15) ne ugorjon a következő lépcsőre.
    return int(math.ceil(round(value, 2) / step) * step)


def score_client(client, config=None):
    """Egy ügyfél pontozása és díjjavaslata. `client`: oszlopkulcs -> érték."""
    config = config or default_config()
    criteria = config["criteria"]
    pricing = config["pricing"]
    top = config["scale_max"]

    points = {c["kod"]: criterion_points(c, client.get(c["kod"]), config) for c in criteria}
    score = _round1(_weighted(criteria, points, top))
    groups = {g: _round1(_weighted([c for c in criteria if c["csoport"] == g], points, top)) for g, _ in GROUPS}

    gate = all(points[k] >= v for k, v in config["value_gate"].items())
    tiers = config["tiers"]
    tier = tiers[0]
    for t in tiers:
        if score >= t["min"]:
            tier = t
    if tier is tiers[-1] and not gate:
        tier = tiers[-2]

    strongest = max(criteria, key=lambda c: points[c["kod"]] + c["suly"] / 1000)
    lost = [c["suly"] * (top - points[c["kod"]]) for c in criteria]
    weakest = "—" if max(lost) == 0 else criteria[lost.index(max(lost))]["nev"]

    labour = (
        pricing["google_alap"] * _yes(client.get("google"))
        + pricing["meta_alap"] * _yes(client.get("meta"))
        + pricing["egyeb_alap"] * _num(client.get("egyeb_csatorna"))
        + pricing["oradij"] * _num(client.get("orak"))
    )
    by_labour = labour * tier["szorzo"]
    by_budget = _num(client.get("keret")) * pricing["keret_arany"]
    raw = max(pricing["min_dij"], by_labour, by_budget)
    if raw == by_labour:
        basis = "Munkaigény × szintszorzó"
    elif raw == by_budget:
        basis = "Hirdetési keret arányában"
    else:
        basis = "Minimumdíj"
    fee = _ceiling(raw, pricing["kerekites"])

    current = None if _blank(client.get("jelenlegi_dij")) else float(client["jelenlegi_dij"])
    diff = None if current is None else fee - current
    diff_pct = None if not current else diff / current
    revenue = client.get("bevetel")
    commission = None if _blank(revenue) else float(revenue) * tier["jutalek"]

    return Result(
        nev=client.get("nev", ""),
        pontok=points,
        csoportok=groups,
        pontszam=score,
        ertekkapu=gate,
        szint=tier["szint"],
        szint_nev=tier["nev"],
        legerosebb=strongest["nev"],
        leggyengebb=weakest,
        munkadij=labour,
        szorzo=tier["szorzo"],
        keret_minimum=by_budget,
        javasolt_dij=fee,
        dij_alapja=basis,
        jelenlegi_dij=current,
        elteres=diff,
        elteres_szazalek=diff_pct,
        jutalek=commission,
        hianyzo_adatok=sum(_blank(client.get(c["kod"])) for c in criteria),
    )


def score_clients(clients, config=None):
    """Több ügyfél pontozása, rangsorral.

    Holtversenyben azonos a rang (mint az Excel RANK), a sorrendben pedig a
    korábban felvett ügyfél áll elöl – ugyanúgy, mint a Rangsor lapon.
    """
    config = config or default_config()
    results = [score_client(c, config) for c in clients if not _blank(c.get("nev"))]
    for r in results:
        r.rang = 1 + sum(o.pontszam > r.pontszam for o in results)
    return sorted(results, key=lambda r: r.rang)  # stabil rendezés: a felvétel sorrendje marad
