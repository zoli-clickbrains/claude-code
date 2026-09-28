"""ICP-pontozó és díjjavaslat-motor.

Ugyanazt a logikát valósítja meg, mint a táblázat képletei (a tesztek ezt
LibreOffice-szal ellenőrzik), így később a belső rendszerbe közvetlenül
beépíthető.
"""

import math
from dataclasses import dataclass, field

from .config import default_config


@dataclass
class Result:
    nev: str
    pontok: dict  # kritériumkód -> 1–5 pont
    pontszam: float  # 0–100
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
    extra: dict = field(default_factory=dict)


def _blank(value):
    return value is None or (isinstance(value, str) and value.strip() == "")


def _band_points(value, bands):
    value = max(float(value), bands[0][0])
    points = bands[0][1]
    for lower, pts in bands:
        if value >= lower:
            points = pts
    return points


def criterion_points(criterion, value, config):
    """Egy kritérium pontja (1–5); hiányzó adatnál a beállított pont."""
    missing = config["missing_points"]
    if _blank(value):
        return missing
    kind = criterion["tipus"]
    if kind == "sav":
        return _band_points(value, criterion["savok"])
    if kind == "skala":
        # Az Excel ROUND-hoz igazodva: .5 felfelé kerekít.
        return min(5, max(1, math.floor(float(value) + 0.5)))
    if kind == "lista":
        lookup = {name.casefold(): pts for name, pts in config["industries"]}
        return lookup.get(str(value).strip().casefold(), missing)
    raise ValueError(f"Ismeretlen kritériumtípus: {kind}")


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

    points = {c["kod"]: criterion_points(c, client.get(c["kod"]), config) for c in criteria}
    raw_score = sum(c["suly"] * (points[c["kod"]] - 1) / 4 for c in criteria)
    # Excel-féle kerekítés (.5 felfelé); a Python round() banki kerekítést használna.
    score = math.floor(raw_score * 10 + 0.5 + 1e-9) / 10

    tier = config["tiers"][0]
    for t in config["tiers"]:
        if score >= t["min"]:
            tier = t

    strongest = max(criteria, key=lambda c: points[c["kod"]] + c["suly"] / 1000)
    lost = [c["suly"] * (5 - points[c["kod"]]) for c in criteria]
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
        pontszam=score,
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
