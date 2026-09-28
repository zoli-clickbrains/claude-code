"""ICP-pontozó, státuszbesoroló és díjjavaslat-motor.

Ugyanazt a logikát valósítja meg, mint a táblázat képletei (a tesztek ezt
LibreOffice-szal ellenőrzik), így később a belső rendszerbe közvetlenül
beépíthető.
"""

import math
from dataclasses import dataclass

from .config import GROUPS, default_config

ACCEPT, WAITLIST, REJECT = "Elfogadjuk", "Várólista (a 4-es hely foglalt)", "Elutasítjuk"


@dataclass
class Result:
    nev: str
    pontok: dict  # kritériumkód -> 1–10 pont
    csoportok: dict  # csoportnév -> 0–100
    pontszam: float  # 0–100, a rangsor alapja
    rendszer: float  # 0–100, a státusz alapja
    kizart: bool  # kizáró feltétel teljesült (pl. értékeinkkel nem összeegyeztethető)
    statusz: int  # 1–5
    statusz_nev: str
    dijmodell: str
    legerosebb: str
    leggyengebb: str
    munkadij: float
    szorzo: float | None
    keret_minimum: float
    javasolt_dij: int | None  # 5. státusznál nincs ajánlat
    dij_alapja: str
    jelenlegi_dij: float | None
    elteres: float | None
    elteres_szazalek: float | None
    jutalekos_modell: bool
    jutalekos_fix_dij: int | None
    jutalek_szazalek: float | None
    varhato_jutalek: float | None
    hianyzo_adatok: int
    dontes: str = ""
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
        lookup = {name.casefold(): pts for name, pts in criterion["opciok"]}
        return lookup.get(str(value).strip().casefold(), missing)
    raise ValueError(f"Ismeretlen kritériumtípus: {kind}")


def _weighted(criteria, points, top):
    """Súlyozott pontszám 0–100 között.

    Ugyanabban a műveleti sorrendben, mint a táblázat képlete, hogy a
    kerekítési határesetek (pl. x,x5) is egyezzenek.
    """
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


def classify(system, groups, knocked_out, config):
    """Státusz (1–5) a rendszerpontszám, a csoportpontszámok és a kizáró feltétel alapján."""
    if knocked_out:
        return 5
    for s in config["statuses"][:3]:
        if system >= s["min"]:
            return s["statusz"]
    strong = all(groups[g] >= v for g, v in config["status4_rule"].items())
    return 4 if system >= config["status4_min_system"] and strong else 5


def score_client(client, config=None):
    """Egy ügyfél pontozása, státusza és díjjavaslata. `client`: oszlopkulcs -> érték.

    A döntést (elfogadás, várólista) a portfólió határozza meg, azt a
    `score_clients` tölti ki.
    """
    config = config or default_config()
    criteria = config["criteria"]
    pricing = config["pricing"]
    top = config["scale_max"]

    points = {c["kod"]: criterion_points(c, client.get(c["kod"]), config) for c in criteria}
    score = _round1(_weighted(criteria, points, top))
    groups = {g: _round1(_weighted([c for c in criteria if c["csoport"] == g], points, top)) for g, _ in GROUPS}
    system = _round1(_weighted([c for c in criteria if c["csoport"] in config["system_groups"]], points, top))
    # Kizárni csak kitöltött adat alapján lehet, a hiányzó mező pontja nem számít.
    knocked_out = any(not _blank(client.get(k)) and points[k] <= v for k, v in config["knockout"].items())
    status = config["statuses"][classify(system, groups, knocked_out, config) - 1]

    strongest = max(criteria, key=lambda c: points[c["kod"]] + c["suly"] / 1000)
    lost = [c["suly"] * (top - points[c["kod"]]) for c in criteria]
    weakest = "—" if max(lost) == 0 else criteria[lost.index(max(lost))]["nev"]

    labour = (
        pricing["google_alap"] * _yes(client.get("google"))
        + pricing["meta_alap"] * _yes(client.get("meta"))
        + pricing["egyeb_alap"] * _num(client.get("egyeb_csatorna"))
        + pricing["oradij"] * _num(client.get("orak"))
    )
    by_budget = _num(client.get("keret")) * pricing["keret_arany"]
    current = None if _blank(client.get("jelenlegi_dij")) else float(client["jelenlegi_dij"])

    fee = diff = diff_pct = None
    commission_model = False
    commission_fee = commission_pct = expected_commission = None
    if status["szorzo"] is None:
        basis = "Elutasítva"
    else:
        by_labour = labour * status["szorzo"]
        raw = max(pricing["min_dij"], by_labour, by_budget)
        if raw == by_labour:
            basis = "Munkaigény × státuszszorzó"
        elif raw == by_budget:
            basis = "Hirdetési keret arányában"
        else:
            basis = "Minimumdíj"
        fee = _ceiling(raw, pricing["kerekites"])
        diff = None if current is None else fee - current
        diff_pct = None if not current else diff / current

        commission_model = (status["jutalek"] > 0
                            and points["jutalek_nyitottsag"] >= config["commission"]["kuszob"])
        if commission_model:
            commission_fee = _ceiling(fee * (1 - config["commission"]["fix_csokkentes"]), pricing["kerekites"])
            commission_pct = status["jutalek"]
            revenue = client.get("bevetel")
            expected_commission = None if _blank(revenue) else float(revenue) * commission_pct

    return Result(
        nev=client.get("nev", ""),
        pontok=points,
        csoportok=groups,
        pontszam=score,
        rendszer=system,
        kizart=knocked_out,
        statusz=status["statusz"],
        statusz_nev=status["nev"],
        dijmodell=status["dijmodell"],
        legerosebb=strongest["nev"],
        leggyengebb=weakest,
        munkadij=labour,
        szorzo=status["szorzo"],
        keret_minimum=by_budget,
        javasolt_dij=fee,
        dij_alapja=basis,
        jelenlegi_dij=current,
        elteres=diff,
        elteres_szazalek=diff_pct,
        jutalekos_modell=commission_model,
        jutalekos_fix_dij=commission_fee,
        jutalek_szazalek=commission_pct,
        varhato_jutalek=expected_commission,
        hianyzo_adatok=sum(_blank(client.get(c["kod"])) for c in criteria),
    )


def score_clients(clients, config=None):
    """Több ügyfél pontozása, rangsorral és döntéssel.

    Holtversenyben azonos a rang (mint az Excel RANK), a sorrendben pedig a
    korábban felvett ügyfél áll elöl – ugyanúgy, mint a Rangsor lapon. A 4.
    státuszból csak a legjobb `max_status4` ügyfelet fogadjuk el, a többi
    várólistára kerül.
    """
    config = config or default_config()
    results = [score_client(c, config) for c in clients if not _blank(c.get("nev"))]
    for r in results:
        r.rang = 1 + sum(o.pontszam > r.pontszam for o in results)
    results.sort(key=lambda r: r.rang)  # stabil rendezés: a felvétel sorrendje marad

    status4_seen = 0
    for r in results:
        if r.statusz == 5:
            r.dontes = REJECT
        elif r.statusz == 4:
            r.dontes = ACCEPT if status4_seen < config["max_status4"] else WAITLIST
            status4_seen += 1
        else:
            r.dontes = ACCEPT
    return results


def portfolio_summary(results, config=None):
    """Az elfogadott ügyfelek státuszeloszlása és a célokhoz mért ellenőrzés."""
    config = config or default_config()
    targets = config["portfolio_targets"]
    accepted = [r for r in results if r.dontes == ACCEPT]
    # 1–4: elfogadott ügyfelek; 5: elutasítottak.
    counts = {k: sum(r.statusz == k for r in accepted) for k in (1, 2, 3, 4)}
    counts[5] = sum(r.statusz == 5 for r in results)
    n = len(accepted)
    top2 = (counts[1] + counts[2]) / n if n else 0.0
    s3 = counts[3] / n if n else 0.0
    return {
        "elfogadott": n,
        "darab": counts,
        "varolista": sum(r.dontes == WAITLIST for r in results),
        "top2_arany": top2,
        "s3_arany": s3,
        "top2_rendben": targets["top2_min"] <= top2 <= targets["top2_max"],
        "s3_rendben": targets["s3_min"] <= s3 <= targets["s3_max"],
        "s4_rendben": counts[4] <= config["max_status4"],
    }
