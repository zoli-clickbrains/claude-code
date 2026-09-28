"""Az ICP-ügyfélminősítő táblázat előállítása élő képletekkel.

A képletek Excelben, Google Táblázatokban és LibreOffice-ban is működnek
(VLOOKUP, INDEX/MATCH, SUMIF, RANK, LARGE, CEILING – nincs dinamikus tömbképlet).
"""

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule, DataBarRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as col
from openpyxl.worksheet.datavalidation import DataValidation

from .config import (CRITERIA, GROUPS, INPUT_COLUMNS, MISSING_POINTS, PRICING, SCALE_MAX, SEGMENTS, TIERS,
                     VALUE_GATE)

MAX_ROWS = 200  # ennyi ügyfélsorra készülnek elő a képletek
FIRST, LAST = 2, MAX_ROWS + 1

S_IN, S_SET, S_RES, S_RANK, S_HELP = "Ügyfelek", "Beállítások", "Eredmény", "Rangsor", "Útmutató"

FT = '#,##0" Ft"'
PCT = "0%"
GATE_OK, GATE_FAIL = "Teljesül", "Nem teljesül"

NAVY = "032D42"
GROUP_COLORS = dict(GROUPS) | {"Díjazás": "5E5E5A"}
HEAD_FONT = Font(bold=True, color="FFFFFF")
INPUT_FILL = PatternFill("solid", fgColor="FFF8E1")
TITLE_FONT = Font(bold=True, size=14, color=NAVY)
SECTION_FONT = Font(bold=True, size=12, color=NAVY)
THIN = Border(bottom=Side(style="thin", color="C3C2B7"))
TIER_FILLS = {
    "A": PatternFill("solid", fgColor="C8EFC0"),
    "B": PatternFill("solid", fgColor="E3F4D9"),
    "C": PatternFill("solid", fgColor="FDEBC2"),
    "D": PatternFill("solid", fgColor="F6CFCF"),
}


def fill(color):
    return PatternFill("solid", fgColor=color)


def ref(sheet, cell):
    return f"'{sheet}'!{cell}"


def abs_ref(sheet, column, row):
    return ref(sheet, f"${column}${row}")


def _head(cell, color=NAVY):
    cell.fill, cell.font = fill(color), HEAD_FONT
    cell.alignment = Alignment(wrap_text=True, vertical="center")


def _header(ws, row, labels, widths=None, colors=None):
    for i, label in enumerate(labels, start=1):
        _head(ws.cell(row=row, column=i, value=label), (colors or {}).get(i, NAVY))
        if widths:
            ws.column_dimensions[col(i)].width = widths[i - 1]
    ws.row_dimensions[row].height = 60


# ---------------------------------------------------------------- Beállítások

class Layout:
    """A Beállítások lap cellacímei, amikre a képletek hivatkoznak."""

    crit_first = 6
    band_col0 = 8  # H oszloptól jobbra, kritériumonként 3 oszlop
    band_row0 = 7

    def __init__(self):
        self.crit_last = self.crit_first + len(CRITERIA) - 1
        self.total_row = self.crit_last + 1
        self.group_first = self.total_row + 4
        self.seg_first = self.group_first + len(GROUPS) + 4
        self.seg_last = self.seg_first + len(SEGMENTS) + 3  # üres sorok új szegmenseknek
        self.param_row0 = self.seg_last + 4
        self.tier_first = self.param_row0 + 3 + len(PRICING) + 4
        self.tier_last = self.tier_first + len(TIERS) - 1

        self.weight = {c["kod"]: abs_ref(S_SET, "E", self.crit_first + i) for i, c in enumerate(CRITERIA)}
        self.names = f"{ref(S_SET, '$B$%d' % self.crit_first)}:$B${self.crit_last}"
        self.total_weight = abs_ref(S_SET, "E", self.total_row)
        self.group_weight = {g: abs_ref(S_SET, "B", self.group_first + i) for i, (g, _) in enumerate(GROUPS)}
        self.bands = {}
        for i, c in enumerate([c for c in CRITERIA if c["tipus"] == "sav"]):
            lo, pt = col(self.band_col0 + 3 * i), col(self.band_col0 + 3 * i + 1)
            last = self.band_row0 + len(c["savok"]) - 1
            self.bands[c["kod"]] = (
                f"{ref(S_SET, f'${lo}${self.band_row0}')}:${pt}${last}",
                abs_ref(S_SET, lo, self.band_row0),
            )
        self.segments = f"{ref(S_SET, f'$A${self.seg_first}')}:$B${self.seg_last}"
        self.segment_names = f"{ref(S_SET, f'$A${self.seg_first}')}:$A${self.seg_last}"
        self.missing = abs_ref(S_SET, "B", self.param_row0)
        self.gate = {k: abs_ref(S_SET, "B", self.param_row0 + 1 + i) for i, k in enumerate(VALUE_GATE)}
        p0 = self.param_row0 + 1 + len(VALUE_GATE)
        self.pricing = {k: abs_ref(S_SET, "B", p0 + i) for i, k in enumerate(PRICING)}
        self.tiers = f"{ref(S_SET, f'$A${self.tier_first}')}:$B${self.tier_last}"

        def tier_col(c):
            return f"{ref(S_SET, f'${c}${self.tier_first}')}:${c}${self.tier_last}"

        self.tier_codes, self.tier_names = tier_col("B"), tier_col("C")
        self.tier_mults, self.tier_commissions = tier_col("D"), tier_col("E")
        self.top_tier = abs_ref(S_SET, "B", self.tier_last)
        self.second_tier = abs_ref(S_SET, "B", self.tier_last - 1)


def _section(ws, row, text, column=1):
    ws.cell(row=row, column=column, value=text).font = SECTION_FONT


def build_settings(ws, lay):
    ws["A1"] = "Beállítások – kritériumok, súlyok, sávok, szintek, díjazás"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = (f"Minden kritérium 1–{SCALE_MAX} pontot kap, a {SCALE_MAX} mindig a számunkra kedvező érték. "
                "A sárga cellák szabadon módosíthatók; minden eredmény azonnal újraszámolódik.")

    _section(ws, 4, "Kritériumok és súlyok")
    for j, label in enumerate(["Kód", "Kritérium", "Csoport", "Típus", "Súly", "Leírás"], 1):
        _head(ws.cell(row=5, column=j, value=label))
    types = {"sav": "sáv", "skala": f"1–{SCALE_MAX} skála", "lista": "lista"}
    for i, c in enumerate(CRITERIA):
        r = lay.crit_first + i
        for j, v in enumerate([c["kod"], c["nev"], c["csoport"], types[c["tipus"]], c["suly"], c["leiras"]], 1):
            ws.cell(row=r, column=j, value=v)
        ws.cell(row=r, column=3).font = Font(color=GROUP_COLORS[c["csoport"]], bold=True)
        ws.cell(row=r, column=5).fill = INPUT_FILL
        ws.cell(row=r, column=6).alignment = Alignment(wrap_text=True, vertical="top")
    t = lay.total_row
    ws.cell(row=t, column=4, value="Összesen").font = Font(bold=True)
    ws.cell(row=t, column=5, value=f"=SUM(E{lay.crit_first}:E{lay.crit_last})").font = Font(bold=True)
    ws.cell(row=t, column=6, value=f'=IF(E{t}=100,"Rendben","Figyelem: a súlyok összege legyen 100!")')
    for w, c in zip([18, 30, 18, 11, 8, 70], "ABCDEF"):
        ws.column_dimensions[c].width = w

    _section(ws, lay.group_first - 2, "Csoportok súlya (automatikus)")
    for j, label in enumerate(["Csoport", "Súly összesen"], 1):
        _head(ws.cell(row=lay.group_first - 1, column=j, value=label))
    for i, (g, color) in enumerate(GROUPS):
        r = lay.group_first + i
        ws.cell(row=r, column=1, value=g).font = Font(color=color, bold=True)
        ws.cell(row=r, column=2, value=f"=SUMIF($C${lay.crit_first}:$C${lay.crit_last},A{r},"
                                       f"$E${lay.crit_first}:$E${lay.crit_last})")

    _section(ws, 4, "Sávok (alsó határ → pont)", lay.band_col0)
    for i, c in enumerate([c for c in CRITERIA if c["tipus"] == "sav"]):
        c0 = lay.band_col0 + 3 * i
        ws.cell(row=5, column=c0, value=c["nev"]).font = Font(bold=True)
        for j, label in enumerate(["Alsó határ", "Pont"]):
            _head(ws.cell(row=6, column=c0 + j, value=label))
        for k, (lower, pts) in enumerate(c["savok"]):
            for j, v in enumerate([lower, pts]):
                cell = ws.cell(row=lay.band_row0 + k, column=c0 + j, value=v)
                cell.fill = INPUT_FILL
                if j == 0:
                    cell.number_format = "#,##0.##"
        ws.column_dimensions[col(c0)].width = 13
        ws.column_dimensions[col(c0 + 1)].width = 7
        ws.column_dimensions[col(c0 + 2)].width = 3

    _section(ws, lay.seg_first - 2, "Értékalapú szegmensek pontértéke")
    for j, label in enumerate(["Szegmens", f"Pont (1–{SCALE_MAX})"], 1):
        _head(ws.cell(row=lay.seg_first - 1, column=j, value=label))
    for r in range(lay.seg_first, lay.seg_last + 1):
        for j in (1, 2):
            ws.cell(row=r, column=j).fill = INPUT_FILL
    for i, (name, pts) in enumerate(SEGMENTS):
        ws.cell(row=lay.seg_first + i, column=1, value=name)
        ws.cell(row=lay.seg_first + i, column=2, value=pts)
    ws.cell(row=lay.seg_first - 2, column=3, value="(új szegmens az üres sorokba vehető fel)")

    _section(ws, lay.param_row0 - 1, "Általános paraméterek, értékkapu és díjazás")
    names = {c["kod"]: c["nev"] for c in CRITERIA}
    params = [(f"Hiányzó adat pontja (1–{SCALE_MAX})", MISSING_POINTS, "0")]
    params += [(f"Értékkapu: „{names[k]}” legalább", v, "0") for k, v in VALUE_GATE.items()]
    params += [(label, value, PCT if k == "keret_arany" else FT) for k, (label, value) in PRICING.items()]
    for i, (label, value, fmt) in enumerate(params):
        r = lay.param_row0 + i
        ws.cell(row=r, column=1, value=label)
        cell = ws.cell(row=r, column=2, value=value)
        cell.fill, cell.number_format = INPUT_FILL, fmt

    _section(ws, lay.tier_first - 2, "Szintek és díjszorzók")
    for j, label in enumerate(["Minimum pontszám", "Szint", "Megnevezés", "Díjszorzó", "Jutalék % (később)"], 1):
        _head(ws.cell(row=lay.tier_first - 1, column=j, value=label))
    for i, t in enumerate(TIERS):
        r = lay.tier_first + i
        for j, v in enumerate([t["min"], t["szint"], t["nev"], t["szorzo"], t["jutalek"]], 1):
            ws.cell(row=r, column=j, value=v).fill = INPUT_FILL
        ws.cell(row=r, column=4).number_format = "0.00"
        ws.cell(row=r, column=5).number_format = "0.0%"
    ws.cell(row=lay.tier_last + 1, column=1, value=(
        "A szintek a minimum pontszám szerint növekvő sorrendben maradjanak. A legfelső szintet csak az "
        "értékkapun átjutó ügyfél kaphatja; aki nem jut át, eggyel lejjebb kerül.")).font = Font(italic=True)
    ws.freeze_panes = "A6"


# -------------------------------------------------------------------- Ügyfelek

def build_input(ws, lay, clients):
    colors = {j: GROUP_COLORS.get(group, NAVY) for j, (*_, group) in enumerate(INPUT_COLUMNS, 1)}
    _header(ws, 1, [h for _, h, *_ in INPUT_COLUMNS], [w for _, _, w, *_ in INPUT_COLUMNS], colors)
    ws.freeze_panes = "B2"

    for i, client in enumerate(clients):
        for j, (key, *_rest) in enumerate(INPUT_COLUMNS, 1):
            value = client.get(key, "")
            ws.cell(row=FIRST + i, column=j, value=None if value == "" else value)

    validations = {
        "lista": DataValidation(type="list", formula1=lay.segment_names, allow_blank=True),
        "igennem": DataValidation(type="list", formula1='"igen,nem"', allow_blank=True),
        "skala": DataValidation(type="whole", operator="between", formula1="1", formula2=str(SCALE_MAX),
                                allow_blank=True),
        "szam": DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0", allow_blank=True),
        "penz": DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0", allow_blank=True),
    }
    validations["skala"].error = f"1 és {SCALE_MAX} közötti egész számot adj meg."
    for dv in validations.values():
        dv.showErrorMessage = True
        ws.add_data_validation(dv)
    for j, (_, _, _, kind, _) in enumerate(INPUT_COLUMNS, 1):
        if kind in validations:
            validations[kind].add(f"{col(j)}{FIRST}:{col(j)}{LAST}")
        if kind == "penz":
            for r in range(FIRST, LAST + 1):
                ws.cell(row=r, column=j).number_format = FT


def input_cell(key, row):
    idx = [k for k, *_ in INPUT_COLUMNS].index(key) + 1
    return ref(S_IN, f"{col(idx)}{row}")


def criterion_formula(c, lay, row):
    x = input_cell(c["kod"], row)
    miss = lay.missing
    if c["tipus"] == "sav":
        rng, first = lay.bands[c["kod"]]
        body = f"VLOOKUP(MAX({x},{first}),{rng},2,TRUE)"
    elif c["tipus"] == "skala":
        body = f"MIN({SCALE_MAX},MAX(1,ROUND({x},0)))"
    else:
        body = f"IFERROR(VLOOKUP(TRIM({x}),{lay.segments},2,FALSE),{miss})"
    return f'=IF($A{row}="","",IF({x}="",{miss},{body}))'


# -------------------------------------------------------------------- Eredmény

RES_FIXED = [
    ("score", "ICP pontszám (0–100)", 12, "0.0"),
    ("rank", "Rang", 7, "0"),
    ("gate", "Értékkapu", 12, None),
    ("tier", "Szint", 7, None),
    ("tier_name", "Szint megnevezése", 18, None),
    ("best", "Legerősebb terület", 24, None),
    ("worst", "Leggyengébb terület", 24, None),
    ("labour", "Munkadíj (Ft/hó)", 14, FT),
    ("mult", "Szintszorzó", 10, "0.00"),
    ("budget_min", "Keretarányos minimum (Ft/hó)", 15, FT),
    ("fee", "Javasolt havi díj (Ft)", 15, FT),
    ("basis", "Díj alapja", 24, None),
    ("current", "Jelenlegi díj (Ft)", 14, FT),
    ("diff", "Eltérés (Ft)", 13, FT),
    ("diff_pct", "Eltérés (%)", 10, PCT),
    ("commission", "Opcionális jutalék (Ft/hó)", 14, FT),
    ("missing", "Hiányzó adatok (db)", 10, "0"),
]


class ResCols:
    def __init__(self):
        n = len(CRITERIA)
        self.points = {c["kod"]: col(2 + i) for i, c in enumerate(CRITERIA)}
        g0 = 2 + n
        self.groups = {g: col(g0 + i) for i, (g, _) in enumerate(GROUPS)}
        f0 = g0 + len(GROUPS)
        self.fixed = {k: col(f0 + i) for i, (k, *_) in enumerate(RES_FIXED)}
        h0 = f0 + len(RES_FIXED) + 1  # egy üres elválasztó oszlop után a segédoszlopok
        self.strength = [col(h0 + i) for i in range(n)]
        self.lost = [col(h0 + n + i) for i in range(n)]
        self.sort_key = col(h0 + 2 * n)
        self.helper_first, self.helper_last = h0, h0 + 2 * n


def _weighted_formula(criteria, lay, rc, r, total):
    terms = "+".join(f"{lay.weight[c['kod']]}*({rc.points[c['kod']]}{r}-1)" for c in criteria)
    return f"ROUND(100*({terms})/({SCALE_MAX - 1}*{total}),1)"


def build_results(ws, lay, rc):
    headers = (["Ügyfél"] + [c["nev"] for c in CRITERIA] + [f"{g} (0–100)" for g, _ in GROUPS]
               + [h for _, h, _, _ in RES_FIXED])
    widths = [28] + [11] * len(CRITERIA) + [17] * len(GROUPS) + [w for *_, w, _ in RES_FIXED]
    colors = {2 + i: GROUP_COLORS[c["csoport"]] for i, c in enumerate(CRITERIA)}
    colors |= {2 + len(CRITERIA) + i: color for i, (_, color) in enumerate(GROUPS)}
    _header(ws, 1, headers, widths, colors)
    helper_headers = ([f"erő: {c['kod']}" for c in CRITERIA] + [f"veszt: {c['kod']}" for c in CRITERIA]
                      + ["rendezőkulcs"])
    for i, h in enumerate(helper_headers):
        ws.cell(row=1, column=rc.helper_first + i, value=h).font = Font(italic=True, color="898781")

    f, p = rc.fixed, lay.pricing
    for r in range(FIRST, LAST + 1):
        blank = f'$A{r}=""'
        ws[f"A{r}"] = f'=IF({input_cell("nev", r)}="","",{input_cell("nev", r)})'
        for c in CRITERIA:
            ws[f"{rc.points[c['kod']]}{r}"] = criterion_formula(c, lay, r)
        for g, _ in GROUPS:
            group_criteria = [c for c in CRITERIA if c["csoport"] == g]
            cell = ws[f"{rc.groups[g]}{r}"]
            cell.value = f'=IF({blank},"",{_weighted_formula(group_criteria, lay, rc, r, lay.group_weight[g])})'
            cell.number_format = "0.0"

        ws[f"{f['score']}{r}"] = f'=IF({blank},"",{_weighted_formula(CRITERIA, lay, rc, r, lay.total_weight)})'
        ws[f"{f['rank']}{r}"] = f'=IF({blank},"",RANK({f["score"]}{r},${f["score"]}${FIRST}:${f["score"]}${LAST}))'
        gate = ",".join(f"{rc.points[k]}{r}>={cell}" for k, cell in lay.gate.items())
        ws[f"{f['gate']}{r}"] = f'=IF({blank},"",IF(AND({gate}),"{GATE_OK}","{GATE_FAIL}"))'
        base_tier = f"VLOOKUP({f['score']}{r},{lay.tiers},2,TRUE)"
        ws[f"{f['tier']}{r}"] = (f'=IF({blank},"",IF(AND({base_tier}={lay.top_tier},{f["gate"]}{r}="{GATE_FAIL}"),'
                                 f'{lay.second_tier},{base_tier}))')
        tier_pos = f"MATCH({f['tier']}{r},{lay.tier_codes},0)"
        ws[f"{f['tier_name']}{r}"] = f'=IF({blank},"",INDEX({lay.tier_names},{tier_pos}))'
        ws[f"{f['mult']}{r}"] = f'=IF({blank},"",INDEX({lay.tier_mults},{tier_pos}))'

        s_rng = f"{rc.strength[0]}{r}:{rc.strength[-1]}{r}"
        l_rng = f"{rc.lost[0]}{r}:{rc.lost[-1]}{r}"
        ws[f"{f['best']}{r}"] = f'=IF({blank},"",INDEX({lay.names},MATCH(MAX({s_rng}),{s_rng},0)))'
        ws[f"{f['worst']}{r}"] = (f'=IF({blank},"",IF(MAX({l_rng})=0,"—",'
                                  f'INDEX({lay.names},MATCH(MAX({l_rng}),{l_rng},0))))')
        for i, c in enumerate(CRITERIA):
            pt = f"{rc.points[c['kod']]}{r}"
            ws[f"{rc.strength[i]}{r}"] = f'=IF({blank},"",{pt}+{lay.weight[c["kod"]]}/1000)'
            ws[f"{rc.lost[i]}{r}"] = f'=IF({blank},"",{lay.weight[c["kod"]]}*({SCALE_MAX}-{pt}))'
        ws[f"{rc.sort_key}{r}"] = f'=IF({blank},"",{f["score"]}{r}-ROW()/100000)'

        def yes(key):
            return f'(TRIM({input_cell(key, r)})="igen")'

        ws[f"{f['labour']}{r}"] = (
            f'=IF({blank},"",{p["google_alap"]}*{yes("google")}+{p["meta_alap"]}*{yes("meta")}'
            f'+{p["egyeb_alap"]}*N({input_cell("egyeb_csatorna", r)})+{p["oradij"]}*N({input_cell("orak", r)}))'
        )
        ws[f"{f['budget_min']}{r}"] = f'=IF({blank},"",N({input_cell("keret", r)})*{p["keret_arany"]})'
        by_labour = f"{f['labour']}{r}*{f['mult']}{r}"
        raw = f"MAX({p['min_dij']},{by_labour},{f['budget_min']}{r})"
        ws[f"{f['fee']}{r}"] = f'=IF({blank},"",CEILING(ROUND({raw},2),{p["kerekites"]}))'
        ws[f"{f['basis']}{r}"] = (
            f'=IF({blank},"",IF({raw}={by_labour},"Munkaigény × szintszorzó",'
            f'IF({raw}={f["budget_min"]}{r},"Hirdetési keret arányában","Minimumdíj")))'
        )
        cur = input_cell("jelenlegi_dij", r)
        ws[f"{f['current']}{r}"] = f'=IF(OR({blank},{cur}=""),"",{cur})'
        ws[f"{f['diff']}{r}"] = f'=IF({f["current"]}{r}="","",{f["fee"]}{r}-{f["current"]}{r})'
        ws[f"{f['diff_pct']}{r}"] = (f'=IF(OR({f["current"]}{r}="",{f["current"]}{r}=0),"",'
                                     f'{f["diff"]}{r}/{f["current"]}{r})')
        rev = input_cell("bevetel", r)
        ws[f"{f['commission']}{r}"] = (f'=IF(OR({blank},{rev}=""),"",'
                                       f'{rev}*INDEX({lay.tier_commissions},{tier_pos}))')
        missing = "+".join(f'({input_cell(c["kod"], r)}="")' for c in CRITERIA)
        ws[f"{f['missing']}{r}"] = f'=IF({blank},"",{missing})'

        for key, _, _, fmt in RES_FIXED:
            if fmt:
                ws[f"{f[key]}{r}"].number_format = fmt

    for i in range(rc.helper_first, rc.helper_last + 1):
        ws.column_dimensions[col(i)].hidden = True
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = f"A1:{f['missing']}{LAST}"

    ws.conditional_formatting.add(f"{f['score']}{FIRST}:{f['score']}{LAST}", DataBarRule(
        start_type="num", start_value=0, end_type="num", end_value=100, color="2E981B"))
    _tier_colors(ws, f"{f['tier']}{FIRST}:{f['tier']}{LAST}")
    _gate_colors(ws, f"{f['gate']}{FIRST}:{f['gate']}{LAST}")
    diff_rng = f"{f['diff']}{FIRST}:{f['diff_pct']}{LAST}"
    ws.conditional_formatting.add(diff_rng, CellIsRule(operator="greaterThan", formula=["0"],
                                                       font=Font(color="1C7A3E", bold=True)))
    ws.conditional_formatting.add(diff_rng, CellIsRule(operator="lessThan", formula=["0"],
                                                       font=Font(color="C23434", bold=True)))


def _tier_colors(ws, rng):
    top_left = rng.split(":")[0]
    for tier, tier_fill in TIER_FILLS.items():
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{top_left}="{tier}"'], fill=tier_fill))


def _gate_colors(ws, rng):
    top_left = rng.split(":")[0]
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{top_left}="{GATE_OK}"'],
                                                   font=Font(color="1C7A3E", bold=True)))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{top_left}="{GATE_FAIL}"'],
                                                   font=Font(color="C23434")))


# --------------------------------------------------------------------- Rangsor

def build_ranking(ws, rc):
    cols = [("Rang", rc.fixed["rank"], 7, "0", NAVY), ("Ügyfél", "A", 30, None, NAVY),
            ("ICP pontszám", rc.fixed["score"], 12, "0.0", NAVY), ("Szint", rc.fixed["tier"], 7, None, NAVY),
            ("Szint megnevezése", rc.fixed["tier_name"], 20, None, NAVY),
            ("Értékkapu", rc.fixed["gate"], 12, None, NAVY)]
    cols += [(g, rc.groups[g], 17, "0.0", color) for g, color in GROUPS]
    cols += [("Javasolt havi díj", rc.fixed["fee"], 16, FT, NAVY),
             ("Jelenlegi díj", rc.fixed["current"], 15, FT, NAVY),
             ("Eltérés (Ft)", rc.fixed["diff"], 14, FT, NAVY),
             ("Eltérés (%)", rc.fixed["diff_pct"], 10, PCT, NAVY),
             ("Leggyengébb terület", rc.fixed["worst"], 26, None, NAVY),
             ("Hiányzó adatok", rc.fixed["missing"], 10, "0", NAVY)]
    ws["A1"] = "ICP-rangsor – pontszám szerint csökkenő sorrendben"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = ("Automatikusan frissül az Ügyfelek lap alapján. A csoportoszlopok 0–100 között mutatják, "
                "hol erős az ügyfél. Holtversenynél a korábban felvett ügyfél áll elöl.")
    _header(ws, 4, [c[0] for c in cols], [c[2] for c in cols], {i: c[4] for i, c in enumerate(cols, 1)})
    helper = col(len(cols) + 2)
    ws.column_dimensions[helper].hidden = True
    ws[f"{helper}4"] = "sor"
    key_rng = f"{ref(S_RES, f'${rc.sort_key}${FIRST}')}:${rc.sort_key}${LAST}"
    for k in range(1, MAX_ROWS + 1):
        r = 4 + k
        ws[f"{helper}{r}"] = f'=IFERROR(MATCH(LARGE({key_rng},{k}),{key_rng},0),"")'
        for j, (_, src, _, fmt, _) in enumerate(cols, 1):
            src_rng = f"{ref(S_RES, f'${src}${FIRST}')}:${src}${LAST}"
            cell = ws.cell(row=r, column=j, value=f'=IF(${helper}{r}="","",INDEX({src_rng},${helper}{r}))')
            if fmt:
                cell.number_format = fmt
            cell.border = THIN
    ws.freeze_panes = "C5"
    last = 4 + MAX_ROWS
    _tier_colors(ws, f"D5:D{last}")
    _gate_colors(ws, f"F5:F{last}")
    ws.conditional_formatting.add(f"C5:C{last}", DataBarRule(
        start_type="num", start_value=0, end_type="num", end_value=100, color="2E981B"))


# -------------------------------------------------------------------- Útmutató

def _guide():
    weights = {g: sum(c["suly"] for c in CRITERIA if c["csoport"] == g) for g, _ in GROUPS}
    counts = {g: sum(c["csoport"] == g for c in CRITERIA) for g, _ in GROUPS}
    group_lines = [(f"• {g}: {counts[g]} szempont, {weights[g]}% súly", None) for g, _ in GROUPS]
    top, second = TIERS[-1], TIERS[-2]
    gate_names = " és ".join(f"„{c['nev']}” ≥ {VALUE_GATE[c['kod']]}" for c in CRITERIA if c["kod"] in VALUE_GATE)
    tier_text = "; ".join(f"{t['szint']} ≥ {t['min']}" for t in reversed(TIERS))
    mult_text = "; ".join(f"{t['szint']} = {t['szorzo']:.2f}".replace(".", ",") for t in reversed(TIERS))
    return [
        ("ICP ügyfélminősítő – útmutató", TITLE_FONT),
        ("", None),
        ("Mire jó?", SECTION_FONT),
        (f"Az ügyfeleket {len(CRITERIA)} szempont alapján 0–100 pontra értékeli, szintbe sorolja, rangsorolja, "
         "és havi díjat javasol az online marketing munkára.", None),
        ("", None),
        ("Kiket tekintünk ideális ügyfélnek?", SECTION_FONT),
        ("Felelősségteljes, fenntartható, zöld szemléletű cégeket, amelyek elsősorban szükségleteket elégítenek "
         "ki, nem igényeket:", None),
        ("• Emberközpontú szolgáltatók (egészség, oktatás, mentális jóllét, fejlesztés)", None),
        ("• Tudatos termelők és alkotók (fenntartható termékek, kézműves márkák, gazdaságok)", None),
        ("• Hatásközpontú szervezetek (közösségi, környezeti és társadalmi projektek)", None),
        (f"Értékkapu: a(z) {top['szint']} szinthez ({top['nev']}) {gate_names} pont kell. Aki nem jut át, "
         f"legfeljebb {second['szint']} szintet kaphat, bármilyen magas is a pontszáma.", None),
        ("", None),
        ("Használat", SECTION_FONT),
        (f"1. Ügyfelek lap: ügyfelenként egy sor. Minden értékelés 1–{SCALE_MAX} között, ahol a {SCALE_MAX} "
         "mindig a számunkra kedvező (pl. „Vásárlószerzés költsége”: 10 = olcsó, könnyű). "
         f"Az üres mező hiányzó adatnak számít (alapból {MISSING_POINTS} pont), ezért érdemes mindent kitölteni. "
         "A fejlécek színe a szempontcsoportot jelöli; a Beállítások lap Leírás oszlopa minden szempontot elmagyaráz.",
         None),
        ("2. Rangsor lap: a pontszám szerinti sorrend, szinttel, értékkapuval, csoportpontszámokkal és "
         "díjjavaslattal. Csak olvasható.", None),
        ("3. Eredmény lap: szempontonkénti pontok és a díjszámítás részletei. Szűrhető.", None),
        ("4. Beállítások lap: a súlyok, sávhatárok, szegmensek, értékkapu, szintek és díjparaméterek (sárga cellák).",
         None),
        ("", None),
        ("Pontozás", SECTION_FONT),
        *group_lines,
        (f"ICP pontszám = 100 × Σ súly × (pont − 1) / ({SCALE_MAX - 1} × Σ súly), így 0 és 100 közé esik. "
         "A csoportpontszámok ugyanígy, csak a csoport szempontjaival számolnak.", None),
        (f"Szintek: {tier_text}.", None),
        ("A „Leggyengébb terület” az a szempont, ahol a legtöbb súlyozott pont veszik el – "
         "itt érdemes fejleszteni vagy tárgyalni.", None),
        ("", None),
        ("Díjjavaslat", SECTION_FONT),
        ("Munkadíj = csatornák alapdíja (Google, Meta, egyéb) + becsült havi óra × óradíj.", None),
        ("Javasolt díj = a három közül a legnagyobb, felfelé kerekítve: "
         "munkadíj × szintszorzó | hirdetési keret × keretarány | minimumdíj.", None),
        (f"A szintszorzó a kockázatot és a többletenergiát árazza: {mult_text}.", None),
        ("Az „Eltérés” a jelenlegi díjhoz képest mutatja, hol érdemes árat emelni (pozitív) "
         "vagy hol fizet az ügyfél a javasoltnál többet (negatív).", None),
        ("Jutalék: a Beállítások „Jutalék % (később)” oszlopa most 0. Ha kitöltöd, a mért bevétel "
         "arányában számol opcionális jutalékot.", None),
        ("", None),
        ("Google Táblázatok", SECTION_FONT),
        ("A fájl feltölthető Google Drive-ra és megnyitható Google Táblázatként; a képletek ott is működnek.", None),
    ]


def build_guide(ws):
    ws.column_dimensions["A"].width = 130
    for i, (text, font) in enumerate(_guide(), 1):
        cell = ws.cell(row=i, column=1, value=text)
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        if font:
            cell.font = font


def build_workbook(clients=()):
    wb = Workbook()
    guide = wb.active
    guide.title = S_HELP
    ws_in = wb.create_sheet(S_IN)
    ws_rank = wb.create_sheet(S_RANK)
    ws_res = wb.create_sheet(S_RES)
    ws_set = wb.create_sheet(S_SET)

    lay, rc = Layout(), ResCols()
    build_guide(guide)
    build_settings(ws_set, lay)
    build_input(ws_in, lay, list(clients))
    build_results(ws_res, lay, rc)
    build_ranking(ws_rank, rc)

    for ws, color in ((ws_in, "FAB219"), (ws_rank, "2E981B"), (ws_res, "134F6C"), (ws_set, "898781")):
        ws.sheet_properties.tabColor = color
    wb.active = 1
    return wb
