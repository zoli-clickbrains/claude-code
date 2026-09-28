"""Az ICP-ügyfélminősítő táblázat előállítása élő képletekkel.

A képletek Excelben, Google Táblázatokban és LibreOffice-ban is működnek
(VLOOKUP, INDEX/MATCH, SUMIF, COUNTIFS, RANK, LARGE, CEILING – nincs dinamikus
tömbképlet).
"""

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule, DataBarRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as col
from openpyxl.worksheet.datavalidation import DataValidation

from .config import (COMMISSION, CRITERIA, GROUPS, INPUT_COLUMNS, KNOCKOUT, LIST_CRITERIA, MAX_STATUS4,
                     MISSING_POINTS, PORTFOLIO_TARGETS, PRICING, SCALE_MAX, STATUS4_MIN_SYSTEM, STATUS4_RULE,
                     STATUSES, SYSTEM_GROUPS)
from .engine import ACCEPT, REJECT, WAITLIST

MAX_ROWS = 200  # ennyi ügyfélsorra készülnek elő a képletek
FIRST, LAST = 2, MAX_ROWS + 1

S_IN, S_SET, S_RES, S_RANK, S_PORT, S_HELP = (
    "Ügyfelek", "Beállítások", "Eredmény", "Rangsor", "Portfólió", "Útmutató")

FT = '#,##0" Ft"'
PCT = "0%"
YES, NO = "Igen", "Nem"

NAVY = "032D42"
GROUP_COLORS = dict(GROUPS) | {"Díjazás": "5E5E5A"}
HEAD_FONT = Font(bold=True, color="FFFFFF")
INPUT_FILL = PatternFill("solid", fgColor="FFF8E1")
TITLE_FONT = Font(bold=True, size=14, color=NAVY)
SECTION_FONT = Font(bold=True, size=12, color=NAVY)
THIN = Border(bottom=Side(style="thin", color="C3C2B7"))
STATUS_FILLS = {1: "C8EFC0", 2: "E3F4D9", 3: "FDEBC2", 4: "F9D3B0", 5: "F6CFCF"}


def fill(color):
    return PatternFill("solid", fgColor=color)


def ref(sheet, cell):
    return f"'{sheet}'!{cell}"


def abs_ref(sheet, column, row):
    return ref(sheet, f"${column}${row}")


def col_range(sheet, column, first, last):
    return f"{ref(sheet, f'${column}${first}')}:${column}${last}"


def _head(cell, color=NAVY):
    cell.fill, cell.font = fill(color), HEAD_FONT
    cell.alignment = Alignment(wrap_text=True, vertical="center")


def _header(ws, row, labels, widths=None, colors=None):
    for i, label in enumerate(labels, start=1):
        _head(ws.cell(row=row, column=i, value=label), (colors or {}).get(i, NAVY))
        if widths:
            ws.column_dimensions[col(i)].width = widths[i - 1]
    ws.row_dimensions[row].height = 60


def _section(ws, row, text, column=1):
    ws.cell(row=row, column=column, value=text).font = SECTION_FONT


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

        # Lista típusú kritériumok táblái egymás alatt, új opcióknak hagyott üres sorokkal.
        self.lists = {}
        row = self.group_first + len(GROUPS) + 4
        for c in LIST_CRITERIA:
            first, last = row, row + len(c["opciok"]) + 2
            self.lists[c["kod"]] = (first, last)
            row = last + 4

        # Paraméterek: (kulcs, címke, érték, formátum)
        names = {c["kod"]: c["nev"] for c in CRITERIA}
        self.params = [("missing", f"Hiányzó adat pontja (1–{SCALE_MAX})", MISSING_POINTS, "0")]
        self.params += [(f"ko:{k}", f"Kizáró: „{names[k]}” pontja legfeljebb", v, "0") for k, v in KNOCKOUT.items()]
        self.params += [("s4_min", "4. státusz: rendszerpontszám legalább", STATUS4_MIN_SYSTEM, "0")]
        self.params += [(f"s4:{g}", f"4. státusz: „{g}” csoportpontszám legalább", v, "0")
                        for g, v in STATUS4_RULE.items()]
        self.params += [("max4", "4. státuszú ügyfelek maximális száma", MAX_STATUS4, "0"),
                        ("comm_min", "Jutalékos modell: nyitottság pontja legalább", COMMISSION["kuszob"], "0"),
                        ("comm_red", "Jutalékos modell: fix díj csökkentése", COMMISSION["fix_csokkentes"], PCT)]
        self.params += [(k, label, value, PCT if k == "keret_arany" else FT) for k, (label, value) in PRICING.items()]
        target_labels = {"top2_min": "Cél: 1–2. státusz aránya legalább", "top2_max": "Cél: 1–2. státusz aránya legfeljebb",
                         "s3_min": "Cél: 3. státusz aránya legalább", "s3_max": "Cél: 3. státusz aránya legfeljebb"}
        self.params += [(k, target_labels[k], v, PCT) for k, v in PORTFOLIO_TARGETS.items()]
        self.param_row0 = row + 1
        self.p = {key: abs_ref(S_SET, "B", self.param_row0 + i) for i, (key, *_) in enumerate(self.params)}

        self.status_first = self.param_row0 + len(self.params) + 4
        self.status_last = self.status_first + len(STATUSES) - 1

        self.weight = {c["kod"]: abs_ref(S_SET, "E", self.crit_first + i) for i, c in enumerate(CRITERIA)}
        self.names = col_range(S_SET, "B", self.crit_first, self.crit_last)
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

    def list_table(self, kod):
        first, last = self.lists[kod]
        return f"{ref(S_SET, f'$A${first}')}:$B${last}"

    def list_names(self, kod):
        first, last = self.lists[kod]
        return col_range(S_SET, "A", first, last)

    def status_col(self, column):
        return col_range(S_SET, column, self.status_first, self.status_last)

    def status_cell(self, column, statusz):
        return abs_ref(S_SET, column, self.status_first + statusz - 1)


STATUS_TABLE = [("Státusz", "statusz", 8, "0"), ("Megnevezés", "nev", 30, None),
                ("Minimum rendszerpontszám", "min", 12, "0"), ("Díjszorzó", "szorzo", 10, "0.00"),
                ("Jutalék % (jutalékos modellben)", "jutalek", 12, "0.0%"), ("Díjmodell", "dijmodell", 40, None),
                ("Leírás", "leiras", 90, None)]
S_MIN, S_NAME, S_MULT, S_PCT, S_MODEL = "C", "B", "D", "E", "F"


def build_settings(ws, lay):
    ws["A1"] = "Beállítások – szempontok, súlyok, sávok, státuszok, díjazás"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = (f"Minden szempont 1–{SCALE_MAX} pontot kap, a {SCALE_MAX} mindig a számunkra kedvező érték. "
                "A sárga cellák szabadon módosíthatók; minden eredmény azonnal újraszámolódik. "
                "A súlyozás ideiglenes váz.")

    _section(ws, 4, "Szempontok és súlyok")
    for j, label in enumerate(["Kód", "Szempont", "Csoport", "Típus", "Súly", "Leírás"], 1):
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
    for w, c in zip([44, 30, 18, 11, 8, 70], "ABCDEF"):
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

    for c in LIST_CRITERIA:
        first, last = lay.lists[c["kod"]]
        _section(ws, first - 2, f"{c['nev']}: választható értékek pontja")
        ws.cell(row=first - 2, column=3, value="(új érték az üres sorokba vehető fel)")
        for j, label in enumerate(["Érték", f"Pont (1–{SCALE_MAX})"], 1):
            _head(ws.cell(row=first - 1, column=j, value=label))
        for r in range(first, last + 1):
            for j in (1, 2):
                ws.cell(row=r, column=j).fill = INPUT_FILL
        for i, (name, pts) in enumerate(c["opciok"]):
            ws.cell(row=first + i, column=1, value=name)
            ws.cell(row=first + i, column=2, value=pts)

    _section(ws, lay.param_row0 - 1, "Kizárás, 4. státusz, jutalék, díjazás, portfólió-célok")
    for i, (_, label, value, fmt) in enumerate(lay.params):
        r = lay.param_row0 + i
        ws.cell(row=r, column=1, value=label)
        cell = ws.cell(row=r, column=2, value=value)
        cell.fill, cell.number_format = INPUT_FILL, fmt

    _section(ws, lay.status_first - 2, "Státuszok, díjszorzók és jutalékok")
    for j, (label, *_rest) in enumerate(STATUS_TABLE, 1):
        _head(ws.cell(row=lay.status_first - 1, column=j, value=label))
    for i, s in enumerate(STATUSES):
        r = lay.status_first + i
        for j, (_, key, _, fmt) in enumerate(STATUS_TABLE, 1):
            cell = ws.cell(row=r, column=j, value=s[key])
            if key in ("min", "szorzo", "jutalek", "dijmodell") and s[key] is not None:
                cell.fill = INPUT_FILL
            if fmt:
                cell.number_format = fmt
            cell.alignment = Alignment(wrap_text=True, vertical="top")
    ws.cell(row=lay.status_last + 1, column=1, value=(
        "Az 1–3. státuszt a rendszerpontszám (marketingezhetőség + üzleti illeszkedés) dönti el, csökkenő "
        "határokkal. A 4. státuszt csak az kapja, aki a 3. határ alatt van, de a fenti feltételeknek megfelel; "
        "mindenki más 5. státuszú, és elutasítjuk."
    )).font = Font(italic=True)
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
        "igennem": DataValidation(type="list", formula1='"igen,nem"', allow_blank=True),
        "skala": DataValidation(type="whole", operator="between", formula1="1", formula2=str(SCALE_MAX),
                                allow_blank=True),
        "szam": DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0", allow_blank=True),
        "penz": DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0", allow_blank=True),
    }
    validations["skala"].error = f"1 és {SCALE_MAX} közötti egész számot adj meg."
    for c in LIST_CRITERIA:
        validations[c["kod"]] = DataValidation(type="list", formula1=lay.list_names(c["kod"]), allow_blank=True)
    for dv in validations.values():
        dv.showErrorMessage = True
        ws.add_data_validation(dv)
    for j, (key, _, _, kind, _) in enumerate(INPUT_COLUMNS, 1):
        dv = validations.get(key if kind == "lista" else kind)
        if dv:
            dv.add(f"{col(j)}{FIRST}:{col(j)}{LAST}")
        if kind == "penz":
            for r in range(FIRST, LAST + 1):
                ws.cell(row=r, column=j).number_format = FT


def input_cell(key, row):
    idx = [k for k, *_ in INPUT_COLUMNS].index(key) + 1
    return ref(S_IN, f"{col(idx)}{row}")


def criterion_formula(c, lay, row):
    x = input_cell(c["kod"], row)
    miss = lay.p["missing"]
    if c["tipus"] == "sav":
        rng, first = lay.bands[c["kod"]]
        body = f"VLOOKUP(MAX({x},{first}),{rng},2,TRUE)"
    elif c["tipus"] == "skala":
        body = f"MIN({SCALE_MAX},MAX(1,ROUND({x},0)))"
    else:
        body = f"IFERROR(VLOOKUP(TRIM({x}),{lay.list_table(c['kod'])},2,FALSE),{miss})"
    return f'=IF($A{row}="","",IF({x}="",{miss},{body}))'


# -------------------------------------------------------------------- Eredmény

RES_FIXED = [
    ("score", "ICP pontszám (0–100)", 12, "0.0"),
    ("system", "Rendszerpontszám (0–100)", 12, "0.0"),
    ("rank", "Rang", 7, "0"),
    ("knockout", "Kizáró feltétel", 10, None),
    ("status", "Státusz", 8, "0"),
    ("status_name", "Státusz megnevezése", 30, None),
    ("decision", "Döntés", 18, None),
    ("fee_model", "Díjmodell", 34, None),
    ("best", "Legerősebb terület", 24, None),
    ("worst", "Leggyengébb terület", 24, None),
    ("labour", "Munkadíj (Ft/hó)", 14, FT),
    ("mult", "Státuszszorzó", 10, "0.00"),
    ("budget_min", "Keretarányos minimum (Ft/hó)", 15, FT),
    ("fee", "Javasolt fix havi díj (Ft)", 15, FT),
    ("basis", "Díj alapja", 24, None),
    ("current", "Jelenlegi díj (Ft)", 14, FT),
    ("diff", "Eltérés (Ft)", 13, FT),
    ("diff_pct", "Eltérés (%)", 10, PCT),
    ("comm_model", "Jutalékos modell ajánlható", 11, None),
    ("comm_fee", "Jutalékos modell: fix díj (Ft/hó)", 15, FT),
    ("comm_pct", "Jutalékos modell: jutalék %", 11, "0.0%"),
    ("comm_expected", "Várható jutalék (Ft/hó)", 15, FT),
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

    def column(self, key):
        c = self.fixed[key]
        return col_range(S_RES, c, FIRST, LAST)


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

    f, p = rc.fixed, lay.p
    status_rng, key_rng = rc.column("status"), col_range(S_RES, rc.sort_key, FIRST, LAST)
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

        score = f"{f['score']}{r}"
        status = f"{f['status']}{r}"
        ws[score] = f'=IF({blank},"",{_weighted_formula(CRITERIA, lay, rc, r, lay.total_weight)})'
        ws[f"{f['rank']}{r}"] = f'=IF({blank},"",RANK({score},${f["score"]}${FIRST}:${f["score"]}${LAST}))'
        system = f"{f['system']}{r}"
        system_criteria = [c for c in CRITERIA if c["csoport"] in SYSTEM_GROUPS]
        system_total = "(" + "+".join(lay.group_weight[g] for g in SYSTEM_GROUPS) + ")"
        ws[system] = f'=IF({blank},"",{_weighted_formula(system_criteria, lay, rc, r, system_total)})'

        # Kizárni csak kitöltött adat alapján lehet.
        knock = ",".join(f'AND({input_cell(k, r)}<>"",{rc.points[k]}{r}<={p[f"ko:{k}"]})' for k in KNOCKOUT)
        ws[f"{f['knockout']}{r}"] = f'=IF({blank},"",IF(OR({knock}),"{YES}","{NO}"))'
        strong = ",".join([f"{system}>={p['s4_min']}"] + [f"{rc.groups[g]}{r}>={p[f's4:{g}']}" for g in STATUS4_RULE])
        by_score = f"IF(AND({strong}),4,5)"
        for s in reversed(STATUSES[:3]):
            by_score = f"IF({system}>={lay.status_cell(S_MIN, s['statusz'])},{s['statusz']},{by_score})"
        ws[status] = f'=IF({blank},"",IF({f["knockout"]}{r}="{YES}",5,{by_score}))'
        ws[f"{f['status_name']}{r}"] = f'=IF({blank},"",INDEX({lay.status_col(S_NAME)},{status}))'
        ws[f"{f['fee_model']}{r}"] = f'=IF({blank},"",INDEX({lay.status_col(S_MODEL)},{status}))'
        ws[f"{f['decision']}{r}"] = (
            f'=IF({blank},"",IF({status}=5,"{REJECT}",IF(AND({status}=4,'
            f'COUNTIFS({status_rng},4,{key_rng},">"&{rc.sort_key}{r})>={p["max4"]}),"{WAITLIST}","{ACCEPT}")))')
        ws[f"{f['mult']}{r}"] = f'=IF(OR({blank},{status}=5),"",INDEX({lay.status_col(S_MULT)},{status}))'

        s_rng = f"{rc.strength[0]}{r}:{rc.strength[-1]}{r}"
        l_rng = f"{rc.lost[0]}{r}:{rc.lost[-1]}{r}"
        ws[f"{f['best']}{r}"] = f'=IF({blank},"",INDEX({lay.names},MATCH(MAX({s_rng}),{s_rng},0)))'
        ws[f"{f['worst']}{r}"] = (f'=IF({blank},"",IF(MAX({l_rng})=0,"—",'
                                  f'INDEX({lay.names},MATCH(MAX({l_rng}),{l_rng},0))))')
        for i, c in enumerate(CRITERIA):
            pt = f"{rc.points[c['kod']]}{r}"
            ws[f"{rc.strength[i]}{r}"] = f'=IF({blank},"",{pt}+{lay.weight[c["kod"]]}/1000)'
            ws[f"{rc.lost[i]}{r}"] = f'=IF({blank},"",{lay.weight[c["kod"]]}*({SCALE_MAX}-{pt}))'
        ws[f"{rc.sort_key}{r}"] = f'=IF({blank},"",{score}-ROW()/100000)'

        def yes(key):
            return f'(TRIM({input_cell(key, r)})="igen")'

        ws[f"{f['labour']}{r}"] = (
            f'=IF({blank},"",{p["google_alap"]}*{yes("google")}+{p["meta_alap"]}*{yes("meta")}'
            f'+{p["egyeb_alap"]}*N({input_cell("egyeb_csatorna", r)})+{p["oradij"]}*N({input_cell("orak", r)}))'
        )
        ws[f"{f['budget_min']}{r}"] = f'=IF({blank},"",N({input_cell("keret", r)})*{p["keret_arany"]})'
        rejected = f"OR({blank},{status}=5)"
        by_labour = f"{f['labour']}{r}*{f['mult']}{r}"
        raw = f"MAX({p['min_dij']},{by_labour},{f['budget_min']}{r})"
        fee = f"{f['fee']}{r}"
        ws[fee] = f'=IF({rejected},"",CEILING(ROUND({raw},2),{p["kerekites"]}))'
        ws[f"{f['basis']}{r}"] = (
            f'=IF({blank},"",IF({status}=5,"Elutasítva",IF({raw}={by_labour},"Munkaigény × státuszszorzó",'
            f'IF({raw}={f["budget_min"]}{r},"Hirdetési keret arányában","Minimumdíj"))))'
        )
        cur_in = input_cell("jelenlegi_dij", r)
        cur = f"{f['current']}{r}"
        ws[cur] = f'=IF(OR({blank},{cur_in}=""),"",{cur_in})'
        ws[f"{f['diff']}{r}"] = f'=IF(OR({cur}="",{fee}=""),"",{fee}-{cur})'
        ws[f"{f['diff_pct']}{r}"] = f'=IF(OR({cur}="",{fee}="",{cur}=0),"",{f["diff"]}{r}/{cur})'

        pct = f"INDEX({lay.status_col(S_PCT)},{status})"
        model = f"{f['comm_model']}{r}"
        ws[model] = (f'=IF({rejected},"",IF(AND({pct}>0,{rc.points["jutalek_nyitottsag"]}{r}>={p["comm_min"]}),'
                     f'"{YES}","{NO}"))')
        ws[f"{f['comm_fee']}{r}"] = (f'=IF({model}="{YES}",CEILING(ROUND({fee}*(1-{p["comm_red"]}),2),'
                                     f'{p["kerekites"]}),"")')
        ws[f"{f['comm_pct']}{r}"] = f'=IF({model}="{YES}",{pct},"")'
        rev = input_cell("bevetel", r)
        ws[f"{f['comm_expected']}{r}"] = f'=IF(OR({model}<>"{YES}",{rev}=""),"",{rev}*{f["comm_pct"]}{r})'
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
    _status_colors(ws, f"{f['status']}{FIRST}:{f['status_name']}{LAST}", f"${f['status']}{FIRST}")
    _decision_colors(ws, f"{f['decision']}{FIRST}:{f['decision']}{LAST}")
    diff_rng = f"{f['diff']}{FIRST}:{f['diff_pct']}{LAST}"
    ws.conditional_formatting.add(diff_rng, CellIsRule(operator="greaterThan", formula=["0"],
                                                       font=Font(color="1C7A3E", bold=True)))
    ws.conditional_formatting.add(diff_rng, CellIsRule(operator="lessThan", formula=["0"],
                                                       font=Font(color="C23434", bold=True)))


def _status_colors(ws, rng, anchor):
    for statusz, color in STATUS_FILLS.items():
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f"{anchor}={statusz}"], fill=fill(color)))


def _decision_colors(ws, rng):
    top_left = rng.split(":")[0]
    for text, color, bold in ((ACCEPT, "1C7A3E", True), (WAITLIST, "B26B00", True), (REJECT, "C23434", True)):
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{top_left}="{text}"'],
                                                       font=Font(color=color, bold=bold)))


# --------------------------------------------------------------------- Rangsor

def build_ranking(ws, rc):
    fx = rc.fixed
    cols = [("Rang", fx["rank"], 7, "0", NAVY), ("Ügyfél", "A", 30, None, NAVY),
            ("ICP pontszám", fx["score"], 12, "0.0", NAVY), ("Rendszerpontszám", fx["system"], 12, "0.0", NAVY),
            ("Státusz", fx["status"], 8, "0", NAVY),
            ("Státusz megnevezése", fx["status_name"], 30, None, NAVY), ("Döntés", fx["decision"], 18, None, NAVY)]
    cols += [(g, rc.groups[g], 17, "0.0", color) for g, color in GROUPS]
    cols += [("Javasolt fix havi díj", fx["fee"], 16, FT, NAVY),
             ("Jelenlegi díj", fx["current"], 15, FT, NAVY),
             ("Eltérés (Ft)", fx["diff"], 14, FT, NAVY),
             ("Eltérés (%)", fx["diff_pct"], 10, PCT, NAVY),
             ("Jutalékos modell", fx["comm_model"], 11, None, NAVY),
             ("Jutalékos: fix díj", fx["comm_fee"], 15, FT, NAVY),
             ("Jutalékos: jutalék %", fx["comm_pct"], 11, "0.0%", NAVY),
             ("Díjmodell", fx["fee_model"], 34, None, NAVY),
             ("Leggyengébb terület", fx["worst"], 26, None, NAVY),
             ("Hiányzó adatok", fx["missing"], 10, "0", NAVY)]
    ws["A1"] = "ICP-rangsor – pontszám szerint csökkenő sorrendben"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = ("Automatikusan frissül az Ügyfelek lap alapján. A csoportoszlopok 0–100 között mutatják, "
                "hol erős az ügyfél. Holtversenynél a korábban felvett ügyfél áll elöl.")
    _header(ws, 4, [c[0] for c in cols], [c[2] for c in cols], {i: c[4] for i, c in enumerate(cols, 1)})
    helper = col(len(cols) + 2)
    ws.column_dimensions[helper].hidden = True
    ws[f"{helper}4"] = "sor"
    key_rng = col_range(S_RES, rc.sort_key, FIRST, LAST)
    for k in range(1, MAX_ROWS + 1):
        r = 4 + k
        ws[f"{helper}{r}"] = f'=IFERROR(MATCH(LARGE({key_rng},{k}),{key_rng},0),"")'
        for j, (_, src, _, fmt, _) in enumerate(cols, 1):
            cell = ws.cell(row=r, column=j,
                           value=f'=IF(${helper}{r}="","",INDEX({col_range(S_RES, src, FIRST, LAST)},${helper}{r}))')
            if fmt:
                cell.number_format = fmt
            cell.border = THIN
    ws.freeze_panes = "C5"
    last = 4 + MAX_ROWS
    _status_colors(ws, f"E5:F{last}", "$E5")
    _decision_colors(ws, f"G5:G{last}")
    ws.conditional_formatting.add(f"C5:D{last}", DataBarRule(
        start_type="num", start_value=0, end_type="num", end_value=100, color="2E981B"))


# ------------------------------------------------------------------- Portfólió

class PortCells:
    """A Portfólió lap cellái (a tesztek is ezeket olvassák)."""

    count_row0 = 5  # státuszonként egy sor: 5..9
    accepted = "C11"
    waitlist = "C12"
    top2, s3, s4 = 15, 16, 17  # ellenőrző sorok


def build_portfolio(ws, lay, rc):
    pc = PortCells
    ws["A1"] = "Portfólió – státuszeloszlás és célok"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = ("Az arányok az elfogadott ügyfelekre vonatkoznak (az elutasított és a várólistás ügyfelek nélkül). "
                "A célok a Beállítások lapon módosíthatók.")
    _header(ws, 4, ["Státusz", "Megnevezés", "Ügyfelek (db)", "Arány"], [10, 42, 16, 12])
    status, decision = rc.column("status"), rc.column("decision")
    for i, s in enumerate(STATUSES):
        r = pc.count_row0 + i
        k = s["statusz"]
        ws.cell(row=r, column=1, value=k)
        ws.cell(row=r, column=2, value=f"={lay.status_cell(S_NAME, k)}")
        if k == 5:
            ws.cell(row=r, column=3, value=f"=COUNTIF({status},5)")
            ws.cell(row=r, column=4, value="elutasítva")
        else:
            ws.cell(row=r, column=3, value=f'=COUNTIFS({status},{k},{decision},"{ACCEPT}")')
            ws.cell(row=r, column=4, value=f"=IF({pc.accepted}=0,\"\",C{r}/{pc.accepted})").number_format = PCT
    _status_colors(ws, f"A{pc.count_row0}:B{pc.count_row0 + 4}", f"$A{pc.count_row0}")
    ws["B11"], ws[pc.accepted] = "Elfogadott ügyfelek összesen", f'=COUNTIF({decision},"{ACCEPT}")'
    ws["B12"], ws[pc.waitlist] = "Várólistán (4-es hely foglalt)", f'=COUNTIF({decision},"{WAITLIST}")'
    for c in ("B11", "B12"):
        ws[c].font = Font(bold=True)

    _header(ws, pc.top2 - 1, ["", "Ellenőrzés", "Cél", "Tény", "Értékelés"], [10, 42, 16, 12, 18])
    ws.row_dimensions[pc.top2 - 1].height = 20
    p = lay.p
    c1, c2, c3, c4 = (f"C{pc.count_row0 + i}" for i in range(4))
    checks = [
        (pc.top2, "1–2. státusz aránya", f'=TEXT({p["top2_min"]},"0%")&"–"&TEXT({p["top2_max"]},"0%")',
         f"=IF({pc.accepted}=0,\"\",({c1}+{c2})/{pc.accepted})", p["top2_min"], p["top2_max"]),
        (pc.s3, "3. státusz aránya", f'=TEXT({p["s3_min"]},"0%")&"–"&TEXT({p["s3_max"]},"0%")',
         f"=IF({pc.accepted}=0,\"\",{c3}/{pc.accepted})", p["s3_min"], p["s3_max"]),
    ]
    for r, label, target, actual, lo, hi in checks:
        ws[f"B{r}"], ws[f"C{r}"], ws[f"D{r}"] = label, target, actual
        ws[f"D{r}"].number_format = "0.0%"
        ws[f"E{r}"] = (f'=IF(D{r}="","—",IF(D{r}<{lo},"Kevesebb a célnál",'
                       f'IF(D{r}>{hi},"Több a célnál","Rendben")))')
    r = pc.s4
    ws[f"B{r}"], ws[f"C{r}"], ws[f"D{r}"] = "4. státusz (db)", f'="legfeljebb "&{p["max4"]}', f"={c4}"
    ws[f"E{r}"] = f'=IF(D{r}<={p["max4"]},"Rendben","Túl sok")'
    check_rng = f"E{pc.top2}:E{pc.s4}"
    ws.conditional_formatting.add(check_rng, FormulaRule(formula=[f'E{pc.top2}="Rendben"'],
                                                         font=Font(color="1C7A3E", bold=True)))
    ws.conditional_formatting.add(check_rng, FormulaRule(formula=[f'AND(E{pc.top2}<>"Rendben",E{pc.top2}<>"—")'],
                                                         font=Font(color="C23434", bold=True)))

    row = pc.s4 + 3
    _section(ws, row, "Státuszok", 2)
    for s in STATUSES:
        row += 1
        ws[f"A{row}"] = s["statusz"]
        ws[f"B{row}"] = f"={lay.status_cell(S_NAME, s['statusz'])}"
        ws[f"B{row}"].font = Font(bold=True)
        row += 1
        ws[f"B{row}"] = f"={lay.status_cell('G', s['statusz'])}"
        ws.merge_cells(f"B{row}:E{row}")
        ws[f"B{row}"].alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[row].height = 32
    ws.column_dimensions["A"].width = 10


# -------------------------------------------------------------------- Útmutató

def _guide():
    weights = {g: sum(c["suly"] for c in CRITERIA if c["csoport"] == g) for g, _ in GROUPS}
    counts = {g: sum(c["csoport"] == g for c in CRITERIA) for g, _ in GROUPS}
    group_lines = [(f"• {g}: {counts[g]} szempont, {weights[g]}% súly", None) for g, _ in GROUPS]
    status_lines = []
    for s in STATUSES:
        rule = (f"rendszerpontszám ≥ {s['min']}" if s["min"] is not None else
                ("kiemelkedő értékek és kapcsolat, legfeljebb 1 ügyfél" if s["statusz"] == 4 else "minden más"))
        commission = f", jutalék {s['jutalek']:.0%}" if s["jutalek"] else ""
        mult = f"szorzó {s['szorzo']:.2f}".replace(".", ",") if s["szorzo"] else "nincs ajánlat"
        status_lines.append((f"{s['statusz']}. {s['nev']} ({rule}; {mult}{commission}): {s['leiras']}", None))
    s4_groups = " és ".join(f"„{g}” ≥ {v}" for g, v in STATUS4_RULE.items())
    return [
        ("ICP ügyfélminősítő – útmutató", TITLE_FONT),
        ("", None),
        ("Mire jó?", SECTION_FONT),
        (f"Az ügyfeleket {len(CRITERIA)} szempont alapján 0–100 pontra értékeli, rangsorolja, öt státuszba "
         "sorolja, és ebből díjat, illetve jutalékos modellt javasol.", None),
        ("", None),
        ("Kiket tekintünk ideális ügyfélnek?", SECTION_FONT),
        ("Felelősségteljes, fenntartható, zöld szemléletű cégeket, amelyek elsősorban szükségleteket elégítenek "
         "ki, nem igényeket:", None),
        ("• Emberközpontú szolgáltatók (egészség, oktatás, mentális jóllét, fejlesztés)", None),
        ("• Tudatos termelők és alkotók (fenntartható termékek, kézműves márkák, gazdaságok)", None),
        ("• Hatásközpontú szervezetek (közösségi, környezeti és társadalmi projektek)", None),
        ("Kizáró feltétel: " + "; ".join(f"„{c['nev']}” legfeljebb {KNOCKOUT[c['kod']]} pont" for c in CRITERIA
                                          if c["kod"] in KNOCKOUT)
         + " (a szegmensnél ez az „Értékeinkkel nem összeegyeztethető” érték). Az ilyen ügyfél automatikusan "
         "5. státuszú, és elutasítjuk. Hiányzó adat nem zár ki.", None),
        ("", None),
        ("Státuszok", SECTION_FONT),
        *status_lines,
        ("A státusz a marketingrendszer állapotát írja le, ezért a rendszerpontszám dönti el: a "
         f"{' és a '.join(SYSTEM_GROUPS)} csoport közös pontszáma (0–100). A rangsor a teljes ICP pontszám "
         "szerint készül.", None),
        (f"A 4. státusz feltétele: a 3. státusz határa alatt van, de rendszerpontszáma ≥ {STATUS4_MIN_SYSTEM}, "
         f"és {s4_groups}. Legfeljebb {MAX_STATUS4} ilyen ügyfelet vállalunk; a továbbiak várólistára kerülnek.",
         None),
        ("Célarány (Portfólió lap): az elfogadott ügyfelek 70–80%-a 1. vagy 2., 20–25%-a 3. státuszú legyen. "
         "Ha nem teljesül, a státuszhatárokat (Beállítások) vagy a portfóliót érdemes felülvizsgálni.", None),
        ("", None),
        ("Használat", SECTION_FONT),
        (f"1. Ügyfelek lap: ügyfelenként egy sor. Minden értékelés 1–{SCALE_MAX} között, ahol a {SCALE_MAX} "
         "mindig a számunkra kedvező (pl. „Vásárlószerzés költsége”: 10 = olcsó, könnyű). "
         f"Az üres mező hiányzó adatnak számít (alapból {MISSING_POINTS} pont), ezért érdemes mindent kitölteni. "
         "A fejlécek színe a szempontcsoportot jelöli; a Beállítások lap Leírás oszlopa minden szempontot elmagyaráz.",
         None),
        ("2. Rangsor lap: sorrend státusszal, döntéssel, csoportpontszámokkal, díj- és jutalékjavaslattal.", None),
        ("3. Portfólió lap: hány ügyfél van az egyes státuszokban, és teljesülnek-e a célarányok.", None),
        ("4. Eredmény lap: szempontonkénti pontok és a díjszámítás részletei. Szűrhető.", None),
        ("5. Beállítások lap: súlyok, sávhatárok, választható értékek, státuszhatárok, díjparaméterek (sárga cellák).",
         None),
        ("", None),
        ("Pontozás", SECTION_FONT),
        *group_lines,
        (f"ICP pontszám = 100 × Σ súly × (pont − 1) / ({SCALE_MAX - 1} × Σ súly), így 0 és 100 közé esik. "
         "A csoportpontszámok ugyanígy, csak a csoport szempontjaival számolnak. A súlyozás ideiglenes váz.", None),
        ("", None),
        ("Díj és jutalék", SECTION_FONT),
        ("Munkadíj = csatornák alapdíja (Google, Meta, egyéb) + becsült havi óra × óradíj.", None),
        ("Javasolt fix díj = a három közül a legnagyobb, felfelé kerekítve: "
         "munkadíj × státuszszorzó | hirdetési keret × keretarány | minimumdíj. Az 5. státusznál nincs ajánlat.", None),
        (f"Jutalékos modell: ha a státuszhoz jutalék tartozik (1–2.) és az ügyfél nyitott rá "
         f"(nyitottság ≥ {COMMISSION['kuszob']} pont), a fix díj {COMMISSION['fix_csokkentes']:.0%}-kal "
         "csökken, és a mért bevétel státusz szerinti %-a jutalékként jár.", None),
        ("", None),
        ("Google Táblázatok", SECTION_FONT),
        ("A fájl feltölthető Google Drive-ra és megnyitható Google Táblázatként; a képletek ott is működnek.", None),
    ]


def build_guide(ws):
    ws.column_dimensions["A"].width = 140
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
    ws_port = wb.create_sheet(S_PORT)
    ws_res = wb.create_sheet(S_RES)
    ws_set = wb.create_sheet(S_SET)

    lay, rc = Layout(), ResCols()
    build_guide(guide)
    build_settings(ws_set, lay)
    build_input(ws_in, lay, list(clients))
    build_results(ws_res, lay, rc)
    build_ranking(ws_rank, rc)
    build_portfolio(ws_port, lay, rc)

    for ws, color in ((ws_in, "FAB219"), (ws_rank, "2E981B"), (ws_port, "2E7D32"), (ws_res, "134F6C"),
                      (ws_set, "898781")):
        ws.sheet_properties.tabColor = color
    wb.active = 1
    return wb
