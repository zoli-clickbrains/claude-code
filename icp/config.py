"""Az ICP-pontozás alapértelmezett beállításai.

Ez az egyetlen forrás: ebből készül a táblázat "Beállítások" lapja, és ezt
használja a Python-motor is. Minden érték JSON-ba szerializálható, így később
adatbázisba vagy a belső rendszer konfigurációjába áttehető.

Minden kritérium 1–10 pontot kap, ahol a 10 mindig a számunkra kedvező érték
(pl. a "Vásárlószerzés költsége" kritériumnál 10 = olcsó, könnyű).
"""

SCALE_MAX = 10

# Kritériumcsoportok a megjelenítés sorrendjében: (név, fejlécszín)
GROUPS = [
    ("Értékek", "2E7D32"),
    ("Marketingezhetőség", "134F6C"),
    ("Üzleti illeszkedés", "6A4C93"),
    ("Kapcsolat", "8C5A14"),
]


def _scale(kod, nev, csoport, suly, leiras):
    return {"kod": kod, "nev": nev, "csoport": csoport, "tipus": "skala", "suly": suly, "leiras": leiras}


def _band(kod, nev, csoport, suly, savok, fejlec, leiras):
    return {"kod": kod, "nev": nev, "csoport": csoport, "tipus": "sav", "suly": suly, "savok": savok,
            "fejlec": fejlec, "leiras": leiras}


M = "Marketingezhetőség"
U = "Üzleti illeszkedés"
K = "Kapcsolat"

# A súlyok összege 100. Csoportonként: Értékek 20, Marketingezhetőség 35,
# Üzleti illeszkedés 22, Kapcsolat 23.
CRITERIA = [
    # --- Értékek: kiket tekintünk ideális ügyfélnek -------------------------
    {"kod": "szegmens", "nev": "Értékalapú szegmens", "csoport": "Értékek", "tipus": "lista", "suly": 12,
     "fejlec": "Értékalapú szegmens",
     "leiras": "Emberközpontú szolgáltató, tudatos termelő/alkotó vagy hatásközpontú szervezet-e."},
    _scale("felelosseg", "Felelősség, fenntarthatóság", "Értékek", 8,
           "Felelősségteljes, fenntartható, zöld; elsősorban szükségletet elégít ki, nem igényt. "
           "1 = egyáltalán nem, 10 = teljesen."),

    # --- Marketingezhetőség (17 szempont) ------------------------------------
    _scale("m_surgosseg", "Sürgősség", M, 3,
           "Mennyire azonnali az igény a termékre (fájdalomcsillapító vs. régi némafilm). 10 = nagyon sürgős."),
    _scale("m_piacmeret", "A piac mérete", M, 2,
           "Hányan akarják megvenni. 10 = óriási piac."),
    _scale("m_arazas", "Árazási potenciál", M, 3,
           "Mi a legmagasabb elérhető ár (piros csoki vs. repülőgép). 10 = nagyon magas ár érhető el."),
    _scale("m_vevoszerzes", "Vásárlószerzés költsége", M, 2,
           "Mennyi pénz és erőfeszítés egy vevő megszerzése. 10 = olcsó és könnyű (autópályás büfé), "
           "1 = drága és nehéz (állami tender)."),
    _scale("m_szallitas", "Leszállítás költsége", M, 2,
           "Mennyibe kerül a termék előállítása és leszállítása. 10 = olcsó (online), 1 = drága (gyár kell)."),
    _scale("m_egyediseg_piac", "Egyediség a piacon", M, 2,
           "Mennyire egyedi a kínálat, mennyire nehéz lemásolni. 10 = alig van konkurencia."),
    _scale("m_gyorsasag", "A piac gyorsasága", M, 1,
           "Milyen gyorsan hozható létre az ajánlat. 10 = akár holnap (fűnyírás), 1 = évek (vendéglő)."),
    _scale("m_befektetes", "Befektetési igény", M, 1,
           "Mennyi pénz és munka kell az eladásig. 10 = szinte semmi (ablakmosás), 1 = nagyon sok (olajkutatás)."),
    _scale("m_upsell_potencial", "Upsell-potenciál", M, 2,
           "Eladás után lehet-e még valamit eladni (borotva → borotvahab). 10 = sok további eladás."),
    _scale("m_orokzold", "Örökzöld potenciál", M, 2,
           "Mennyi plusz munka és pénz kell a szolgáltatás fenntartásához. 10 = szinte semmi, "
           "1 = a fenntartás felemészti (gyógyfürdő-példa)."),
    _scale("m_egyedi_tulajdonsag", "Egyedi tulajdonság", M, 2,
           "Van-e reklámozható egyedi tulajdonsága (életre szóló élesség, guriga nélküli tekercs). 10 = erős, egyértelmű."),
    _scale("m_demonstralhato", "Látványos demonstrálhatóság", M, 3,
           "Látványosan megmutatható-e a hatás (átmegy rajta a kocsi, előtte–utána). 10 = nagyon látványos."),
    _scale("m_dizajn", "Dizájn, vonzó megjelenés", M, 2,
           "Szép-e, vonzó-e a szemnek, van-e körítése. 10 = kiemelkedően szép."),
    _scale("m_egyszeru", "Egyszerű használat", M, 2,
           "Mennyire könnyű használni. 10 = nagyon egyszerű."),
    _scale("m_problema", "Problémát old meg", M, 3,
           "Problémát old meg (nyer) vagy megelőz (át kell fordítani). 10 = égető problémát old meg, 1 = csak megelőz."),
    _scale("m_ujdonsag", "Újdonság", M, 2,
           "Mennyire új a piacon (az újat lehet a legmagasabb áron eladni). 10 = teljesen új."),
    _scale("m_upsell_kinalat", "Van hozzá upsell-ajánlat", M, 1,
           "Van-e már konkrét, upsellben kínálható kiegészítő termék/szolgáltatás. 10 = kész upsell-kínálat."),

    # --- Üzleti illeszkedés ----------------------------------------------------
    _scale("iparagi", "Iparági illeszkedés", U, 3,
           "Mennyire mérhető és skálázható az online hirdetés az iparágban. 10 = kiválóan."),
    _band("arbevetel", "Éves árbevétel", U, 3,
          [(0, 1), (30, 3), (80, 5), (150, 7), (300, 9), (1000, 10)],
          "Éves árbevétel (M Ft)", "Millió Ft-ban. A cég mérete és fizetőképessége (kkv-khoz és szervezetekhez kalibrálva)."),
    _band("keret", "Havi hirdetési keret", U, 6,
          [(0, 1), (150_000, 3), (300_000, 5), (600_000, 7), (1_000_000, 9), (2_000_000, 10)],
          "Havi hirdetési keret (Ft)", "Havi médiaköltés Ft-ban (Google + Meta + egyéb), kkv-khoz és szervezetekhez kalibrálva."),
    _scale("erettseg", "Marketing-érettség", U, 4,
           "Van-e konverziómérés, CRM, saját marketinges, tiszta cél. 1 = semmi, 10 = teljes."),
    _band("roas", "Jelenlegi ROAS", U, 3,
          [(0, 1), (1, 3), (2, 5), (3, 7), (4, 9), (6, 10)],
          "Jelenlegi ROAS", "Mért bevétel / költés (a dashboardról). Ha nem mérhető, maradjon üresen."),
    _scale("novekedes", "Növekedési potenciál", U, 3,
           "Mennyit nőhet a keret és a bevétel 12 hónap alatt. 1 = stagnál, 10 = erősen nő."),

    # --- Kapcsolat, együttműködés ----------------------------------------------
    _band("keses", "Fizetési fegyelem", K, 4,
          [(0, 10), (1, 8), (8, 6), (16, 4), (31, 2), (61, 1)],
          "Átlagos fizetési késés (nap)", "Átlagos késés napokban. 0 = mindig időben fizet."),
    _band("hossz", "Együttműködés hossza", K, 1,
          [(0, 1), (3, 3), (6, 5), (12, 7), (24, 9), (36, 10)],
          "Együttműködés hossza (hónap)", "Hány hónapja dolgozunk együtt."),
    _scale("egyuttmukodes", "Döntéshozó, együttműködés", K, 3,
           "Elérhető-e a döntéshozó, gyors-e a jóváhagyás. 1 = nehézkes, 10 = kiváló."),
    _scale("kezelhetoseg", "Kezelhetőség", K, 2,
           "Kiszolgálási terhelés. 1 = sok egyeztetés, 10 = önállóan futtatható."),
    _scale("ajanlas", "Ajánlási, referenciaérték", K, 2,
           "Ajánl-e minket, használható-e esettanulmánynak. 1 = nem, 10 = aktívan ajánl."),
    _scale("elvarasok", "Reális elvárások", K, 4,
           "Reálisak-e az elvárásai az eredményekkel, határidőkkel, költségekkel szemben. 10 = teljesen."),
    _scale("emberi", "Emberi oldal", K, 4,
           "Bizalommal van felénk, nem mikromenedzsel, nem sürget. 10 = teljes bizalom."),
    _scale("szakmai", "Szakmai színvonal", K, 3,
           "Szakmailag magas szintet képvisel, arccal és névvel vállalja, amit csinál. 10 = kiemelkedő."),
]

# Az értékalapú szegmensek pontértéke (lista típusú kritérium).
SEGMENTS = [
    ("Emberközpontú szolgáltató (egészség, oktatás, mentális jóllét, fejlesztés)", 10),
    ("Tudatos termelő, alkotó (fenntartható termék, kézműves márka, gazdaság)", 10),
    ("Hatásközpontú szervezet (közösségi, környezeti, társadalmi projekt)", 10),
    ("Felelős szemléletű egyéb vállalkozás", 7),
    ("Általános vállalkozás", 4),
    ("Értékeinkkel nem összeegyeztethető", 1),
]

# Hiányzó adat esetén adott pont. Szándékosan az átlag alatti, hogy a
# hiányos adatlap ne jusson jobb helyre, mint a kitöltött.
MISSING_POINTS = 3

# Szintek: növekvő minimum pontszám szerint. A legfelső szint csak az
# értékkapun átjutó ügyfeleké; aki nem jut át, eggyel lejjebb kerül.
# szorzo: a munkadíj szorzója (a kockázatosabb, több energiát igénylő ügyfél drágább).
# jutalek: jövőbeli, opcionális jutalék a mért bevétel %-ában (egyelőre 0).
TIERS = [
    {"szint": "D", "min": 0, "nev": "Nem ideális", "szorzo": 1.30, "jutalek": 0.0},
    {"szint": "C", "min": 50, "nev": "Fejlesztendő", "szorzo": 1.15, "jutalek": 0.0},
    {"szint": "B", "min": 65, "nev": "Jó ügyfél", "szorzo": 1.05, "jutalek": 0.0},
    {"szint": "A", "min": 80, "nev": "Ideális ügyfél (ICP)", "szorzo": 1.00, "jutalek": 0.0},
]

# Értékkapu: a legfelső szinthez legalább ennyi pont kell ezeken a kritériumokon.
VALUE_GATE = {"szegmens": 8, "felelosseg": 7}

# Díjazási paraméterek (Ft, illetve arány).
PRICING = {
    "oradij": ("Óradíj (Ft/óra)", 12_000),
    "google_alap": ("Google Ads alapdíj (Ft/hó)", 60_000),
    "meta_alap": ("Meta alapdíj (Ft/hó)", 50_000),
    "egyeb_alap": ("Egyéb csatorna alapdíja (Ft/hó/csatorna)", 40_000),
    "min_dij": ("Minimum havi díj (Ft)", 150_000),
    "keret_arany": ("Díj minimuma a hirdetési keret arányában", 0.08),
    "kerekites": ("Kerekítés (Ft)", 10_000),
}


def _input_header(c):
    if "fejlec" in c:
        return c["fejlec"]
    return f"{c['nev']} (1–{SCALE_MAX})"


# Az "Ügyfelek" lap oszlopai: (kulcs, fejléc, szélesség, típus, csoport)
# típus: szoveg | szam | penz | igennem | skala | lista
_CRIT_INPUT_TYPES = {"lista": "lista", "skala": "skala", "sav": "szam"}
INPUT_COLUMNS = (
    [("nev", "Ügyfél neve", 28, "szoveg", None)]
    + [(c["kod"], _input_header(c), 34 if c["tipus"] == "lista" else 12,
        "penz" if c["kod"] == "keret" else _CRIT_INPUT_TYPES[c["tipus"]], c["csoport"])
       for c in CRITERIA]
    + [
        ("google", "Google Ads (igen/nem)", 11, "igennem", "Díjazás"),
        ("meta", "Meta (igen/nem)", 11, "igennem", "Díjazás"),
        ("egyeb_csatorna", "Egyéb csatornák száma", 11, "szam", "Díjazás"),
        ("orak", "Becsült havi munkaóra", 11, "szam", "Díjazás"),
        ("jelenlegi_dij", "Jelenlegi havi díj (Ft)", 15, "penz", "Díjazás"),
        ("bevetel", "Havi mért bevétel (Ft)", 15, "penz", "Díjazás"),
        ("megjegyzes", "Megjegyzés", 30, "szoveg", None),
    ]
)


def default_config():
    """A teljes alapértelmezett konfiguráció egy szótárban."""
    return {
        "scale_max": SCALE_MAX,
        "criteria": CRITERIA,
        "segments": SEGMENTS,
        "missing_points": MISSING_POINTS,
        "tiers": TIERS,
        "value_gate": VALUE_GATE,
        "pricing": {k: v for k, (_, v) in PRICING.items()},
    }
