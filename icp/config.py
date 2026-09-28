"""Az ICP-pontozás alapértelmezett beállításai.

Ez az egyetlen forrás: ebből készül a táblázat "Beállítások" lapja, és ezt
használja a Python-motor is. Minden érték JSON-ba szerializálható, így később
adatbázisba vagy a belső rendszer konfigurációjába áttehető.
"""

# Az "Ügyfelek" lap oszlopai: (kulcs, fejléc, szélesség, típus)
# típus: szoveg | szam | penz | igennem | skala | lista
INPUT_COLUMNS = [
    ("nev", "Ügyfél neve", 26, "szoveg"),
    ("iparag", "Iparág", 24, "lista"),
    ("arbevetel", "Éves árbevétel (M Ft)", 14, "szam"),
    ("keret", "Havi hirdetési keret (Ft)", 16, "penz"),
    ("google", "Google Ads (igen/nem)", 12, "igennem"),
    ("meta", "Meta (igen/nem)", 12, "igennem"),
    ("egyeb_csatorna", "Egyéb csatornák száma", 12, "szam"),
    ("orak", "Becsült havi munkaóra", 12, "szam"),
    ("roas", "Jelenlegi ROAS", 11, "szam"),
    ("keses", "Átlagos fizetési késés (nap)", 13, "szam"),
    ("hossz", "Együttműködés hossza (hónap)", 13, "szam"),
    ("erettseg", "Marketing-érettség (1–5)", 12, "skala"),
    ("egyuttmukodes", "Döntéshozó, együttműködés (1–5)", 14, "skala"),
    ("novekedes", "Növekedési potenciál (1–5)", 12, "skala"),
    ("ajanlas", "Ajánlási, referenciaérték (1–5)", 13, "skala"),
    ("kezelhetoseg", "Kezelhetőség (1–5)", 12, "skala"),
    ("jelenlegi_dij", "Jelenlegi havi díj (Ft)", 15, "penz"),
    ("bevetel", "Havi mért bevétel (Ft)", 15, "penz"),
    ("megjegyzes", "Megjegyzés", 30, "szoveg"),
]

# Kritériumok. A súlyok összege 100.
# tipus: "sav" = számérték sávokba sorolva, "skala" = kézi 1–5 értékelés,
# "lista" = kategória pontértéke.
# A sávok (alsó határ, pont) párok, növekvő alsó határral.
CRITERIA = [
    {
        "kod": "iparag",
        "nev": "Iparági illeszkedés",
        "csoport": "Illeszkedés",
        "tipus": "lista",
        "suly": 10,
        "leiras": "Mennyire mérhető és skálázható az online hirdetés az iparágban.",
    },
    {
        "kod": "arbevetel",
        "nev": "Éves árbevétel",
        "csoport": "Illeszkedés",
        "tipus": "sav",
        "suly": 10,
        "savok": [(0, 1), (50, 2), (200, 3), (500, 4), (2000, 5)],
        "leiras": "Millió Ft-ban. A cég mérete és fizetőképessége.",
    },
    {
        "kod": "keret",
        "nev": "Havi hirdetési keret",
        "csoport": "Illeszkedés",
        "tipus": "sav",
        "suly": 15,
        "savok": [(0, 1), (300_000, 2), (1_000_000, 3), (3_000_000, 4), (10_000_000, 5)],
        "leiras": "Havi médiaköltés Ft-ban (Google + Meta + egyéb).",
    },
    {
        "kod": "erettseg",
        "nev": "Marketing-érettség",
        "csoport": "Illeszkedés",
        "tipus": "skala",
        "suly": 10,
        "leiras": "Van-e konverziómérés, CRM, saját marketinges, tiszta cél. 1 = semmi, 5 = teljes.",
    },
    {
        "kod": "roas",
        "nev": "Jelenlegi ROAS",
        "csoport": "Teljesítmény",
        "tipus": "sav",
        "suly": 10,
        "savok": [(0, 1), (1, 2), (2, 3), (4, 4), (6, 5)],
        "leiras": "Mért bevétel / költés (a dashboardról). Ha nem mérhető, maradjon üresen.",
    },
    {
        "kod": "novekedes",
        "nev": "Növekedési potenciál",
        "csoport": "Teljesítmény",
        "tipus": "skala",
        "suly": 10,
        "leiras": "Mennyit nőhet a keret és a bevétel 12 hónap alatt. 1 = stagnál, 5 = erősen nő.",
    },
    {
        "kod": "keses",
        "nev": "Fizetési fegyelem",
        "csoport": "Kapcsolat",
        "tipus": "sav",
        "suly": 10,
        "savok": [(0, 5), (1, 4), (8, 3), (16, 2), (31, 1)],
        "leiras": "Átlagos késés napokban. 0 = mindig időben fizet.",
    },
    {
        "kod": "hossz",
        "nev": "Együttműködés hossza",
        "csoport": "Kapcsolat",
        "tipus": "sav",
        "suly": 5,
        "savok": [(0, 1), (3, 2), (6, 3), (12, 4), (24, 5)],
        "leiras": "Hány hónapja dolgozunk együtt.",
    },
    {
        "kod": "egyuttmukodes",
        "nev": "Döntéshozó, együttműködés",
        "csoport": "Kapcsolat",
        "tipus": "skala",
        "suly": 10,
        "leiras": "Elérhető-e a döntéshozó, gyors-e a jóváhagyás. 1 = nehézkes, 5 = kiváló.",
    },
    {
        "kod": "kezelhetoseg",
        "nev": "Kezelhetőség",
        "csoport": "Kapcsolat",
        "tipus": "skala",
        "suly": 5,
        "leiras": "Kiszolgálási terhelés. 1 = sok egyeztetés, sürgetés; 5 = önállóan futtatható.",
    },
    {
        "kod": "ajanlas",
        "nev": "Ajánlási, referenciaérték",
        "csoport": "Kapcsolat",
        "tipus": "skala",
        "suly": 5,
        "leiras": "Ajánl-e minket, használható-e esettanulmánynak. 1 = nem, 5 = aktívan ajánl.",
    },
]

# Iparágak pontértéke (lista típusú kritériumhoz).
INDUSTRIES = [
    ("E-kereskedelem, webshop", 5),
    ("Online szolgáltatás, SaaS", 5),
    ("Oktatás, képzés", 4),
    ("Egészség, szépség", 4),
    ("Turizmus, vendéglátás", 4),
    ("B2B szolgáltatás", 3),
    ("Helyi szolgáltatás", 3),
    ("Ingatlan, építőipar", 3),
    ("Gyártás, ipar", 2),
    ("Egyéb", 2),
]

# Hiányzó adat esetén adott pont (1–5). Szándékosan az átlag alatti, hogy a
# hiányos adatlap ne jusson jobb helyre, mint a kitöltött.
MISSING_POINTS = 2

# Szintek: növekvő minimum pontszám szerint.
# szorzo: a munkadíj szorzója (a kockázatosabb, több energiát igénylő ügyfél drágább).
# jutalek: jövőbeli, opcionális jutalék a mért bevétel %-ában (egyelőre 0).
TIERS = [
    {"szint": "D", "min": 0, "nev": "Nem ideális", "szorzo": 1.30, "jutalek": 0.0},
    {"szint": "C", "min": 50, "nev": "Fejlesztendő", "szorzo": 1.15, "jutalek": 0.0},
    {"szint": "B", "min": 65, "nev": "Jó ügyfél", "szorzo": 1.05, "jutalek": 0.0},
    {"szint": "A", "min": 80, "nev": "Ideális ügyfél (ICP)", "szorzo": 1.00, "jutalek": 0.0},
]

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


def default_config():
    """A teljes alapértelmezett konfiguráció egy szótárban."""
    return {
        "criteria": CRITERIA,
        "industries": INDUSTRIES,
        "missing_points": MISSING_POINTS,
        "tiers": TIERS,
        "pricing": {k: v for k, (_, v) in PRICING.items()},
    }
