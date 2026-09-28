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


def _list(kod, nev, csoport, suly, opciok, leiras):
    return {"kod": kod, "nev": nev, "csoport": csoport, "tipus": "lista", "suly": suly, "opciok": opciok,
            "fejlec": nev, "leiras": leiras}


M = "Marketingezhetőség"
U = "Üzleti illeszkedés"
K = "Kapcsolat"

# Ideiglenes váz-súlyozás (összesen 100): Értékek 20, Marketingezhetőség 30
# (15 × 2), Üzleti illeszkedés 20, Kapcsolat 30. A hirdetési keret súlya
# szándékosan alacsony.
CRITERIA = [
    # --- Értékek: kiket tekintünk ideális ügyfélnek -------------------------
    _list("szegmens", "Értékalapú szegmens", "Értékek", 10, [
        ("Emberközpontú szolgáltató (egészség, oktatás, mentális jóllét, fejlesztés)", 10),
        ("Tudatos termelő, alkotó (fenntartható termék, kézműves márka, gazdaság)", 10),
        ("Hatásközpontú szervezet (közösségi, környezeti, társadalmi projekt)", 10),
        ("Felelős szemléletű egyéb vállalkozás", 7),
        ("Általános vállalkozás", 4),
        ("Értékeinkkel nem összeegyeztethető", 1),
    ], "Emberközpontú szolgáltató, tudatos termelő/alkotó vagy hatásközpontú szervezet-e."),
    _scale("felelosseg", "Felelősség, fenntarthatóság", "Értékek", 10,
           "Felelősségteljes, fenntartható, zöld; elsősorban szükségletet elégít ki, nem igényt. "
           "1 = egyáltalán nem, 10 = teljesen."),

    # --- Marketingezhetőség (15 szempont) ------------------------------------
    _scale("m_surgosseg", "Sürgősség", M, 2,
           "Mennyire azonnali az igény a termékre (fájdalomcsillapító vs. régi némafilm). 10 = nagyon sürgős."),
    _scale("m_piacmeret", "A piac mérete", M, 2,
           "Hányan akarják megvenni. 10 = óriási piac."),
    _scale("m_arazas", "Árazási potenciál", M, 2,
           "Mi a legmagasabb elérhető ár (piros csoki vs. repülőgép). 10 = nagyon magas ár érhető el."),
    _scale("m_vevoszerzes", "Vásárlószerzés költsége", M, 2,
           "Mennyi pénz és erőfeszítés egy vevő megszerzése. 10 = olcsó és könnyű (autópályás büfé), "
           "1 = drága és nehéz (állami tender)."),
    _scale("m_szallitas", "Leszállítás költsége", M, 2,
           "Mennyibe kerül a termék előállítása és leszállítása. 10 = olcsó (online), 1 = drága (gyár kell)."),
    _scale("m_egyediseg", "Egyediség, egyedi tulajdonság", M, 2,
           "Mennyire egyedi a kínálat a piacon, mennyire nehéz lemásolni, és van-e reklámozható egyedi "
           "tulajdonsága (életre szóló élesség, guriga nélküli tekercs). 10 = alig van konkurencia, "
           "erős egyedi tulajdonsággal."),
    _scale("m_gyorsasag", "A piac gyorsasága", M, 2,
           "Milyen gyorsan hozható létre az ajánlat. 10 = akár holnap (fűnyírás), 1 = évek (vendéglő)."),
    _scale("m_befektetes", "Befektetési igény", M, 2,
           "Mennyi pénz és munka kell az eladásig. 10 = szinte semmi (ablakmosás), 1 = nagyon sok (olajkutatás)."),
    _scale("m_upsell", "Upsell-potenciál", M, 2,
           "Eladás után lehet-e még valamit eladni, és van-e hozzá upsellben kínálható kiegészítő "
           "(borotva → borotvahab; a frizbihez nincs). 10 = sok további eladás, kész upsell-kínálattal."),
    _scale("m_orokzold", "Örökzöld potenciál", M, 2,
           "Mennyi plusz munka és pénz kell a szolgáltatás fenntartásához. 10 = szinte semmi, "
           "1 = a fenntartás felemészti (gyógyfürdő-példa)."),
    _scale("m_demonstralhato", "Látványos demonstrálhatóság", M, 2,
           "Látványosan megmutatható-e a hatás (átmegy rajta a kocsi, előtte–utána). 10 = nagyon látványos."),
    _scale("m_dizajn", "Dizájn, vonzó megjelenés", M, 2,
           "Szép-e, vonzó-e a szemnek, van-e körítése. 10 = kiemelkedően szép."),
    _scale("m_egyszeru", "Egyszerű használat", M, 2,
           "Mennyire könnyű használni. 10 = nagyon egyszerű."),
    _scale("m_problema", "Problémát old meg", M, 2,
           "Problémát old meg (nyer) vagy megelőz (át kell fordítani). 10 = égető problémát old meg, 1 = csak megelőz."),
    _scale("m_ujdonsag", "Újdonság", M, 2,
           "Mennyire új a piacon (az újat lehet a legmagasabb áron eladni). 10 = teljesen új."),

    # --- Üzleti illeszkedés ----------------------------------------------------
    _scale("iparagi", "Iparági illeszkedés", U, 3,
           "Mennyire mérhető és skálázható az online hirdetés az iparágban. 10 = kiválóan."),
    _band("arbevetel", "Éves árbevétel", U, 3,
          [(0, 1), (30, 3), (80, 5), (150, 7), (300, 9), (1000, 10)],
          "Éves árbevétel (M Ft)",
          "Millió Ft-ban. A cég mérete és fizetőképessége (kkv-khoz és szervezetekhez kalibrálva)."),
    _band("keret", "Havi hirdetési keret", U, 2,
          [(0, 1), (150_000, 3), (300_000, 5), (600_000, 7), (1_000_000, 9), (2_000_000, 10)],
          "Havi hirdetési keret (Ft)",
          "Havi médiaköltés Ft-ban (Google + Meta + egyéb), kkv-khoz és szervezetekhez kalibrálva."),
    _scale("erettseg", "Marketing-érettség", U, 4,
           "Van-e konverziómérés, CRM, saját marketinges, tiszta cél. 1 = semmi, 10 = teljes."),
    _band("roas", "Jelenlegi ROAS", U, 3,
          [(0, 1), (1, 3), (2, 5), (3, 7), (4, 9), (6, 10)],
          "Jelenlegi ROAS", "Mért bevétel / költés (a dashboardról). Ha nem mérhető, maradjon üresen."),
    _scale("novekedes", "Növekedési potenciál", U, 3,
           "Mennyit nőhet a keret és a bevétel 12 hónap alatt. 1 = stagnál, 10 = erősen nő."),
    _list("jutalek_nyitottsag", "Nyitottság a jutalékos modellre", U, 2, [
        ("Igen, kifejezetten ezt szeretné", 10),
        ("Nyitott rá", 8),
        ("Megfontolja", 5),
        ("Nem, csak fix díjat szeretne", 2),
    ], "Gondolkodik-e jutalékos (teljesítményalapú) együttműködésben."),

    # --- Kapcsolat, együttműködés ----------------------------------------------
    _band("keses", "Fizetési fegyelem", K, 4,
          [(0, 10), (1, 8), (8, 6), (16, 4), (31, 2), (61, 1)],
          "Átlagos fizetési késés (nap)", "Átlagos késés napokban. 0 = mindig időben fizet."),
    _band("hossz", "Együttműködés hossza", K, 2,
          [(0, 1), (3, 3), (6, 5), (12, 7), (24, 9), (36, 10)],
          "Együttműködés hossza (hónap)", "Hány hónapja dolgozunk együtt."),
    _scale("egyuttmukodes", "Döntéshozó, együttműködés", K, 3,
           "Elérhető-e a döntéshozó, gyors-e a jóváhagyás. 1 = nehézkes, 10 = kiváló."),
    _scale("kezelhetoseg", "Kezelhetőség", K, 3,
           "Kiszolgálási terhelés. 1 = sok egyeztetés, 10 = önállóan futtatható."),
    _scale("ajanlas", "Ajánlási, referenciaérték", K, 3,
           "Ajánl-e minket, használható-e esettanulmánynak. 1 = nem, 10 = aktívan ajánl."),
    _scale("elvarasok", "Reális elvárások", K, 4,
           "Reálisak-e az elvárásai az eredményekkel, határidőkkel, költségekkel szemben. 10 = teljesen."),
    _scale("emberi", "Emberi oldal", K, 4,
           "Bizalommal van felénk, nem mikromenedzsel, nem sürget. 10 = teljes bizalom."),
    _scale("szakmai", "Szakmai színvonal", K, 4,
           "Szakmailag magas szintet képvisel, arccal és névvel vállalja, amit csinál. 10 = kiemelkedő."),
    _list("tapasztalat", "Korábbi ügynökségi tapasztalat", K, 3, [
        ("Igen, jó tapasztalattal", 10),
        ("Igen, vegyes tapasztalattal", 7),
        ("Nem, még nem dolgozott szakemberrel", 5),
        ("Igen, rossz tapasztalattal", 4),
    ], "Dolgozott-e már marketingügynökséggel vagy szakemberrel, és milyen tapasztalattal."),
]

LIST_CRITERIA = [c for c in CRITERIA if c["tipus"] == "lista"]

# Hiányzó adat esetén adott pont. Szándékosan az átlag alatti, hogy a
# hiányos adatlap ne jusson jobb helyre, mint a kitöltött.
MISSING_POINTS = 3

# A státusz a marketingrendszer állapotát írja le, ezért a "rendszerpontszám"
# dönti el: a SYSTEM_GROUPS csoportok súlyozott pontszáma (0–100). A többi
# csoport (értékek, kapcsolat) a 4. státusz feltétele. A rangsor továbbra is a
# teljes ICP pontszám szerint készül.
SYSTEM_GROUPS = ["Marketingezhetőség", "Üzleti illeszkedés"]

# Státuszok. Az 1–3. státuszt a rendszerpontszám dönti el (min = alsó határ).
# A 4. státuszt csak az kaphatja, aki a 3. határ alatt van, de az Értékek és a
# Kapcsolat csoportban kiemelkedő (STATUS4_RULE). Mindenki más az 5.-be kerül,
# és elutasítjuk.
# szorzo: a munkadíj szorzója; jutalek: a jutalékos modellben ajánlott % a mért bevételből.
STATUSES = [
    {"statusz": 1, "nev": "Stabil rendszer", "min": 70, "szorzo": 0.90, "jutalek": 0.05,
     "dijmodell": "Fenntartási díj; jutalékos modellben is ajánlható",
     "leiras": "A kampányok és a kapcsolódó marketingelemek jól működnek, az eredmények kiszámíthatók. "
               "A feladat elsősorban az ellenőrzés, fenntartás és kisebb finomhangolás."},
    {"statusz": 2, "nev": "Menedzselt rendszer", "min": 55, "szorzo": 1.00, "jutalek": 0.03,
     "dijmodell": "Havi menedzsmentdíj; jutalékkal kombinálható",
     "leiras": "A rendszer működik, de rendszeres operatív kezelést igényel: ismétlődő kampányfrissítések, "
               "szezonális promóciók, kreatívcserék, tervezhető optimalizálási feladatok."},
    {"statusz": 3, "nev": "Fejlesztés alatt álló rendszer", "min": 40, "szorzo": 1.15, "jutalek": 0.0,
     "dijmodell": "Havi díj fejlesztési felárral, fix díjas",
     "leiras": "Az alapok részben adottak, de több ponton optimalizálásra van szükség: kampányok, mérések, "
               "ajánlatok, kreatívok vagy landing oldalak fejlesztése."},
    {"statusz": 4, "nev": "Stratégiai fejlesztést igénylő rendszer", "min": None, "szorzo": 1.30, "jutalek": 0.0,
     "dijmodell": "Stratégiai fejlesztési díj, fix díjas",
     "leiras": "A növekedéshez már nem elég a meglévő rendszer kezelése: új ötletek, kampánytípusok, csatornák, "
               "az ajánlati vagy értékesítési logika újragondolása kell."},
    {"statusz": 5, "nev": "Kritikus helyreállítást igénylő rendszer", "min": None, "szorzo": None,
     "jutalek": 0.0, "dijmodell": "Nem vállaljuk",
     "leiras": "A jelenlegi marketingrendszer nem alkalmas stabil eredmények elérésére. Automatikusan "
               "kipontozódik, elutasítjuk."},
]

# A 4. státusz feltételei: legalább ennyi rendszerpontszám (alatta a rendszer
# kritikus), és a felsorolt csoportokban legalább ennyi csoportpontszám (0–100).
STATUS4_MIN_SYSTEM = 25
STATUS4_RULE = {"Értékek": 80, "Kapcsolat": 75}
MAX_STATUS4 = 1  # ennyi 4-es státuszú ügyfelet vállalunk; a többi várólistára kerül

# Kizáró feltételek: ha a kritérium pontja legfeljebb ennyi, az ügyfél 5.
# státuszú, és elutasítjuk (értékeinkkel nem összeegyeztethető, illetve nem
# felelős szemléletű cég).
KNOCKOUT = {"szegmens": 1, "felelosseg": 3}

# Jutalékos modell: akkor ajánljuk, ha a nyitottság pontja legalább ennyi, és a
# státuszhoz tartozó jutalék nagyobb 0-nál. Ilyenkor a fix díj ennyivel csökken.
COMMISSION = {"kuszob": 7, "fix_csokkentes": 0.25}

# Portfólió-célok az elfogadott ügyfelek arányában.
PORTFOLIO_TARGETS = {"top2_min": 0.70, "top2_max": 0.80, "s3_min": 0.20, "s3_max": 0.25}

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
    + [(c["kod"], _input_header(c), 30 if c["tipus"] == "lista" else 12,
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
        "missing_points": MISSING_POINTS,
        "statuses": STATUSES,
        "system_groups": SYSTEM_GROUPS,
        "status4_min_system": STATUS4_MIN_SYSTEM,
        "status4_rule": STATUS4_RULE,
        "max_status4": MAX_STATUS4,
        "knockout": KNOCKOUT,
        "commission": COMMISSION,
        "portfolio_targets": PORTFOLIO_TARGETS,
        "pricing": {k: v for k, (_, v) in PRICING.items()},
    }
