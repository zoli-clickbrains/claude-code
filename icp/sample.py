"""Kitalált mintaügyfelek a táblázat kipróbálásához és a tesztekhez."""

from .config import CRITERIA

_OPTIONS = {c["kod"]: [name for name, _ in c["opciok"]] for c in CRITERIA if c["tipus"] == "lista"}
EMBER, TUDATOS, HATAS, FELELOS, ALTALANOS, ELLENTETES = _OPTIONS["szegmens"]
J_IGEN, J_NYITOTT, J_TALAN, J_NEM = _OPTIONS["jutalek_nyitottsag"]
T_JO, T_VEGYES, T_NINCS, T_ROSSZ = _OPTIONS["tapasztalat"]
MARKETING_CODES = [c["kod"] for c in CRITERIA if c["csoport"] == "Marketingezhetőség"]


def _client(nev, szegmens, felelosseg, marketing, uzlet, kapcsolat, dijazas):
    """`marketing`: a marketingezhetőségi pontok szóközzel elválasztva ("-" = hiányzó adat)."""
    values = marketing.split(" ")
    assert len(values) == len(MARKETING_CODES), nev
    client = {"nev": nev, "szegmens": szegmens, "felelosseg": felelosseg}
    client |= {k: ("" if v == "-" else int(v)) for k, v in zip(MARKETING_CODES, values)}
    return client | uzlet | kapcsolat | dijazas


def _uzlet(iparagi, arbevetel, keret, erettseg, roas, novekedes, jutalek_nyitottsag):
    return dict(iparagi=iparagi, arbevetel=arbevetel, keret=keret, erettseg=erettseg, roas=roas,
                novekedes=novekedes, jutalek_nyitottsag=jutalek_nyitottsag)


def _kapcsolat(keses, hossz, egyuttmukodes, kezelhetoseg, ajanlas, elvarasok, emberi, szakmai, tapasztalat):
    return dict(keses=keses, hossz=hossz, egyuttmukodes=egyuttmukodes, kezelhetoseg=kezelhetoseg,
                ajanlas=ajanlas, elvarasok=elvarasok, emberi=emberi, szakmai=szakmai, tapasztalat=tapasztalat)


def _dijazas(google, meta, egyeb_csatorna, orak, jelenlegi_dij, bevetel):
    return dict(google=google, meta=meta, egyeb_csatorna=egyeb_csatorna, orak=orak,
                jelenlegi_dij=jelenlegi_dij, bevetel=bevetel)


SAMPLE_CLIENTS = [
    _client("Minta Mentálhigiénés Központ", EMBER, 9, "9 7 7 6 8 7 7 7 9 8 7 7 8 10 6",
            _uzlet(8, 180, 900_000, 8, 4.5, 8, J_NYITOTT),
            _kapcsolat(0, 26, 9, 8, 9, 9, 10, 9, T_JO),
            _dijazas("igen", "igen", 0, 12, 260_000, 4_050_000)),
    _client("Minta Kézműves Sajtműhely", TUDATOS, 10, "5 6 7 7 6 9 6 6 8 8 8 9 9 5 7",
            _uzlet(8, 95, 450_000, 7, 5.2, 7, J_IGEN),
            _kapcsolat(2, 30, 9, 9, 10, 9, 10, 10, T_VEGYES),
            _dijazas("igen", "igen", 0, 8, 180_000, 2_340_000)),
    _client("Minta Online Nyelviskola", EMBER, 8, "7 8 6 8 9 7 8 8 9 9 7 7 9 8 6",
            _uzlet(9, 120, 650_000, 8, 4.2, 7, J_TALAN),
            _kapcsolat(3, 20, 8, 9, 8, 8, 8, 8, T_JO),
            _dijazas("igen", "igen", 0, 10, 250_000, 2_730_000)),
    _client("Minta Fejlesztő Központ", EMBER, 9, "8 7 7 7 8 8 7 7 7 8 8 8 8 9 7",
            _uzlet(8, 70, 350_000, 7, 3.6, 8, J_NYITOTT),
            _kapcsolat(0, 15, 9, 8, 9, 8, 9, 9, T_NINCS),
            _dijazas("igen", "igen", 0, 8, 170_000, 1_260_000)),
    _client("Minta Biogazdaság", TUDATOS, 10, "5 7 6 7 5 7 5 4 7 7 7 8 9 5 5",
            _uzlet(7, 60, 300_000, 6, 3.8, 6, J_NYITOTT),
            _kapcsolat(5, 9, 8, 7, 8, 8, 9, 8, T_VEGYES),
            _dijazas("igen", "igen", 0, 8, 150_000, 1_140_000)),
    _client("Minta Fenntartható Divat Webshop", TUDATOS, 8, "4 7 7 6 6 8 6 6 8 7 7 10 9 4 7",
            _uzlet(9, 420, 2_200_000, 8, 4.8, 8, J_NEM),
            _kapcsolat(6, 7, 7, 6, 7, 6, 7, 8, T_VEGYES),
            _dijazas("igen", "igen", 1, 22, 420_000, 10_560_000)),
    _client("Minta Jógastúdió", EMBER, 8, "6 6 5 7 8 6 8 8 7 8 7 8 8 6 5",
            _uzlet(7, 45, 200_000, 6, 3.2, 6, J_TALAN),
            _kapcsolat(4, 12, 8, 7, 8, 7, 8, 7, T_NINCS),
            _dijazas("nem", "igen", 0, 6, 150_000, 640_000)),
    _client("Minta Közösségi Alapítvány", HATAS, 10, "6 6 4 5 8 7 7 7 6 6 7 7 8 8 6",
            _uzlet(6, 40, 200_000, 5, "", 6, J_NEM),
            _kapcsolat(8, 14, 8, 8, 9, 8, 9, 8, T_NINCS),
            _dijazas("nem", "igen", 1, 8, 150_000, "")),
    _client("Minta Gyorsdivat Webáruház", ALTALANOS, 3, "5 9 6 7 7 5 9 7 8 7 6 8 9 3 8",
            _uzlet(10, 2500, 12_000_000, 8, 4.8, 9, J_NYITOTT),
            _kapcsolat(10, 7, 6, 4, 5, 5, 5, 6, T_ROSSZ),
            _dijazas("igen", "igen", 2, 40, 900_000, 57_600_000)),
    _client("Minta Kerámiaműhely", TUDATOS, 10, "3 4 5 3 4 7 3 4 4 5 5 9 6 2 5",
            _uzlet(4, 20, 100_000, 2, "", 5, J_NEM),
            _kapcsolat(0, 3, 9, 8, 9, 8, 10, 9, T_NINCS),
            _dijazas("nem", "igen", 0, 10, "", "")),
    _client("Minta Környezetvédelmi Egyesület", HATAS, 10, "4 4 3 3 5 6 4 4 3 4 5 5 7 4 5",
            _uzlet(4, 15, 80_000, 3, "", 4, J_NEM),
            _kapcsolat(0, 0, 8, 8, 8, 8, 10, 9, T_NINCS),
            _dijazas("nem", "igen", 0, 8, "", "")),
    _client("Minta Gyógytorna Rendelő", EMBER, 8, "7 5 5 5 7 5 6 6 5 6 6 5 7 7 4",
            _uzlet(6, 40, 150_000, 4, 2.5, 6, J_TALAN),
            _kapcsolat(5, 10, 7, 7, 7, 7, 8, 8, T_NINCS),
            _dijazas("igen", "igen", 0, 8, 150_000, 375_000)),
    _client("Minta Ingatlaniroda", ELLENTETES, 2, "6 5 8 3 7 3 8 7 4 5 5 6 7 4 4",
            _uzlet(6, 90, 400_000, 3, 1.4, 4, J_NYITOTT),
            _kapcsolat(40, 2, 3, 2, 2, 2, 3, 4, T_ROSSZ),
            _dijazas("igen", "igen", 0, 10, 200_000, 560_000)),
    _client("Minta Új Érdeklődő Kft.", ALTALANOS, "", "8 - - - - - - - - - - - - 9 -",
            _uzlet(7, "", 300_000, 5, "", "", ""),
            _kapcsolat("", 0, "", "", "", "", "", "", ""),
            _dijazas("igen", "nem", 0, 8, "", "")),
]
