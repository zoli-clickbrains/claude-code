"""Kitalált mintaügyfelek a táblázat kipróbálásához és a tesztekhez."""

from .config import CRITERIA, SEGMENTS

EMBER, TUDATOS, HATAS, FELELOS, ALTALANOS, ELLENTETES = (name for name, _ in SEGMENTS)
MARKETING_CODES = [c["kod"] for c in CRITERIA if c["csoport"] == "Marketingezhetőség"]


def _client(nev, szegmens, felelosseg, marketing, uzlet, kapcsolat, dijazas):
    """`marketing`: a 17 marketingezhetőségi pont szóközzel elválasztva ("" = hiányzó adat)."""
    values = marketing.split(" ")
    assert len(values) == len(MARKETING_CODES), nev
    client = {"nev": nev, "szegmens": szegmens, "felelosseg": felelosseg}
    client |= {k: ("" if v == "-" else int(v)) for k, v in zip(MARKETING_CODES, values)}
    return client | uzlet | kapcsolat | dijazas


def _uzlet(iparagi, arbevetel, keret, erettseg, roas, novekedes):
    return dict(iparagi=iparagi, arbevetel=arbevetel, keret=keret, erettseg=erettseg, roas=roas,
                novekedes=novekedes)


def _kapcsolat(keses, hossz, egyuttmukodes, kezelhetoseg, ajanlas, elvarasok, emberi, szakmai):
    return dict(keses=keses, hossz=hossz, egyuttmukodes=egyuttmukodes, kezelhetoseg=kezelhetoseg,
                ajanlas=ajanlas, elvarasok=elvarasok, emberi=emberi, szakmai=szakmai)


def _dijazas(google, meta, egyeb_csatorna, orak, jelenlegi_dij, bevetel):
    return dict(google=google, meta=meta, egyeb_csatorna=egyeb_csatorna, orak=orak,
                jelenlegi_dij=jelenlegi_dij, bevetel=bevetel)


SAMPLE_CLIENTS = [
    _client("Minta Mentálhigiénés Központ", EMBER, 9, "9 7 7 6 8 7 7 7 9 8 7 6 7 8 10 6 8",
            _uzlet(8, 180, 900_000, 7, 4.5, 8),
            _kapcsolat(0, 26, 9, 8, 9, 9, 10, 9),
            _dijazas("igen", "igen", 0, 16, 260_000, 4_050_000)),
    _client("Minta Kézműves Sajtműhely", TUDATOS, 10, "5 6 7 7 6 9 6 6 8 8 9 8 9 9 5 7 7",
            _uzlet(8, 95, 450_000, 6, 5.2, 7),
            _kapcsolat(2, 18, 9, 9, 10, 8, 10, 10),
            _dijazas("igen", "igen", 0, 10, 180_000, 2_340_000)),
    _client("Minta Biogazdaság", TUDATOS, 10, "5 7 6 7 5 7 5 4 7 7 7 7 8 9 5 5 6",
            _uzlet(7, 60, 300_000, 5, 3.8, 6),
            _kapcsolat(5, 9, 8, 7, 8, 7, 9, 8),
            _dijazas("igen", "igen", 0, 8, 150_000, 1_140_000)),
    _client("Minta Közösségi Alapítvány", HATAS, 10, "6 6 4 5 8 7 7 7 6 6 6 7 7 8 8 6 5",
            _uzlet(6, 40, 200_000, 5, "", 6),
            _kapcsolat(8, 14, 8, 8, 9, 8, 9, 8),
            _dijazas("nem", "igen", 1, 8, 150_000, "")),
    _client("Minta Online Nyelviskola", EMBER, 8, "7 8 6 8 9 6 8 8 9 9 6 7 7 9 8 6 8",
            _uzlet(9, 120, 650_000, 7, 4.2, 7),
            _kapcsolat(3, 20, 8, 9, 8, 8, 8, 8),
            _dijazas("igen", "igen", 0, 12, 250_000, 2_730_000)),
    _client("Minta Fenntartható Divat Webshop", TUDATOS, 8, "4 7 7 6 6 7 6 6 8 7 8 7 10 9 4 7 8",
            _uzlet(9, 420, 2_200_000, 8, 4.8, 8),
            _kapcsolat(12, 7, 6, 5, 6, 5, 6, 8),
            _dijazas("igen", "igen", 1, 24, 420_000, 10_560_000)),
    _client("Minta Gyorsdivat Webáruház", ALTALANOS, 3, "5 9 6 7 7 4 9 7 8 7 5 6 8 9 3 8 8",
            _uzlet(10, 2500, 12_000_000, 8, 4.8, 9),
            _kapcsolat(10, 7, 6, 4, 5, 5, 5, 6),
            _dijazas("igen", "igen", 2, 40, 900_000, 57_600_000)),
    _client("Minta Ipari Gépgyártó Zrt.", ALTALANOS, 5, "4 4 9 3 2 7 2 2 6 5 7 6 4 5 7 5 6",
            _uzlet(4, 5200, 900_000, 4, "", 5),
            _kapcsolat(25, 9, 4, 5, 4, 5, 4, 7),
            _dijazas("igen", "nem", 1, 16, 280_000, "")),
    _client("Minta Ingatlaniroda", ELLENTETES, 2, "6 5 8 3 7 3 8 7 4 5 3 5 6 7 4 4 3",
            _uzlet(6, 90, 400_000, 3, 1.4, 4),
            _kapcsolat(40, 2, 3, 2, 2, 2, 3, 4),
            _dijazas("igen", "igen", 0, 10, 200_000, 560_000)),
    _client("Minta Új Érdeklődő Kft.", EMBER, "", "8 - - - - - - - - - - - - - 9 - -",
            _uzlet(7, "", 300_000, 5, "", ""),
            _kapcsolat("", 0, "", "", "", "", "", ""),
            _dijazas("igen", "nem", 0, 8, "", "")),
]
