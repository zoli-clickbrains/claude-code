# ICP ügyfélminősítő

B2B online marketing ügyfelek pontozása (0–100), A–D szintbe sorolása, rangsorolása és havi díjjavaslat.

Két, egymással egyező megvalósítás:

- **Táblázat** (`build/ICP_ugyfelminosito.xlsx`): élő képletekkel, Excelben és Google Táblázatokban is működik. Most ezt használjuk.
- **Python-motor** (`icp/engine.py`): ugyanaz a logika kódban, a belső rendszerbe (ad-automation) való későbbi beépítéshez.

A tesztek LibreOffice-szal újraszámolják a táblázatot, és ellenőrzik, hogy minden pont, szint és díj egyezik a motor eredményével.

## Használat

```bash
pip install -r requirements.txt

python -m icp tablazat build/ICP_ugyfelminosito.xlsx          # üres sablon
python -m icp tablazat build/ICP_ugyfelminosito_minta.xlsx --minta   # kitalált mintaügyfelekkel
python -m icp pontoz build/ICP_ugyfelminosito_minta.xlsx --csv build/eredmeny.csv

python -m pytest
```

## A táblázat lapjai

| Lap | Tartalom |
|---|---|
| Útmutató | Rövid leírás |
| Ügyfelek | Bemenet: ügyfelenként egy sor, legördülő listákkal és ellenőrzött 1–5 skálákkal |
| Rangsor | Pontszám szerinti sorrend szinttel, díjjavaslattal és a jelenlegi díjtól való eltéréssel |
| Eredmény | Kritériumonkénti pontok és a díjszámítás részletei |
| Beállítások | Súlyok, sávhatárok, iparági pontok, szintek, díjparaméterek (sárga cellák) |

## Pontozás

11 kritérium, mindegyik 1–5 pontot kap. A kritériumok, a súlyuk és hogy honnan kapnak pontot:

| Csoport | Kritérium | Súly | Pontozás |
|---|---|---|---|
| Illeszkedés | Iparági illeszkedés | 10 | lista |
| | Éves árbevétel | 10 | sávok (M Ft) |
| | Havi hirdetési keret | 15 | sávok (Ft) |
| | Marketing-érettség | 10 | 1–5 |
| Teljesítmény | Jelenlegi ROAS | 10 | sávok (a dashboardról) |
| | Növekedési potenciál | 10 | 1–5 |
| Kapcsolat | Fizetési fegyelem | 10 | sávok (késés napban) |
| | Együttműködés hossza | 5 | sávok (hónap) |
| | Döntéshozó, együttműködés | 10 | 1–5 |
| | Kezelhetőség | 5 | 1–5 |
| | Ajánlási, referenciaérték | 5 | 1–5 |

A pontszám így számol: `ICP pontszám = Σ súly × (pont − 1) / 4`. Ha egy adat hiányzik, 2 pontot kap, hogy a hiányos adatlap ne kerüljön előrébb a kitöltöttnél.

A szintek határai:

- **A:** legalább 80 pont (ideális ügyfél)
- **B:** legalább 65 pont
- **C:** legalább 50 pont
- **D:** 50 pont alatt

## Díjjavaslat

- **Munkadíj:** Google-alapdíj + Meta-alapdíj + egyéb csatornák díja + becsült havi óraszám × óradíj.
- **Javasolt havi díj:** a három érték közül a legnagyobb, felfelé kerekítve 10 000 Ft-ra:
  - munkadíj × szintszorzó (A 1,00 · B 1,05 · C 1,15 · D 1,30),
  - a hirdetési keret 8%-a,
  - minimum 150 000 Ft.
- **Jutalék:** később, szintenként megadott %-ban számolható a mért bevételből. A beállításokban most 0%.

Minden alapérték az `icp/config.py` fájlban van, a táblázatban pedig a Beállítások lapon.
