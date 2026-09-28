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
| Útmutató | Rövid leírás, az ideális ügyfél meghatározása |
| Ügyfelek | Bemenet: ügyfelenként egy sor, csoportszínű fejlécekkel, legördülő listákkal és ellenőrzött 1–10 skálákkal |
| Rangsor | Pontszám szerinti sorrend szinttel, értékkapuval, csoportpontszámokkal, díjjavaslattal és a jelenlegi díjtól való eltéréssel |
| Eredmény | Szempontonkénti pontok és a díjszámítás részletei |
| Beállítások | Súlyok, sávhatárok, szegmensek, értékkapu, szintek, díjparaméterek (sárga cellák) |

## Kiket tekintünk ideális ügyfélnek?

Felelősségteljes, fenntartható, zöld szemléletű cégeket, amelyek elsősorban szükségleteket elégítenek ki, nem igényeket:

- **Emberközpontú szolgáltatók:** egészség, oktatás, mentális jóllét, fejlesztés.
- **Tudatos termelők és alkotók:** fenntartható termékek, kézműves márkák, gazdaságok.
- **Hatásközpontú szervezetek:** közösségi, környezeti és társadalmi projektek.

**Értékkapu:** A szintet csak az kaphat, akinek az „Értékalapú szegmens” pontja legalább 8, a „Felelősség, fenntarthatóság” pontja pedig legalább 7. Aki nem jut át, legfeljebb B szintű lehet, bármilyen magas a pontszáma.

## Pontozás

33 szempont, mindegyik 1–10 pontot kap. A 10 mindig a számunkra kedvező érték (pl. „Vásárlószerzés költsége”: 10 = olcsó, könnyű).

| Csoport | Szempontok | Súly |
|---|---|---|
| Értékek | Értékalapú szegmens (lista), Felelősség, fenntarthatóság | 20 |
| Marketingezhetőség | A 17 szempont: sürgősség, piacméret, árazási potenciál, vásárlószerzés költsége, leszállítás költsége, egyediség a piacon, a piac gyorsasága, befektetési igény, upsell-potenciál, örökzöld potenciál, egyedi tulajdonság, látványos demonstrálhatóság, dizájn, egyszerű használat, problémamegoldás, újdonság, upsell-ajánlat | 35 |
| Üzleti illeszkedés | Iparági illeszkedés, éves árbevétel, havi hirdetési keret, marketing-érettség, jelenlegi ROAS, növekedési potenciál | 22 |
| Kapcsolat | Fizetési fegyelem, együttműködés hossza, döntéshozó, kezelhetőség, ajánlási érték, reális elvárások, emberi oldal, szakmai színvonal | 23 |

A szempontonkénti súlyok és leírások a Beállítások lapon és az `icp/config.py` fájlban vannak.

A pontszám így számol: `ICP pontszám = 100 × Σ súly × (pont − 1) / (9 × Σ súly)`. A négy csoport külön is kap 0–100-as pontszámot. Ha egy adat hiányzik, 3 pontot kap, hogy a hiányos adatlap ne kerüljön előrébb a kitöltöttnél.

Az árbevétel- és a hirdetésikeret-sávok kkv-khoz és szervezetekhez vannak kalibrálva (pl. a havi 1 millió Ft-os keret már 9 pont).

A szintek határai:

- **A:** legalább 80 pont és átjut az értékkapun (ideális ügyfél)
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
