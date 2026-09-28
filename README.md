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
| Útmutató | Rövid leírás, az ideális ügyfél, a státuszok és a díjazás magyarázata |
| Ügyfelek | Bemenet: ügyfelenként egy sor, csoportszínű fejlécekkel, legördülő listákkal és ellenőrzött 1–10 skálákkal |
| Rangsor | ICP pontszám szerinti sorrend státusszal, döntéssel, csoportpontszámokkal, fix díjjal és jutalékos ajánlattal |
| Portfólió | Hány elfogadott ügyfél van az egyes státuszokban, és teljesülnek-e a célarányok |
| Eredmény | Szempontonkénti pontok és a díjszámítás részletei |
| Beállítások | Súlyok, sávhatárok, választható értékek, státuszhatárok, kizáró feltételek, díjparaméterek (sárga cellák) |

## Kiket tekintünk ideális ügyfélnek?

Felelősségteljes, fenntartható, zöld szemléletű cégeket, amelyek elsősorban szükségleteket elégítenek ki, nem igényeket:

- **Emberközpontú szolgáltatók:** egészség, oktatás, mentális jóllét, fejlesztés.
- **Tudatos termelők és alkotók:** fenntartható termékek, kézműves márkák, gazdaságok.
- **Hatásközpontú szervezetek:** közösségi, környezeti és társadalmi projektek.

**Kizáró feltételek:** az „Értékeinkkel nem összeegyeztethető” szegmens és a legfeljebb 3 pontos felelősség. Az ilyen ügyfél automatikusan 5. státuszú, és elutasítjuk. Hiányzó adat nem zár ki.

## Pontozás

33 szempont, mindegyik 1–10 pontot kap. A 10 mindig a számunkra kedvező érték (pl. „Vásárlószerzés költsége”: 10 = olcsó, könnyű). A súlyozás ideiglenes váz; a hirdetési keret súlya szándékosan alacsony (2).

| Csoport | Szempontok | Súly |
|---|---|---|
| Értékek | Értékalapú szegmens (lista), felelősség és fenntarthatóság | 20 |
| Marketingezhetőség | 15 szempont: sürgősség, piacméret, árazási potenciál, vásárlószerzés költsége, leszállítás költsége, egyediség és egyedi tulajdonság, a piac gyorsasága, befektetési igény, upsell-potenciál, örökzöld potenciál, látványos demonstrálhatóság, dizájn, egyszerű használat, problémamegoldás, újdonság | 30 |
| Üzleti illeszkedés | Iparági illeszkedés, éves árbevétel, havi hirdetési keret, marketing-érettség, jelenlegi ROAS, növekedési potenciál, nyitottság a jutalékos modellre (lista) | 20 |
| Kapcsolat | Fizetési fegyelem, együttműködés hossza, döntéshozó, kezelhetőség, ajánlási érték, reális elvárások, emberi oldal, szakmai színvonal, korábbi ügynökségi tapasztalat (lista) | 30 |

Az eredeti 17 marketingezhetőségi szempontból kettő-kettő összevonva: „Egyediség a piacon” + „Egyedi tulajdonság”, illetve „Upsell-potenciál” + „Van hozzá upsell”.

Két összesített pontszám készül (0–100):

- **ICP pontszám:** mind a 33 szempont, `100 × Σ súly × (pont − 1) / (9 × Σ súly)`. Ez adja a rangsort.
- **Rendszerpontszám:** csak a marketingezhetőség és az üzleti illeszkedés. Ez dönti el a státuszt, mert a státusz a marketingrendszer állapotát írja le.

Ha egy adat hiányzik, 3 pontot kap, hogy a hiányos adatlap ne kerüljön előrébb a kitöltöttnél.

## Státuszok és díjazás

| Státusz | Feltétel | Díjszorzó | Jutalék (jutalékos modellben) |
|---|---|---|---|
| 1. Stabil rendszer | rendszerpontszám ≥ 70 | 0,90 | 5% |
| 2. Menedzselt rendszer | rendszerpontszám ≥ 55 | 1,00 | 3% |
| 3. Fejlesztés alatt álló rendszer | rendszerpontszám ≥ 40 | 1,15 | – |
| 4. Stratégiai fejlesztést igénylő rendszer | 40 alatt, de ≥ 25, és Értékek ≥ 80, Kapcsolat ≥ 75 | 1,30 | – |
| 5. Kritikus helyreállítást igénylő rendszer | minden más, vagy kizáró feltétel | nincs ajánlat, elutasítjuk | – |

- **4-es hely:** legfeljebb 1 ügyfél lehet 4. státuszú (a legmagasabb ICP pontszámú); a többi várólistára kerül.
- **Célarány** (Portfólió lap, az elfogadott ügyfelekre): 70–80% legyen 1. vagy 2. státuszú, 20–25% pedig 3. státuszú.
- **Munkadíj:** Google-alapdíj + Meta-alapdíj + egyéb csatornák díja + becsült havi óraszám × óradíj.
- **Javasolt fix havi díj:** a három érték közül a legnagyobb, felfelé kerekítve 10 000 Ft-ra:
  - munkadíj × státuszszorzó,
  - a hirdetési keret 8%-a,
  - minimum 150 000 Ft.
- **Jutalékos modell:** akkor ajánljuk, ha a státuszhoz jutalék tartozik (1–2.), és az ügyfél nyitott rá (a nyitottság legalább 7 pont). Ilyenkor a fix díj 25%-kal csökken, és a mért bevétel státusz szerinti %-a jutalékként jár.

Minden alapérték az `icp/config.py` fájlban van, a táblázatban pedig a Beállítások lapon.
