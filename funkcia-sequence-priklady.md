# Funkcia SEQUENCE v Microsoft Excel – praktické príklady

Tento dokument slúži ako sprievodca k cvičnému zošitu `funkcia-sequence-microsoft-excel-prakticke-cvicenia.xlsx`. Obsahuje syntax, praktické scenáre, vzorce pripravené na rýchle kopírovanie a príklady pre dynamické polia, dátumy, identifikátory, dvojrozmerné matice, operátor `#` a chybu `#SPILL!`.

> **Poznámka:** príklady používajú anglické názvy funkcií a oddeľovač argumentov `;`, ktorý je bežný pri slovenských regionálnych nastaveniach Excelu. Pri desatinných číslach používame čiarku, napr. `0,5`.

## Obsah

1. [Čo robí funkcia SEQUENCE](#co-robi-sequence)
2. [Syntax funkcie SEQUENCE](#syntax)
3. [Rýchle kopírovanie – najpoužívanejšie vzorce](#rychle-kopirovanie)
4. [Základná zvislá postupnosť](#zvisla-postupnost)
5. [Horizontálna postupnosť](#horizontalna-postupnost)
6. [Vlastný začiatok](#vlastny-zaciatok)
7. [Vlastný krok](#vlastny-krok)
8. [Klesajúca postupnosť](#klesajuca-postupnost)
9. [Desatinný krok](#desatinny-krok)
10. [Dvojrozmerné pole](#dvojrozmerne-pole)
11. [Dynamické parametre z buniek](#dynamicke-parametre)
12. [Dátumy a časové intervaly](#datumy-a-cas)
13. [Dynamické pole a operátor #](#operator-spill)
14. [Násobilka pomocou SEQUENCE](#nasobilka)
15. [Chyba #SPILL!](#spill-error)
16. [SEQUENCE a Excel Table](#excel-table)
17. [Praktické scenáre podľa hárkov](#scenare)
18. [Odporúčané cvičenia](#cvicenia)
19. [Najčastejšie chyby](#chyby)
20. [Oficiálna dokumentácia](#dokumentacia)

<a id="co-robi-sequence"></a>
## 1. Čo robí funkcia SEQUENCE

Funkcia `SEQUENCE` generuje postupnosť čísel ako dynamické pole. Namiesto ručného vypĺňania buniek môžeme jedným vzorcom vytvoriť desiatky, stovky alebo celé dvojrozmerné matice hodnôt.

Príklad:

```excel
=SEQUENCE(10)
```

Výsledok:

```text
1
2
3
4
5
6
7
8
9
10
```

Vzorec zadávame iba do jednej bunky. Excel výsledok automaticky rozleje do ďalších buniek.

<a id="syntax"></a>
## 2. Syntax funkcie SEQUENCE

```excel
=SEQUENCE(rows;[columns];[start];[step])
```

| Argument | Význam | Príklad |
|---|---|---|
| `rows` | počet riadkov výsledného poľa | `120` |
| `columns` | počet stĺpcov výsledného poľa | `12` |
| `start` | prvá hodnota postupnosti | `1001` |
| `step` | rozdiel medzi nasledujúcimi hodnotami | `5` |

Ak voliteľné argumenty neuvedieme, Excel štandardne používa hodnotu `1`.

<a id="rychle-kopirovanie"></a>
## 3. Rýchle kopírovanie – najpoužívanejšie vzorce

### Základné sekvencie

```excel
=SEQUENCE(120)
```

```excel
=SEQUENCE(1;12)
```

```excel
=SEQUENCE(10;12)
```

### Vlastný začiatok

```excel
=SEQUENCE(120;1;1001)
```

```excel
=SEQUENCE(150;1;50001)
```

```excel
=SEQUENCE(120;1;260001)
```

### Vlastný krok

```excel
=SEQUENCE(120;1;1000;10)
```

```excel
=SEQUENCE(150;1;90000;5)
```

```excel
=SEQUENCE(120;1;4000;10)
```

### Klesajúca postupnosť

```excel
=SEQUENCE(120;1;120;-1)
```

```excel
=SEQUENCE(120;1;100;-0,5)
```

### Desatinné hodnoty

```excel
=SEQUENCE(130;1;0;0,5)
```

```excel
=SEQUENCE(120;1;10;2,5)
```

### Dátumy

```excel
=SEQUENCE(120;1;DATE(2026;9;1);1)
```

```excel
=SEQUENCE(18;1;DATE(2026;9;1);7)
```

### Časové intervaly

```excel
=SEQUENCE(48;1;0;1/48)
```

```excel
=SEQUENCE(96;1;0;1/96)
```

### Názvy mesiacov

```excel
=TEXT(DATE(YEAR(TODAY());SEQUENCE(1;12);1);"mmm")
```

### Dynamické parametre z buniek

```excel
=SEQUENCE(M5;M6;M7;M8)
```

### Práca s celým spill rozsahom

```excel
=SUM(M6#)
```

```excel
=ROWS(M6#)
```

```excel
=COLUMNS(M6#)
```

```excel
=COUNTA(M6#)
```

```excel
=AVERAGE(M6#)
```

```excel
=MAX(M6#)
```

```excel
=MIN(M6#)
```

```excel
=MAX(M6#)-MIN(M6#)
```

### Násobilka

```excel
=SEQUENCE(10)*TRANSPOSE(SEQUENCE(10))
```

```excel
=SEQUENCE(12)*TRANSPOSE(SEQUENCE(12))
```

<a id="zvisla-postupnost"></a>
## 4. Základná zvislá postupnosť

Najjednoduchší príklad vytvorí čísla od 1 po 120:

```excel
=SEQUENCE(120)
```

Praktické použitie v hárku `01_Zamestnanci`: poradové čísla zamestnancov.

Praktické použitie v hárku `03_ProjektovyPlan`: `TaskID` pre 120 projektových úloh.

Praktické použitie v hárku `06_Skolenia`: poradové čísla účastníkov.

<a id="horizontalna-postupnost"></a>
## 5. Horizontálna postupnosť

Ak chceme hodnoty do stĺpcov namiesto riadkov, nastavíme jeden riadok a požadovaný počet stĺpcov:

```excel
=SEQUENCE(1;12)
```

Výsledok:

```text
1  2  3  4  5  6  7  8  9  10  11  12
```

Použitie:

- mesiace 1 až 12,
- projektové sprinty,
- účtovné obdobia,
- skladové regály,
- kapacitné obdobia.

<a id="vlastny-zaciatok"></a>
## 6. Vlastný začiatok

### EmployeeID od 1001

```excel
=SEQUENCE(120;1;1001)
```

### OrderID od 50001

```excel
=SEQUENCE(150;1;50001)
```

### InvoiceID od 260001

```excel
=SEQUENCE(120;1;260001)
```

### Číslo dokladu od 2600001

```excel
=SEQUENCE(120;1;2600001)
```

Tretí argument určuje prvú hodnotu celej sekvencie.

<a id="vlastny-krok"></a>
## 7. Vlastný krok

### Interný kód po desiatkach

```excel
=SEQUENCE(120;1;1000;10)
```

Výsledok začína:

```text
1000
1010
1020
1030
...
```

### Logistický kód s krokom 5

```excel
=SEQUENCE(150;1;90000;5)
```

### GL kódy od 4000 s krokom 10

```excel
=SEQUENCE(120;1;4000;10)
```

### Paletové čísla po 100

```excel
=SEQUENCE(130;1;100;100)
```

<a id="klesajuca-postupnost"></a>
## 8. Klesajúca postupnosť

Záporným krokom vytvoríme zostupnú postupnosť:

```excel
=SEQUENCE(120;1;120;-1)
```

Výsledok začína:

```text
120
119
118
117
...
```

Použitie:

- reverzné poradie,
- auditný index,
- klesajúca kapacita,
- kontrolné poradie,
- odpočítavanie.

<a id="desatinny-krok"></a>
## 9. Desatinný krok

### Interval po 0,5

```excel
=SEQUENCE(130;1;0;0,5)
```

### Cenové úrovne od 10 s krokom 2,5

```excel
=SEQUENCE(120;1;10;2,5)
```

### Klesajúce ceny po 0,5

```excel
=SEQUENCE(120;1;100;-0,5)
```

Pri slovenskom regionálnom nastavení používame desatinnú čiarku `,` a argumenty oddeľujeme bodkočiarkou `;`.

<a id="dvojrozmerne-pole"></a>
## 10. Dvojrozmerné pole

Funkcia `SEQUENCE` nemusí vytvárať iba jeden stĺpec alebo jeden riadok.

### Pole 10 × 12

```excel
=SEQUENCE(10;12)
```

### Pole 10 × 8 so začiatkom 100

```excel
=SEQUENCE(10;8;100)
```

### Pole 10 × 8 so začiatkom 100 a krokom 5

```excel
=SEQUENCE(10;8;100;5)
```

Použitie:

- matica sprintov a úloh,
- skladové pozície,
- účtovná matica stredisko × obdobie,
- kapacitný plán,
- pomocné indexy pre analytické modely.

<a id="dynamicke-parametre"></a>
## 11. Dynamické parametre z buniek

Najpraktickejšie použitie vznikne vtedy, keď argumenty nezadáme napevno, ale načítame ich z buniek.

V hárku `14_DynamickeParametre` sú ovládacie hodnoty v bunkách:

| Bunka | Parameter | Príklad |
|---|---|---:|
| `M5` | Rows | 10 |
| `M6` | Columns | 1 |
| `M7` | Start | 100 |
| `M8` | Step | 5 |

Vzorec:

```excel
=SEQUENCE(M5;M6;M7;M8)
```

Zmenou jednej vstupnej bunky okamžite zmeníme veľkosť alebo hodnoty celého dynamického poľa.

Ďalšie varianty:

```excel
=SEQUENCE(M5)
```

```excel
=SEQUENCE(M5;M6)
```

```excel
=SEQUENCE(M5;M6;M7)
```

```excel
=SEQUENCE(M5;M6;M7;M8)
```

<a id="datumy-a-cas"></a>
## 12. Dátumy a časové intervaly

Excel ukladá dátumy a časy ako číselné hodnoty, preto môžeme `SEQUENCE` použiť aj na generovanie kalendárov a harmonogramov.

### 120 po sebe idúcich dní

```excel
=SEQUENCE(120;1;DATE(2026;9;1);1)
```

### 18 týždňov

```excel
=SEQUENCE(18;1;DATE(2026;9;1);7)
```

### 12 mesiacov ako názvy

```excel
=TEXT(DATE(YEAR(TODAY());SEQUENCE(1;12);1);"mmm")
```

### 30-minútové intervaly

```excel
=SEQUENCE(48;1;0;1/48)
```

Po vytvorení nastavíme bunkám formát času, napríklad `hh:mm`.

### 15-minútové intervaly

```excel
=SEQUENCE(96;1;0;1/96)
```

Použitie:

- rezervačný kalendár,
- harmonogram školení,
- projektové termíny,
- plán inkasa,
- pracovné zmeny,
- časové sloty.

<a id="operator-spill"></a>
## 13. Dynamické pole a operátor #

Ak zadáme do bunky `M6`:

```excel
=SEQUENCE(120)
```

celé dynamické pole môžeme následne adresovať ako:

```excel
M6#
```

### Súčet celého dynamického poľa

```excel
=SUM(M6#)
```

### Počet riadkov

```excel
=ROWS(M6#)
```

### Počet stĺpcov

```excel
=COLUMNS(M6#)
```

### Počet neprázdnych hodnôt

```excel
=COUNTA(M6#)
```

### Priemer

```excel
=AVERAGE(M6#)
```

### Maximum

```excel
=MAX(M6#)
```

### Minimum

```excel
=MIN(M6#)
```

### Rozpätie

```excel
=MAX(M6#)-MIN(M6#)
```

Hlavná výhoda operátora `#` je, že odkaz automaticky reaguje na zmenu veľkosti dynamického poľa.

<a id="nasobilka"></a>
## 14. Násobilka pomocou SEQUENCE

### Násobilka 10 × 10

```excel
=SEQUENCE(10)*TRANSPOSE(SEQUENCE(10))
```

### Násobilka 12 × 12

```excel
=SEQUENCE(12)*TRANSPOSE(SEQUENCE(12))
```

Takto kombinujeme zvislé dynamické pole s jeho transponovanou horizontálnou verziou.

<a id="spill-error"></a>
## 15. Chyba #SPILL!

Ak sa dynamické pole nemôže rozliať do potrebného rozsahu, Excel zobrazí chybu `#SPILL!`.

Typický test:

```excel
=SEQUENCE(120)
```

Potom vložíme hodnotu do bunky, do ktorej sa má výsledok rozliať. Excel nemôže prepísať existujúci obsah a zobrazí chybu.

Postup riešenia:

1. označíme bunku s chybou `#SPILL!`,
2. zistíme plánovaný spill rozsah,
3. nájdeme blokujúcu bunku,
4. obsah blokujúcej bunky odstránime alebo premiestnime,
5. overíme, že sa pole správne rozlialo.

Na precvičenie použijeme hárok `11_Spill_Error`.

<a id="excel-table"></a>
## 16. SEQUENCE a Excel Table

Dynamické polia majú obmedzenia pri použití priamo vo vypočítavaných stĺpcoch Excel Table. Preto testujeme dva scenáre:

### Mimo Excel Table

```excel
=SEQUENCE(120)
```

### Horizontálne pole vedľa tabuľky

```excel
=SEQUENCE(1;8)
```

### Dvojrozmerné pole vedľa tabuľky

```excel
=SEQUENCE(10;8)
```

Na precvičenie použijeme hárok `12_Excel_Tabulka`, kde porovnáme správanie dynamického poľa vo vnútri tabuľky a mimo nej.

<a id="scenare"></a>
## 17. Praktické scenáre podľa hárkov

| Hárok | Scenár | Hlavné precvičované oblasti |
|---|---|---|
| `01_Zamestnanci` | HR databáza | poradové čísla, EmployeeID, vlastný začiatok, krok, klesanie, `#` |
| `02_Objednavky` | objednávky e-shopu | OrderID, logistické kódy, horizontálna SEQUENCE, 2D pole |
| `03_ProjektovyPlan` | projektové riadenie | TaskID, WBS, dátumy, týždne, sprinty |
| `04_Faktury` | fakturácia | InvoiceID, variabilný symbol, dávkové kódy, dátumy, `#SPILL!` |
| `05_Sklad` | sklad a inventúra | SKU, palety, regály, 2D pozície, desatinný krok |
| `06_Skolenia` | vzdelávanie | účastníci, certifikáty, moduly, termíny |
| `07_Kalendar` | kalendár | dni, týždne, mesiace, 15- a 30-minútové intervaly |
| `08_CenoveIntervaly` | cenotvorba | desatinné kroky, cenové úrovne, klesajúce ceny |
| `09_GL_Kody` | účtovníctvo | GL kódy, strediská, doklady, obdobia |
| `10_KapacitnyPlan` | plánovanie kapacít | 2D matica, rows, columns, dynamické parametre |
| `11_Spill_Error` | diagnostika | `#SPILL!`, blokujúce bunky, zmena veľkosti poľa |
| `12_Excel_Tabulka` | Excel Table | dynamické pole v tabuľke a mimo tabuľky |
| `13_Nasobilka_2D` | matice | násobilka, 2D SEQUENCE, rozmery poľa |
| `14_DynamickeParametre` | parametrický model | `rows`, `columns`, `start`, `step` z buniek |

<a id="cvicenia"></a>
## 18. Odporúčané cvičenia

### Cvičenie 1 – 120 zamestnancov

Vytvorte poradové čísla:

```excel
=SEQUENCE(120)
```

Potom EmployeeID od 1001:

```excel
=SEQUENCE(120;1;1001)
```

A interný kód po desiatkach:

```excel
=SEQUENCE(120;1;1000;10)
```

### Cvičenie 2 – 150 objednávok

Vygenerujte OrderID od 50001:

```excel
=SEQUENCE(150;1;50001)
```

Logistický kód od 90000 s krokom 5:

```excel
=SEQUENCE(150;1;90000;5)
```

### Cvičenie 3 – projektový kalendár

Vygenerujte 120 projektových dní od 1. septembra 2026:

```excel
=SEQUENCE(120;1;DATE(2026;9;1);1)
```

Vytvorte týždenné míľniky:

```excel
=SEQUENCE(18;1;DATE(2026;9;1);7)
```

### Cvičenie 4 – faktúry

Vygenerujte interné čísla faktúr:

```excel
=SEQUENCE(120;1;260001)
```

Vygenerujte variabilné symboly:

```excel
=SEQUENCE(120;1;2026090001)
```

### Cvičenie 5 – sklad

Vytvorte SKU kódy:

```excel
=SEQUENCE(130;1;7000;5)
```

Vytvorte paletové čísla:

```excel
=SEQUENCE(130;1;100;100)
```

Vytvorte hmotnostné intervaly po 0,5:

```excel
=SEQUENCE(130;1;0;0,5)
```

### Cvičenie 6 – školenia

Vytvorte ParticipantID:

```excel
=SEQUENCE(120;1;3001)
```

Vytvorte čísla certifikátov po 10:

```excel
=SEQUENCE(120;1;10000;10)
```

### Cvičenie 7 – kalendár a čas

Vytvorte 48 polhodinových časových slotov:

```excel
=SEQUENCE(48;1;0;1/48)
```

Vytvorte 96 pätnásťminútových slotov:

```excel
=SEQUENCE(96;1;0;1/96)
```

### Cvičenie 8 – cenové intervaly

Vytvorte hodnoty od 0 po 0,5:

```excel
=SEQUENCE(120;1;0;0,5)
```

Vytvorte cenové úrovne od 10 s krokom 2,5:

```excel
=SEQUENCE(120;1;10;2,5)
```

### Cvičenie 9 – GL kódy

Vytvorte hlavné účtovné kódy:

```excel
=SEQUENCE(120;1;4000;10)
```

Vytvorte kódy stredísk:

```excel
=SEQUENCE(120;1;100;5)
```

### Cvičenie 10 – dynamická matica

Do vstupných buniek nastavte počet riadkov, počet stĺpcov, začiatok a krok a použite:

```excel
=SEQUENCE(M5;M6;M7;M8)
```

Následne meňte jednotlivé parametre a sledujte zmenu rozmerov a hodnôt poľa.

<a id="chyby"></a>
## 19. Najčastejšie chyby

### Nedostatok voľného priestoru

Ak sa výsledok nemá kam rozliať, vznikne `#SPILL!`.

### Zámena počtu prvkov a poslednej hodnoty

```excel
=SEQUENCE(10;1;100;5)
```

nevytvorí hodnoty od 100 po 10. Vytvorí **10 hodnôt** so začiatkom 100 a krokom 5.

### Nesprávny desatinný oddeľovač

Pri slovenskom nastavení používame napríklad:

```excel
=SEQUENCE(10;1;0;0,5)
```

nie:

```text
=SEQUENCE(10,1,0,0.5)
```

### Úprava bunky vo vnútri spill poľa

Jednotlivé výsledkové bunky dynamického poľa nie sú samostatné vzorce. Zdrojový vzorec upravujeme v ľavej hornej bunke poľa.

### Dynamické pole vo vnútri Excel Table

Ak potrebujeme, aby sa výsledok rozlieval do viacerých buniek, umiestnime dynamický vzorec mimo vypočítavaného stĺpca Excel Table.

<a id="dokumentacia"></a>
## 20. Oficiálna dokumentácia

Microsoft Support – SEQUENCE function:

https://support.microsoft.com/en-us/excel/functions/sequence-function

Funkcia je určená pre dynamické polia a je dostupná v podporovaných moderných verziách Excelu vrátane Microsoft 365, Excel 2024 a Excel 2021.
