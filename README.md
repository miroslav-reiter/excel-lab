# 🧪 Online Kurzy Microsoft Excel (excel-lab)

Testovací repozitár na praktické precvičovanie programu **Microsoft Excel** od úplných základov až po pokročilú analýzu dát, kontingenčné tabuľky, automatizáciu a programovanie vo **VBA (Visual Basic for Applications)**.

Repozitár nadväzuje na štyri online kurzy VITA Academy:

1. [Microsoft Excel I. Začiatočník](https://www.vita.sk/online-kurz-microsoft-excel-i-zaciatocnik/)
2. [Microsoft Excel II. Mierne Pokročilý](https://www.vita.sk/online-kurz-microsoft-excel-ii-mierne-pokrocily/)
3. [Microsoft Excel III. Pokročilý](https://www.vita.sk/online-kurz-microsoft-excel-iii-pokrocily/)
4. [Microsoft Excel IV. Expert: Programovanie vo VBA](https://www.vita.sk/online-kurz-microsoft-excel-iv-expert-programovanie-vo-vba/)

Cieľom repozitára je precvičiť prácu s bunkami, tabuľkami, formátovaním, vzorcami, funkciami, filtrami, grafmi, veľkými dátami, kontingenčnými tabuľkami, importom dát, makrami a objektovým modelom Microsoft Excelu na realistických firemných príkladoch.

> **Poznámka k vzorcom:** príklady používajú najmä anglické názvy funkcií. Podľa jazykového a regionálneho nastavenia Excelu môže byť oddeľovač argumentov `,` alebo `;`.

## 📑 Obsah (TOC)

1. [Excel I. Začiatočník](#excel-i-zaciatocnik)
   1. [Navigácia a práca s bunkami](#excel-i-01)
   2. [Relatívne a absolútne odkazy](#excel-i-02)
   3. [Základné funkcie](#excel-i-03)
   4. [Excelová tabuľka](#excel-i-04)
   5. [Odstránenie duplicít](#excel-i-05)
   6. [Filtrovanie a triedenie](#excel-i-06)
   7. [Formáty čísel a percentá](#excel-i-07)
   8. [Podmienené formátovanie](#excel-i-08)
   9. [Graf predaja](#excel-i-09)
   10. [Projekt – faktúra](#excel-i-10)
   11. [Projekt – evidencia dochádzky](#excel-i-11)
2. [Excel II. Mierne Pokročilý](#excel-ii-mierne-pokrocily)
   1. [Číselné a dátumové postupnosti](#excel-ii-12)
   2. [Vlastné zoznamy](#excel-ii-13)
   3. [Efektívne vkladanie dát](#excel-ii-14)
   4. [Pokročilá selekcia a Prejsť na](#excel-ii-15)
   5. [Dátumy a časy](#excel-ii-16)
   6. [Automatické filtre a divoké karty](#excel-ii-17)
   7. [Vlastné zoradenie](#excel-ii-18)
   8. [Paste Special](#excel-ii-19)
   9. [Pokročilé grafy](#excel-ii-20)
   10. [Tlač a zlomy strán](#excel-ii-21)
3. [Excel III. Pokročilý](#excel-iii-pokrocily)
   1. [Limity pracovného hárka](#excel-iii-22)
   2. [Volatilné funkcie](#excel-iii-23)
   3. [VLOOKUP](#excel-iii-24)
   4. [HLOOKUP](#excel-iii-25)
   5. [MATCH a INDEX](#excel-iii-26)
   6. [XLOOKUP ako moderná alternatíva](#excel-iii-27)
   7. [CHOOSE](#excel-iii-28)
   8. [Goal Seek – Hľadanie riešenia](#excel-iii-29)
   9. [SUMPRODUCT](#excel-iii-30)
   10. [Ošetrenie chýb](#excel-iii-31)
   11. [Projekt – HR absencie](#excel-iii-32)
   12. [Projekt – firemný rozpočet](#excel-iii-33)
   13. [Kontingenčná tabuľka](#excel-iii-34)
   14. [Unikátne hodnoty](#excel-iii-35)
   15. [Rýchla analýza](#excel-iii-36)
   16. [Import CSV](#excel-iii-37)
   17. [Informačné funkcie](#excel-iii-38)
4. [Excel IV. Expert – Programovanie vo VBA](#excel-iv-expert-vba)
   1. [Prvé makro](#excel-iv-39)
   2. [Range](#excel-iv-40)
   3. [Cells](#excel-iv-41)
   4. [Práca s hárkami](#excel-iv-42)
   5. [AutoFit](#excel-iv-43)
   6. [Premenné a Option Explicit](#excel-iv-44)
   7. [Podmienka If](#excel-iv-45)
   8. [For cyklus](#excel-iv-46)
   9. [For Step](#excel-iv-47)
   10. [For Each](#excel-iv-48)
   11. [Do While](#excel-iv-49)
   12. [Do Until a Exit Do](#excel-iv-50)
   13. [MsgBox a InputBox](#excel-iv-51)
   14. [Replace](#excel-iv-52)
   15. [Copy a Paste](#excel-iv-53)
   16. [Clear a Delete](#excel-iv-54)
   17. [Používateľská funkcia UDF](#excel-iv-55)
   18. [Automatické vytvorenie reportu](#excel-iv-56)
   19. [Optimalizácia VBA kódu](#excel-iv-57)
   20. [Ovládacie tlačidlo](#excel-iv-58)

## 📌 Čo je Microsoft Excel

Microsoft Excel je tabuľkový procesor určený na zadávanie, organizovanie, výpočty, analýzu a vizualizáciu dát. Dáta ukladáme do pracovných zošitov, ktoré obsahujú jeden alebo viac hárkov. Základným prvkom je bunka identifikovaná adresou, napríklad `A1`, `C15` alebo `F250`.

Excel môžeme používať ako jednoduchú tabuľku, ale aj ako analytický nástroj s funkciami, kontingenčnými tabuľkami, grafmi, importom dát a automatizáciou pomocou VBA.

## 🎯 Na čo používame Microsoft Excel v praxi

Typické použitie:

1. firemné reporty a manažérske prehľady
2. účtovníctvo, financie a rozpočty
3. evidencia faktúr a pohľadávok
4. evidencia dochádzky a HR reporty
5. obchodné a marketingové analýzy
6. skladové hospodárstvo a produktové cenníky
7. plánovanie projektov a harmonogramov
8. štatistické výpočty a analýza dát
9. vizualizácia údajov pomocou grafov
10. import a čistenie CSV, TXT a XML dát
11. kontingenčné tabuľky a súhrnné reporty
12. automatizácia opakovaných úloh pomocou makier a VBA

## 🧭 Štyri úrovne tréningu

| Úroveň | Kurz | Hlavné témy | Praktický výsledok |
|---|---|---|---|
| I. | Excel I. Začiatočník | prostredie, bunky, tabuľky, filtre, formátovanie, vzorce, základné funkcie, grafy | samostatná práca s bežným firemným zošitom |
| II. | Excel II. Mierne Pokročilý | postupnosti, selekcia, efektívne vkladanie dát, pokročilé filtrovanie, triedenie, grafy, tlač | rýchlejšia a systematickejšia práca s väčšími tabuľkami |
| III. | Excel III. Pokročilý | veľké dáta, lookupy, Goal Seek, SUMPRODUCT, rozpočty, HR, PivotTable, importy, informačné funkcie | pokročilá analýza a reporting |
| IV. | Excel IV. Expert – VBA | makrá, objektový model, Range, Cells, Worksheets, UDF, cykly, podmienky, ovládacie prvky | automatizované pracovné postupy a vlastné VBA riešenia |

## 🖥️ Odporúčané prostredie

Pre väčšinu cvičení odporúčame **Microsoft Excel pre Microsoft 365** alebo aktuálnu desktopovú verziu Microsoft Excelu pre Windows alebo macOS.

Pre kurz IV. používame desktopovú aplikáciu Excel. **Excel pre web nedokáže vytvárať, spúšťať ani upravovať VBA makrá.**

Odporúčané formáty:

| Formát | Použitie |
|---|---|
| `.xlsx` | štandardný zošit bez VBA makier |
| `.xlsm` | zošit s VBA makrami |
| `.xlsb` | binárny zošit, vhodný aj pri väčších súboroch |
| `.csv` | výmena jednoduchých tabuľkových dát |
| `.txt` | textové vstupy a exporty |
| `.xml` | štruktúrované dátové vstupy a výstupy |

## ⚙️ Príprava prostredia

### Microsoft Excel pre Windows

1. Spustíme Microsoft Excel
2. Vytvoríme nový prázdny zošit
3. Skontrolujeme kartu **Home**, **Insert**, **Page Layout**, **Formulas**, **Data**, **Review** a **View**
4. Pre VBA aktivujeme kartu **Developer** cez `File > Options > Customize Ribbon`
5. Zošity s makrami ukladáme ako `.xlsm`

### Microsoft Excel pre macOS

Väčšina cvičení z úrovní I až III funguje rovnako. VBA je dostupné aj v desktopovej verzii Excelu pre Mac, niektoré ovládacie prvky a systémové integrácie sa však môžu správať odlišne ako vo Windows.

### Excel pre web

Excel pre web môžeme používať na veľkú časť základných a analytických cvičení. Pri VBA prejdeme do desktopovej aplikácie.

## 🗂️ Navrhovaná štruktúra repozitára

```text
excel-lab/
├── README.md
├── data/
│   ├── employees.csv
│   ├── sales.csv
│   ├── invoices.csv
│   ├── attendance.csv
│   ├── budget.csv
│   ├── products.csv
│   └── lookup-products.csv
├── exercises/
│   ├── 01-excel-basics.md
│   ├── 02-tables-formatting.md
│   ├── 03-formulas-functions.md
│   ├── 04-filtering-sorting.md
│   ├── 05-charts-printing.md
│   ├── 06-lookups-analysis.md
│   ├── 07-pivottables.md
│   └── 08-vba.md
├── solutions/
│   ├── excel-01-basics.xlsx
│   ├── excel-02-intermediate.xlsx
│   ├── excel-03-advanced.xlsx
│   └── excel-04-vba.xlsm
├── vba/
│   ├── 01-first-macro.bas
│   ├── 02-range-cells.bas
│   ├── 03-worksheets.bas
│   ├── 04-conditions-loops.bas
│   ├── 05-udf.bas
│   └── 06-report-automation.bas
└── projects/
    ├── invoice/
    ├── attendance/
    ├── budget/
    ├── sales-report/
    └── hr-absence/
```

## 🚀 Rýchly štart

```bash
git clone https://github.com/miroslav-reiter/excel-lab.git
cd excel-lab
```

Následne otvoríme tréningové zošity v Microsoft Exceli a pokračujeme podľa úrovne kurzu.

---

<a id="excel-i-zaciatocnik"></a>
# 🟢 Excel I. Začiatočník

Kurz I. pokrýva používateľské prostredie, pracovné zošity a hárky, prácu s bunkami, tabuľkami, filtrovanie, triedenie, formátovanie, základné vzorce, funkcie, grafy a praktické firemné dokumenty.

<a id="excel-i-01"></a>
## 1. Navigácia a práca s bunkami

Vytvoríme tabuľku:

| A | B | C | D |
|---|---|---|---|
| Produkt | Cena | Množstvo | Tržba |
| Notebook | 850 | 2 | |
| Monitor | 240 | 4 | |
| Myš | 35 | 10 | |

Do bunky `D2` zadáme:

```excel
=B2*C2
```

Vzorec potiahneme smerom nadol.

Precvičíme:

1. výber bunky a rozsahu
2. presun pomocou šípok
3. `Ctrl + Arrow`
4. `Ctrl + Shift + Arrow`
5. editáciu bunky pomocou `F2`
6. kopírovanie a prilepenie
7. automatické dopĺňanie vzorcov

<a id="excel-i-02"></a>
## 2. Relatívne a absolútne odkazy

Do bunky `F1` zadáme sadzbu DPH:

```text
0,23
```

Do bunky `E2` vypočítame DPH:

```excel
=D2*$F$1
```

Rozdiel:

```text
D2      relatívny odkaz
$D$2    absolútny odkaz
D$2     zmiešaný odkaz – fixovaný riadok
$D2     zmiešaný odkaz – fixovaný stĺpec
```

Pri editácii odkazu môžeme použiť `F4`.

<a id="excel-i-03"></a>
## 3. Základné funkcie

Predpokladajme hodnoty tržieb v rozsahu `D2:D100`.

```excel
=SUM(D2:D100)
=AVERAGE(D2:D100)
=MIN(D2:D100)
=MAX(D2:D100)
=COUNT(D2:D100)
```

Praktická úloha:

1. vypočítame celkovú tržbu
2. priemernú objednávku
3. najnižšiu tržbu
4. najvyššiu tržbu
5. počet číselných záznamov

<a id="excel-i-04"></a>
## 4. Excelová tabuľka

Označíme dáta a použijeme:

```text
Ctrl + T
```

Tabuľke nastavíme názov:

```text
tblSales
```

Výhody:

1. automatické filtre
2. automatické rozšírenie
3. jednotné formátovanie
4. štruktúrované odkazy
5. riadok súčtu

Príklad štruktúrovaného odkazu:

```excel
=[@Cena]*[@Množstvo]
```

<a id="excel-i-05"></a>
## 5. Odstránenie duplicít

Pripravíme zoznam zákazníkov s opakovanými e-mailami.

Použijeme:

```text
Data > Remove Duplicates
```

Pred odstránením duplicít si údaje zálohujeme alebo vytvoríme kópiu hárka.

<a id="excel-i-06"></a>
## 6. Filtrovanie a triedenie

Na dátach `sales.csv` precvičíme:

1. filter podľa regiónu
2. filter podľa obchodníka
3. číselný filter `Tržba > 1000`
4. zoradenie podľa tržby zostupne
5. viacúrovňové zoradenie `Región > Obchodník > Tržba`

<a id="excel-i-07"></a>
## 7. Formáty čísel a percentá

Príklady:

```text
1250,5       → 1 250,50 €
0,235        → 23,50 %
2026-09-12   → 12. 9. 2026
0,5          → 12:00
```

Dôležité pravidlo: percentuálny formát nemení matematický význam hodnoty. Hodnota `0,23` reprezentuje `23 %`.

<a id="excel-i-08"></a>
## 8. Podmienené formátovanie

Pre tržby použijeme pravidlá:

```text
Tržba < 500          červené zvýraznenie
Tržba 500 až 2000    stredná kategória
Tržba > 2000         zelené zvýraznenie
```

Precvičíme:

1. Highlight Cells Rules
2. Data Bars
3. Color Scales
4. vlastné pravidlo so vzorcom

Príklad vlastného pravidla:

```excel
=$D2>2000
```

<a id="excel-i-09"></a>
## 9. Graf predaja

Z tabuľky:

| Mesiac | Tržba |
|---|---:|
| Január | 12000 |
| Február | 14500 |
| Marec | 13800 |
| Apríl | 17200 |

vytvoríme:

1. stĺpcový graf na porovnanie mesiacov
2. čiarový graf na sledovanie trendu

Upravíme:

1. názov grafu
2. legendu
3. os Y
4. dátové menovky
5. formát čísel

<a id="excel-i-10"></a>
## 10. Projekt – faktúra

Navrhneme jednoduchú faktúru so stĺpcami:

```text
Kód | Produkt | Množstvo | Cena bez DPH | DPH | Cena s DPH
```

Vzorce:

```excel
=Množstvo*Cena_bez_DPH
=Medzisúčet*Sadzba_DPH
=Medzisúčet+DPH
```

Cieľom je precvičiť vzorce, odkazy, formátovanie, tlač a praktickú prácu so zošitom.

<a id="excel-i-11"></a>
## 11. Projekt – evidencia dochádzky

Stĺpce:

```text
Dátum | Zamestnanec | Príchod | Odchod | Prestávka | Odpracované hodiny
```

Príklad výpočtu:

```excel
=(D2-C2)-E2
```

Výsledok naformátujeme ako čas alebo vlastný formát podľa požadovaného reportu.

---

<a id="excel-ii-mierne-pokrocily"></a>
# 🟡 Excel II. Mierne Pokročilý

Kurz II. rozvíja efektívnu prácu s prostredím, postupnosťami, vkladaním dát, selekciami, tabuľkami, vlastným triedením, filtrami, komentármi, grafmi a tlačou.

<a id="excel-ii-12"></a>
## 12. Číselné a dátumové postupnosti

Do buniek zadáme:

```text
1
2
```

Označíme ich a potiahneme úchyt výplne.

Precvičíme:

```text
1, 2, 3, 4, ...
10, 20, 30, 40, ...
1.1.2026, 2.1.2026, 3.1.2026, ...
Pondelok, Utorok, Streda, ...
```

Moderný dynamický príklad:

```excel
=SEQUENCE(100)
```

alebo mriežka 10 × 5:

```excel
=SEQUENCE(10,5,1,1)
```

<a id="excel-ii-13"></a>
## 13. Vlastné zoznamy

Vytvoríme vlastný zoznam:

```text
Junior
Medior
Senior
Lead
Manager
```

Použijeme ho pri automatickom dopĺňaní a vlastnom triedení.

<a id="excel-ii-14"></a>
## 14. Efektívne vkladanie dát

Precvičíme:

```text
Enter
Shift + Enter
Tab
Shift + Tab
Ctrl + Enter
Alt + Enter
```

Príklad: označíme rozsah `B2:B50`, zadáme text `Bratislava` a použijeme:

```text
Ctrl + Enter
```

Hodnota sa vloží do celého označeného rozsahu.

<a id="excel-ii-15"></a>
## 15. Pokročilá selekcia a Prejsť na

Použijeme:

```text
Ctrl + G
```

alebo:

```text
F5
```

Precvičíme **Go To Special**:

1. prázdne bunky
2. vzorce
3. konštanty
4. viditeľné bunky
5. komentáre a poznámky

<a id="excel-ii-16"></a>
## 16. Dátumy a časy

Príklady:

```excel
=TODAY()
=NOW()
=YEAR(A2)
=MONTH(A2)
=DAY(A2)
```

Praktická úloha: z dátumu faktúry vytvoríme stĺpce `Rok`, `Mesiac` a `Deň`.

<a id="excel-ii-17"></a>
## 17. Automatické filtre a divoké karty

Pri textových filtroch môžeme pracovať s divokými kartami:

```text
*    ľubovoľný počet znakov
?    jeden ľubovoľný znak
~    escape znak
```

Príklad:

```text
Rei*
```

vyhľadá texty začínajúce na `Rei`.

<a id="excel-ii-18"></a>
## 18. Vlastné zoradenie

Úroveň priority:

```text
Kritická
Vysoká
Stredná
Nízka
```

Namiesto abecedného zoradenia vytvoríme vlastné poradie a použijeme:

```text
Data > Sort > Custom List
```

<a id="excel-ii-19"></a>
## 19. Paste Special

Precvičíme:

```text
Values
Formulas
Formats
Transpose
Multiply
Divide
Add
Subtract
```

Praktický príklad: zvýšime všetky ceny o `1,05` pomocou operácie **Multiply** bez vytvárania pomocného stĺpca.

<a id="excel-ii-20"></a>
## 20. Pokročilé grafy

Na mesačných dátach vytvoríme:

1. stĺpcový graf
2. čiarový graf
3. koláčový graf

Porovnáme, kedy je každý typ vhodný.

Zásada: koláčový graf používame najmä na jednoduchý podiel niekoľkých kategórií na celku. Pri veľkom počte kategórií použijeme radšej stĺpcový alebo pruhový graf.

<a id="excel-ii-21"></a>
## 21. Tlač a zlomy strán

Precvičíme:

```text
Page Layout
Print Area
Orientation
Margins
Scale to Fit
Page Break Preview
Headers & Footers
```

Úloha: pripraviť report na jednu stranu A4 bez odrezaných stĺpcov.

---

<a id="excel-iii-pokrocily"></a>
# 🟠 Excel III. Pokročilý

Kurz III. sa zameriava na veľké dátové súbory, volatilné a vyhľadávacie funkcie, Goal Seek, SUMPRODUCT, firemné rozpočty, HR absencie, kontingenčné tabuľky, analýzu dát, importy a informačné funkcie.

<a id="excel-iii-22"></a>
## 22. Limity pracovného hárka

Aktuálne desktopové verzie Excelu podporujú maximálne:

```text
1 048 576 riadkov
16 384 stĺpcov
```

Posledný stĺpec je:

```text
XFD
```

Pri veľkých dátach sledujeme:

1. dostupnú RAM
2. počet vzorcov
3. volatilné funkcie
4. veľkosť súboru
5. počet formátov
6. externé prepojenia
7. dátový model

<a id="excel-iii-23"></a>
## 23. Volatilné funkcie

Medzi funkcie, ktoré môžu spôsobovať častejšie prepočítavanie zošita, patria napríklad:

```excel
=NOW()
=TODAY()
=RAND()
=RANDBETWEEN(1,100)
=OFFSET(...)
=INDIRECT(...)
```

Pri veľkých zošitoch ich používame uvážene.

<a id="excel-iii-24"></a>
## 24. VLOOKUP

Tabuľka produktov:

| Kód | Produkt | Cena |
|---|---|---:|
| P001 | Notebook | 850 |
| P002 | Monitor | 240 |
| P003 | Myš | 35 |

Vyhľadanie ceny:

```excel
=VLOOKUP(A2,$H$2:$J$100,3,FALSE)
```

Dávame pozor na:

1. správny vyhľadávací stĺpec
2. číslo výsledného stĺpca
3. presnú zhodu `FALSE`
4. fixovanie rozsahu

<a id="excel-iii-25"></a>
## 25. HLOOKUP

Pri horizontálnej tabuľke:

```excel
=HLOOKUP(B1,$B$10:$M$12,3,FALSE)
```

Používame ho vtedy, keď hľadané hodnoty ležia v prvom riadku tabuľky.

<a id="excel-iii-26"></a>
## 26. MATCH a INDEX

Pozícia hodnoty:

```excel
=MATCH(A2,$H$2:$H$100,0)
```

Vyhľadanie hodnoty:

```excel
=INDEX($J$2:$J$100,MATCH(A2,$H$2:$H$100,0))
```

Táto kombinácia je flexibilnejšia než klasický `VLOOKUP`, pretože výsledný stĺpec nemusí byť napravo od vyhľadávacieho stĺpca.

<a id="excel-iii-27"></a>
## 27. XLOOKUP ako moderná alternatíva

V podporovaných verziách Excelu môžeme použiť:

```excel
=XLOOKUP(A2,$H$2:$H$100,$J$2:$J$100,"Nenájdené")
```

Tento príklad je doplnkový k témam lookupov z kurzu.

<a id="excel-iii-28"></a>
## 28. CHOOSE

Príklad:

```excel
=CHOOSE(B2,"Nízka","Stredná","Vysoká","Kritická")
```

Ak `B2 = 3`, výsledkom bude:

```text
Vysoká
```

<a id="excel-iii-29"></a>
## 29. Goal Seek – Hľadanie riešenia

Model:

```text
Cena produktu       100 €
Počet kusov         100
Tržba               10 000 €
Cieľová tržba       15 000 €
```

Použijeme:

```text
Data > What-If Analysis > Goal Seek
```

Nastavenie:

```text
Set cell:       Tržba
To value:       15000
By changing:    Počet kusov
```

Excel vypočíta požadovaný počet kusov.

<a id="excel-iii-30"></a>
## 30. SUMPRODUCT

Máme stĺpce:

```text
Cena
Množstvo
```

Celkovú hodnotu vypočítame:

```excel
=SUMPRODUCT(B2:B100,C2:C100)
```

Bez pomocného stĺpca získame súčet všetkých `Cena × Množstvo`.

<a id="excel-iii-31"></a>
## 31. Ošetrenie chýb

Typické chyby:

```text
#DIV/0!
#N/A
#VALUE!
#REF!
#NAME?
#NUM!
#SPILL!
```

Príklad:

```excel
=IFERROR(VLOOKUP(A2,$H$2:$J$100,3,FALSE),"Nenájdené")
```

Chybu neprekrývame automaticky bez analýzy príčiny. `IFERROR` používame až vtedy, keď rozumieme tomu, prečo môže chyba vzniknúť.

<a id="excel-iii-32"></a>
## 32. Projekt – HR absencie

Stĺpce:

```text
ID | Zamestnanec | Oddelenie | Dátum | Typ absencie | Počet hodín
```

Úlohy:

1. filtrovať absencie podľa oddelenia
2. vypočítať celkové hodiny absencie
3. zistiť priemer na zamestnanca
4. vytvoriť kontingenčnú tabuľku
5. vytvoriť graf absencií podľa mesiaca

<a id="excel-iii-33"></a>
## 33. Projekt – firemný rozpočet

Stĺpce:

```text
Mesiac | Oddelenie | Rozpočet | Skutočnosť | Odchýlka | Odchýlka %
```

Vzorce:

```excel
=Skutočnosť-Rozpočet
```

```excel
=IFERROR((Skutočnosť-Rozpočet)/Rozpočet,0)
```

Pridáme podmienené formátovanie pre prekročenie rozpočtu.

<a id="excel-iii-34"></a>
## 34. Kontingenčná tabuľka

Zo `sales.csv` vytvoríme PivotTable:

```text
Rows:       Región
Columns:    Mesiac
Values:     Sum of Tržba
Filter:     Produkt
```

Následne:

1. zmeníme agregáciu zo `Sum` na `Average`
2. zoradíme regióny podľa tržby
3. pridáme percento z celku
4. obnovíme dáta po zmene zdroja

<a id="excel-iii-35"></a>
## 35. Unikátne hodnoty

Klasický spôsob:

```text
Data > Remove Duplicates
```

Moderný dynamický spôsob:

```excel
=UNIQUE(A2:A1000)
```

Zoradený unikátny zoznam:

```excel
=SORT(UNIQUE(A2:A1000))
```

<a id="excel-iii-36"></a>
## 36. Rýchla analýza

Označíme rozsah a použijeme nástroj **Quick Analysis**.

Precvičíme:

1. formátovanie
2. grafy
3. totals
4. tables
5. sparklines

<a id="excel-iii-37"></a>
## 37. Import CSV

Pri importe CSV sledujeme:

```text
oddeľovač
kódovanie
desatinný oddeľovač
dátumový formát
typy stĺpcov
vedúce nuly
```

Príklad rizikového kódu produktu:

```text
001245
```

Ak Excel stĺpec automaticky interpretuje ako číslo, vedúce nuly sa môžu stratiť. Takýto stĺpec importujeme ako text.

<a id="excel-iii-38"></a>
## 38. Informačné funkcie

Príklady:

```excel
=INFO("osversion")
=SHEET()
=SHEETS()
=ISTEXT(A2)
=TYPE(A2)
=CELL("address",A2)
=CELL("filename",A2)
```

Výsledky závisia od prostredia, typu hodnoty a konkrétnej funkcie.

---

<a id="excel-iv-expert-vba"></a>
# 🔴 Excel IV. Expert – Programovanie vo VBA

Kurz IV. sa zameriava na automatizáciu, záznam makier, objektový model Excelu, objekty `Range`, `Cells` a `Worksheets`, používateľské funkcie, premenné, podmienky, cykly, ovládacie prvky a optimalizáciu.

<a id="excel-iv-39"></a>
## 39. Prvé makro

V editore VBA vložíme modul:

```text
Developer > Visual Basic > Insert > Module
```

Prvé makro:

```vb
Option Explicit

Sub HelloExcel()
    MsgBox "Ahoj Excel!"
End Sub
```

<a id="excel-iv-40"></a>
## 40. Range

Zapísanie hodnoty do bunky:

```vb
Sub WriteToRange()
    Worksheets("Sheet1").Range("A1").Value = "VITA Academy"
End Sub
```

Zápis do rozsahu:

```vb
Sub FillRange()
    Worksheets("Sheet1").Range("A1:A10").Value = "Excel"
End Sub
```

<a id="excel-iv-41"></a>
## 41. Cells

Alternatívne adresovanie:

```vb
Sub WriteWithCells()
    Worksheets("Sheet1").Cells(1, 1).Value = "A1"
    Worksheets("Sheet1").Cells(2, 3).Value = "C2"
End Sub
```

Syntax:

```text
Cells(riadok, stĺpec)
```

<a id="excel-iv-42"></a>
## 42. Práca s hárkami

Pridanie hárka:

```vb
Sub AddWorksheet()
    Worksheets.Add
End Sub
```

Pridanie a pomenovanie:

```vb
Sub AddReportWorksheet()
    Dim ws As Worksheet

    Set ws = Worksheets.Add
    ws.Name = "Report"
End Sub
```

Počet hárkov:

```vb
Sub CountWorksheets()
    MsgBox Worksheets.Count
End Sub
```

<a id="excel-iv-43"></a>
## 43. AutoFit

```vb
Sub AutoFitReport()
    With Worksheets("Report")
        .Columns.AutoFit
        .Rows.AutoFit
    End With
End Sub
```

<a id="excel-iv-44"></a>
## 44. Premenné a Option Explicit

```vb
Option Explicit

Sub VariablesExample()
    Dim revenue As Double
    Dim customerName As String
    Dim rowCount As Long

    revenue = 12500.5
    customerName = "VITA Academy"
    rowCount = 1000

    MsgBox customerName & ": " & revenue
End Sub
```

`Option Explicit` vyžaduje deklaráciu premenných a pomáha odhaliť preklepy v názvoch.

<a id="excel-iv-45"></a>
## 45. Podmienka If

```vb
Sub CheckRevenue()
    Dim revenue As Double

    revenue = Range("A1").Value

    If revenue >= 10000 Then
        Range("B1").Value = "Cieľ splnený"
    Else
        Range("B1").Value = "Cieľ nesplnený"
    End If
End Sub
```

<a id="excel-iv-46"></a>
## 46. For cyklus

```vb
Sub NumberRows()
    Dim i As Long

    For i = 2 To 101
        Cells(i, 1).Value = i - 1
    Next i
End Sub
```

<a id="excel-iv-47"></a>
## 47. For Step

```vb
Sub EverySecondRow()
    Dim i As Long

    For i = 2 To 100 Step 2
        Cells(i, 1).Value = "Párny riadok"
    Next i
End Sub
```

<a id="excel-iv-48"></a>
## 48. For Each

```vb
Sub HighlightNegativeValues()
    Dim cell As Range

    For Each cell In Range("D2:D100")
        If IsNumeric(cell.Value) And cell.Value < 0 Then
            cell.Font.Bold = True
        End If
    Next cell
End Sub
```

<a id="excel-iv-49"></a>
## 49. Do While

```vb
Sub DoWhileExample()
    Dim i As Long

    i = 2

    Do While Cells(i, 1).Value <> ""
        Cells(i, 2).Value = UCase(Cells(i, 1).Value)
        i = i + 1
    Loop
End Sub
```

<a id="excel-iv-50"></a>
## 50. Do Until a Exit Do

```vb
Sub FindStopValue()
    Dim i As Long

    i = 2

    Do Until Cells(i, 1).Value = ""
        If Cells(i, 1).Value = "STOP" Then
            Exit Do
        End If

        i = i + 1
    Loop
End Sub
```

<a id="excel-iv-51"></a>
## 51. MsgBox a InputBox

```vb
Sub AskUser()
    Dim name As String

    name = InputBox("Zadajte meno používateľa:")

    If name <> "" Then
        MsgBox "Ahoj " & name
    End If
End Sub
```

<a id="excel-iv-52"></a>
## 52. Replace

```vb
Sub ReplaceText()
    Worksheets("Data").Range("A:A").Replace _
        What:="s.r.o.", _
        Replacement:="s. r. o.", _
        LookAt:=xlPart
End Sub
```

<a id="excel-iv-53"></a>
## 53. Copy a Paste

Bez použitia schránky môžeme kopírovať efektívnejšie:

```vb
Sub CopyValues()
    Worksheets("Report").Range("A1:D100").Value = _
        Worksheets("Data").Range("A1:D100").Value
End Sub
```

Pri kopírovaní iba hodnôt je tento prístup vhodnejší než opakované `Copy` a `Paste`.

<a id="excel-iv-54"></a>
## 54. Clear a Delete

Vymazanie obsahu:

```vb
Range("A2:D100").ClearContents
```

Vymazanie obsahu aj formátovania:

```vb
Range("A2:D100").Clear
```

Odstránenie celého riadka:

```vb
Rows(10).Delete
```

<a id="excel-iv-55"></a>
## 55. Používateľská funkcia UDF

Vytvoríme vlastnú funkciu na výpočet ceny s DPH:

```vb
Option Explicit

Function PriceWithVAT(priceWithoutVAT As Double, vatRate As Double) As Double
    PriceWithVAT = priceWithoutVAT * (1 + vatRate)
End Function
```

V Exceli ju použijeme:

```excel
=PriceWithVAT(A2,0.23)
```

<a id="excel-iv-56"></a>
## 56. Automatické vytvorenie reportu

```vb
Option Explicit

Sub CreateReport()
    Dim wsData As Worksheet
    Dim wsReport As Worksheet
    Dim lastRow As Long

    Set wsData = Worksheets("Data")

    On Error Resume Next
    Set wsReport = Worksheets("Report")
    On Error GoTo 0

    If wsReport Is Nothing Then
        Set wsReport = Worksheets.Add
        wsReport.Name = "Report"
    Else
        wsReport.Cells.Clear
    End If

    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row

    wsData.Range("A1:D" & lastRow).Copy Destination:=wsReport.Range("A1")

    wsReport.Columns.AutoFit

    MsgBox "Report bol vytvorený."
End Sub
```

<a id="excel-iv-57"></a>
## 57. Optimalizácia VBA kódu

Pri väčších operáciách môžeme dočasne vypnúť prekresľovanie obrazovky a automatický prepočet.

```vb
Option Explicit

Sub OptimizedProcess()
    On Error GoTo CleanUp

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' Spracovanie dát

CleanUp:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
```

Dôležité je nastavenia na konci vždy obnoviť, aj keď počas makra vznikne chyba.

<a id="excel-iv-58"></a>
## 58. Ovládacie tlačidlo

Vložíme tlačidlo:

```text
Developer > Insert > Button
```

Tlačidlu priradíme napríklad makro:

```text
CreateReport
```

Takto môže používateľ spustiť report bez vstupu do VBA editora.

---

# 🧠 Excel cheat sheet

## 📍 Najčastejšie klávesové skratky

| Akcia | Skratka |
|---|---|
| Uložiť zošit | `Ctrl + S` |
| Kopírovať | `Ctrl + C` |
| Vystrihnúť | `Ctrl + X` |
| Prilepiť | `Ctrl + V` |
| Späť | `Ctrl + Z` |
| Znova | `Ctrl + Y` |
| Upraviť bunku | `F2` |
| Prepnúť typ odkazu vo vzorci | `F4` |
| Prejsť na | `Ctrl + G` alebo `F5` |
| Vytvoriť tabuľku | `Ctrl + T` |
| Zapnúť/vypnúť filter | `Ctrl + Shift + L` |
| Vložiť aktuálny dátum | `Ctrl + ;` |
| Nový riadok v bunke | `Alt + Enter` |
| Vybrať súvislú oblasť | `Ctrl + Shift + Arrow` |

## 🧮 Základné funkcie

| Úloha | Funkcia |
|---|---|
| Súčet | `SUM` |
| Priemer | `AVERAGE` |
| Minimum | `MIN` |
| Maximum | `MAX` |
| Počet čísel | `COUNT` |
| Logická podmienka | `IF` |
| Ošetrenie chyby | `IFERROR` |
| Vyhľadanie | `VLOOKUP`, `HLOOKUP`, `XLOOKUP` |
| Pozícia hodnoty | `MATCH` |
| Hodnota podľa pozície | `INDEX` |
| Súčet súčinov | `SUMPRODUCT` |
| Unikátne hodnoty | `UNIQUE` |
| Zoradenie poľa | `SORT` |
| Postupnosť | `SEQUENCE` |
| Dnešný dátum | `TODAY` |
| Dátum a čas | `NOW` |

## 🔗 Typy odkazov

| Typ | Príklad |
|---|---|
| Relatívny | `A1` |
| Absolútny | `$A$1` |
| Fixovaný stĺpec | `$A1` |
| Fixovaný riadok | `A$1` |

## ⚠️ Najčastejšie chybové hodnoty

| Chyba | Typický význam |
|---|---|
| `#DIV/0!` | delenie nulou |
| `#N/A` | hodnota nebola nájdená |
| `#VALUE!` | nesprávny typ hodnoty alebo argumentu |
| `#REF!` | neplatný odkaz |
| `#NAME?` | Excel nerozpoznal názov |
| `#NUM!` | problém s číselnou hodnotou |
| `#SPILL!` | dynamické pole sa nemôže rozliať do cieľového rozsahu |

## 🧱 VBA objektový model – základ

```text
Application
└── Workbook
    └── Worksheet
        ├── Range
        ├── Cells
        ├── Rows
        └── Columns
```

Príklady:

```vb
Application.ActiveWorkbook.Name
ThisWorkbook.Name
Worksheets("Data").Name
Range("A1").Value
Cells(1, 1).Value
Rows(1).Hidden = True
Columns("A").AutoFit
```

---

# 🧪 Praktické projekty

## Projekt 1 – Faktúra

Precvičíme:

1. formátovanie buniek
2. absolútne odkazy
3. DPH
4. medzisúčty
5. tlač
6. PDF export

## Projekt 2 – Dochádzka

Precvičíme:

1. dátumy
2. časy
3. rozdiel časov
4. podmienené formátovanie
5. SUM a AVERAGE

## Projekt 3 – Sales report

Precvičíme:

1. Excel Tables
2. filtre
3. triedenie
4. lookupy
5. SUMPRODUCT
6. grafy
7. PivotTable

## Projekt 4 – Firemný rozpočet

Precvičíme:

1. rozpočet vs. skutočnosť
2. percentuálnu odchýlku
3. podmienené formátovanie
4. Goal Seek
5. manažérsky report

## Projekt 5 – HR absencie

Precvičíme:

1. import dát
2. filtre
3. dátumy
4. unikátne hodnoty
5. PivotTable
6. grafy

## Projekt 6 – VBA report generator

Precvičíme:

1. makrá
2. `Range`
3. `Cells`
4. `Worksheets`
5. premenné
6. podmienky
7. cykly
8. dynamický posledný riadok
9. kopírovanie dát
10. AutoFit
11. tlačidlo na spustenie procedúry

---

# 🧭 Odporúčané poradie tréningu

1. Začneme orientáciou v prostredí Excelu, bunkami, hárkami a zošitmi
2. Precvičíme vkladanie dát, selekciu a formátovanie
3. Prejdeme na Excel Tables, filtre, triedenie a odstránenie duplicít
4. Naučíme sa relatívne, absolútne a zmiešané odkazy
5. Precvičíme základné funkcie a praktické výpočty
6. Vytvoríme grafy a pripravíme report na tlač
7. Doplníme efektívne vkladanie dát, postupnosti a pokročilú selekciu
8. Prejdeme na lookupy, Goal Seek, SUMPRODUCT a prácu s chybami
9. Vytvoríme kontingenčné tabuľky a súhrnné reporty
10. Precvičíme importy CSV a informačné funkcie
11. Aktivujeme Developer a zaznamenáme prvé makro
12. Prejdeme na objektový model, `Range`, `Cells` a `Worksheets`
13. Doplníme premenné, podmienky a cykly
14. Vytvoríme vlastnú UDF funkciu
15. Zautomatizujeme kompletný firemný report pomocou VBA

# ⚠️ Časté chyby

## Číslo uložené ako text

Príklad:

```text
"1250"
```

môže vyzerať ako číslo, ale Excel s ním nemusí počítať ako s číslom.

Overíme:

```excel
=ISNUMBER(A2)
```

## Dátum uložený ako text

Dátum môže vyzerať správne, ale byť uložený ako reťazec. Overíme formát a zdroj importu.

## Nesprávne fixovaný rozsah

Chybne:

```excel
=VLOOKUP(A2,H2:J100,3,FALSE)
```

Po kopírovaní vzorca sa rozsah posúva.

Správne:

```excel
=VLOOKUP(A2,$H$2:$J$100,3,FALSE)
```

## Zlúčené bunky v dátovej tabuľke

Merge Cells používame opatrne. V dátových tabuľkách komplikujú filtrovanie, triedenie, importy aj automatizáciu.

## Prázdne riadky v dátach

Súvislú dátovú tabuľku zbytočne neprerušujeme prázdnymi riadkami a stĺpcami.

## Ručné farby namiesto dátovej logiky

Farbu nepoužívame ako jediný nositeľ informácie. Ak je kategória dôležitá pre analýzu, vytvoríme samostatný stĺpec.

## Hard-coded hodnoty vo vzorcoch

Menej vhodné:

```excel
=A2*1.23
```

Lepšie:

```excel
=A2*(1+$F$1)
```

Sadzbu DPH uložíme do samostatnej bunky.

## Použitie celých stĺpcov v náročných vzorcoch

Pri veľkom množstve náročných vzorcov môže byť:

```excel
=SUMPRODUCT(A:A,B:B)
```

zbytočne náročné.

Použijeme radšej reálny rozsah alebo Excel Table.

## Neuloženie VBA zošita ako XLSM

Ak zošit s VBA uložíme ako `.xlsx`, VBA projekt sa neuloží.

Použijeme:

```text
Excel Macro-Enabled Workbook (*.xlsm)
```

## Makro závislé od ActiveCell

Menej stabilný kód:

```vb
ActiveCell.Value = "OK"
```

Lepšie explicitné adresovanie:

```vb
Worksheets("Data").Range("A1").Value = "OK"
```

# 🔒 Bezpečnostné poznámky

Pri Excel súboroch a VBA makrách dodržiavame tieto pravidlá:

1. nespúšťame makrá z neznámych alebo nedôveryhodných zdrojov
2. neaktivujeme obsah iba preto, že nás na to vyzve neznámy dokument
3. pred spustením makra skontrolujeme VBA kód
4. produkčné dáta nahrádzame anonymizovanými tréningovými dátami
5. do GitHub repozitára neukladáme heslá, API tokeny, osobné údaje ani dôverné firemné exporty
6. pri práci s externými CSV, XLSX a XLSM súbormi overujeme pôvod dát
7. makrá zapisujúce alebo odstraňujúce dáta testujeme najprv na kópii zošita
8. pri automatizácii používame explicitné názvy zošitov, hárkov a rozsahov

# 📚 Užitočné odkazy a zdroje

## VITA Academy – online kurzy

1. Microsoft Excel I. Začiatočník  
   https://www.vita.sk/online-kurz-microsoft-excel-i-zaciatocnik/
2. Microsoft Excel II. Mierne Pokročilý  
   https://www.vita.sk/online-kurz-microsoft-excel-ii-mierne-pokrocily/
3. Microsoft Excel III. Pokročilý  
   https://www.vita.sk/online-kurz-microsoft-excel-iii-pokrocily/
4. Microsoft Excel IV. Expert: Programovanie vo VBA  
   https://www.vita.sk/online-kurz-microsoft-excel-iv-expert-programovanie-vo-vba/

## Microsoft dokumentácia

1. Excel specifications and limits  
   https://support.microsoft.com/en-us/excel/excel-specifications-and-limits
2. Klávesové skratky v Exceli  
   https://support.microsoft.com/sk-sk/accessibility/excel/keyboard-shortcuts-in-excel
3. Vytváranie a formátovanie tabuliek  
   https://support.microsoft.com/sk-sk/excel/get-started/create-and-format-tables
4. Overview of PivotTables and PivotCharts  
   https://support.microsoft.com/en-us/excel/overview-of-pivottables-and-pivotcharts
5. VBA Range object  
   https://learn.microsoft.com/sk-sk/office/vba/api/excel.range(object)
6. VBA makrá v Exceli pre web  
   https://support.microsoft.com/sk-sk/excel/work-with-vba-macros-in-excel-for-the-web

# ✅ Cieľ repozitára

Po prejdení príkladov by sme mali vedieť:

1. samostatne vytvárať a upravovať Excel zošity
2. efektívne pracovať s bunkami, rozsahmi a tabuľkami
3. filtrovať, triediť a formátovať dáta
4. používať vzorce, funkcie a rôzne typy odkazov
5. vytvárať grafy a tlačové reporty
6. analyzovať väčšie dátové súbory
7. používať lookupy, SUMPRODUCT a informačné funkcie
8. vytvárať a analyzovať kontingenčné tabuľky
9. importovať CSV a ďalšie dátové formáty
10. nahrávať a upravovať makrá
11. programovať základné aj praktické VBA procedúry
12. automatizovať opakované úlohy v Microsoft Exceli
