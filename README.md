# 🧪 Online Kurzy Microsoft Excel (excel-lab)

Praktický repozitár na precvičovanie programu **Microsoft Excel** od úplných základov až po pokročilú analýzu dát, kontingenčné tabuľky, automatizáciu a programovanie vo **VBA (Visual Basic for Applications)**.

Repozitár nadväzuje na štyri online kurzy VITA Academy:

1. [Microsoft Excel I. Začiatočník](https://www.vita.sk/online-kurz-microsoft-excel-i-zaciatocnik/)
2. [Microsoft Excel II. Mierne Pokročilý](https://www.vita.sk/online-kurz-microsoft-excel-ii-mierne-pokrocily/)
3. [Microsoft Excel III. Pokročilý](https://www.vita.sk/online-kurz-microsoft-excel-iii-pokrocily/)
4. [Microsoft Excel IV. Expert: Programovanie vo VBA](https://www.vita.sk/online-kurz-microsoft-excel-iv-expert-programovanie-vo-vba/)

> **Pravidlo názvov súborov a priečinkov:** názvy v repozitári sú po slovensky, ale zapisujeme ich bez diakritiky a bez medzier. Technické názvy funkcií a objektov Excel/VBA zostávajú v originálnom tvare.

## 📑 Obsah

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
   6. [Automatické filtre a zástupné znaky](#excel-ii-17)
   7. [Vlastné zoradenie](#excel-ii-18)
   8. [Prilepiť špeciálne](#excel-ii-19)
   9. [Pokročilé grafy](#excel-ii-20)
   10. [Tlač a zlomy strán](#excel-ii-21)
3. [Excel III. Pokročilý](#excel-iii-pokrocily)
   1. [Limity pracovného hárka](#excel-iii-22)
   2. [Volatilné funkcie](#excel-iii-23)
   3. [VLOOKUP](#excel-iii-24)
   4. [HLOOKUP](#excel-iii-25)
   5. [MATCH a INDEX](#excel-iii-26)
   6. [XLOOKUP](#excel-iii-27)
   7. [CHOOSE](#excel-iii-28)
   8. [Hľadanie riešenia](#excel-iii-29)
   9. [SUMPRODUCT](#excel-iii-30)
   10. [Ošetrenie chýb](#excel-iii-31)
   11. [Projekt – absencie zamestnancov](#excel-iii-32)
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
   5. [Automatická šírka a výška](#excel-iv-43)
   6. [Premenné a Option Explicit](#excel-iv-44)
   7. [Podmienka If](#excel-iv-45)
   8. [Cyklus For](#excel-iv-46)
   9. [For Step](#excel-iv-47)
   10. [For Each](#excel-iv-48)
   11. [Do While](#excel-iv-49)
   12. [Do Until a Exit Do](#excel-iv-50)
   13. [MsgBox a InputBox](#excel-iv-51)
   14. [Nahradenie textu](#excel-iv-52)
   15. [Kopírovanie a prilepenie](#excel-iv-53)
   16. [Vymazanie obsahu a riadkov](#excel-iv-54)
   17. [Používateľská funkcia UDF](#excel-iv-55)
   18. [Automatické vytvorenie prehľadu](#excel-iv-56)
   19. [Optimalizácia VBA kódu](#excel-iv-57)
   20. [Ovládacie tlačidlo](#excel-iv-58)
5. [Praktické súbory repozitára](#prakticke-subory)
6. [Praktické projekty](#prakticke-projekty)

<a id="prakticke-subory"></a>
## 🗂️ Praktické súbory repozitára

```text
excel-lab/
├── README.md
├── data/
│   ├── zamestnanci.csv
│   ├── predaj.csv
│   ├── faktury.csv
│   ├── dochadzka.csv
│   ├── rozpocet.csv
│   ├── produkty.csv
│   └── vyhladavanie-produktov.csv
├── cvicenia/
│   ├── 01-zaklady-excelu.md
│   ├── 02-tabulky-a-formatovanie.md
│   ├── 03-vzorce-a-funkcie.md
│   ├── 04-filtrovanie-a-triedenie.md
│   ├── 05-grafy-a-tlac.md
│   ├── 06-vyhladavanie-a-analyza.md
│   ├── 07-kontingencne-tabulky.md
│   └── 08-vba-automatizacia.md
├── riesenia/
│   ├── excel-01-zaciatocnik.xlsx
│   ├── excel-02-mierne-pokrocily.xlsx
│   ├── excel-03-pokrocily.xlsx
│   └── excel-04-expert-vba.xlsm
├── vba-ukazky/
│   ├── 01-prve-makro.bas
│   ├── 02-rozsahy-a-bunky.bas
│   ├── 03-harky.bas
│   ├── 04-podmienky-a-cykly.bas
│   ├── 05-pouzivatelska-funkcia.bas
│   └── 06-automatizacia-prehladu.bas
└── projekty/
    ├── faktura/zadanie.md
    ├── dochadzka/zadanie.md
    ├── rozpocet/zadanie.md
    ├── prehlad-predaja/zadanie.md
    └── absencie-zamestnancov/zadanie.md
```

### Dáta

1. [`zamestnanci.csv`](data/zamestnanci.csv)
2. [`predaj.csv`](data/predaj.csv)
3. [`faktury.csv`](data/faktury.csv)
4. [`dochadzka.csv`](data/dochadzka.csv)
5. [`rozpocet.csv`](data/rozpocet.csv)
6. [`produkty.csv`](data/produkty.csv)
7. [`vyhladavanie-produktov.csv`](data/vyhladavanie-produktov.csv)

### Cvičenia

1. [`01-zaklady-excelu.md`](cvicenia/01-zaklady-excelu.md)
2. [`02-tabulky-a-formatovanie.md`](cvicenia/02-tabulky-a-formatovanie.md)
3. [`03-vzorce-a-funkcie.md`](cvicenia/03-vzorce-a-funkcie.md)
4. [`04-filtrovanie-a-triedenie.md`](cvicenia/04-filtrovanie-a-triedenie.md)
5. [`05-grafy-a-tlac.md`](cvicenia/05-grafy-a-tlac.md)
6. [`06-vyhladavanie-a-analyza.md`](cvicenia/06-vyhladavanie-a-analyza.md)
7. [`07-kontingencne-tabulky.md`](cvicenia/07-kontingencne-tabulky.md)
8. [`08-vba-automatizacia.md`](cvicenia/08-vba-automatizacia.md)

### Riešenia

1. [`excel-01-zaciatocnik.xlsx`](riesenia/excel-01-zaciatocnik.xlsx)
2. [`excel-02-mierne-pokrocily.xlsx`](riesenia/excel-02-mierne-pokrocily.xlsx)
3. [`excel-03-pokrocily.xlsx`](riesenia/excel-03-pokrocily.xlsx)
4. [`excel-04-expert-vba.xlsm`](riesenia/excel-04-expert-vba.xlsm)

### VBA ukážky

1. [`01-prve-makro.bas`](vba-ukazky/01-prve-makro.bas)
2. [`02-rozsahy-a-bunky.bas`](vba-ukazky/02-rozsahy-a-bunky.bas)
3. [`03-harky.bas`](vba-ukazky/03-harky.bas)
4. [`04-podmienky-a-cykly.bas`](vba-ukazky/04-podmienky-a-cykly.bas)
5. [`05-pouzivatelska-funkcia.bas`](vba-ukazky/05-pouzivatelska-funkcia.bas)
6. [`06-automatizacia-prehladu.bas`](vba-ukazky/06-automatizacia-prehladu.bas)

<a id="excel-i-zaciatocnik"></a>
# 🟢 Excel I. Začiatočník

<a id="excel-i-01"></a>
## 1. Navigácia a práca s bunkami

```excel
=B2*C2
```

<a id="excel-i-02"></a>
## 2. Relatívne a absolútne odkazy

```excel
=D2*$F$1
```

<a id="excel-i-03"></a>
## 3. Základné funkcie

```excel
=SUM(D2:D100)
=AVERAGE(D2:D100)
=MIN(D2:D100)
=MAX(D2:D100)
=COUNT(D2:D100)
```

<a id="excel-i-04"></a>
## 4. Excelová tabuľka

Použijeme `Ctrl + T` a štruktúrovaný odkaz `=[@Cena]*[@Mnozstvo]`.

<a id="excel-i-05"></a>
## 5. Odstránenie duplicít

Použijeme `Dáta > Odstrániť duplicity`.

<a id="excel-i-06"></a>
## 6. Filtrovanie a triedenie

Precvičíme textové, číselné a dátumové filtre a viacúrovňové triedenie.

<a id="excel-i-07"></a>
## 7. Formáty čísel a percentá

Precvičíme menu EUR, percentá, dátumy a čas.

<a id="excel-i-08"></a>
## 8. Podmienené formátovanie

Zvýrazníme hodnoty podľa pravidiel a hraníc.

<a id="excel-i-09"></a>
## 9. Graf predaja

Vytvoríme stĺpcový graf mesačných tržieb.

<a id="excel-i-10"></a>
## 10. Projekt – faktúra

Zadanie: [`projekty/faktura/zadanie.md`](projekty/faktura/zadanie.md)

<a id="excel-i-11"></a>
## 11. Projekt – evidencia dochádzky

Zadanie: [`projekty/dochadzka/zadanie.md`](projekty/dochadzka/zadanie.md)

<a id="excel-ii-mierne-pokrocily"></a>
# 🟡 Excel II. Mierne Pokročilý

<a id="excel-ii-12"></a>
## 12. Číselné a dátumové postupnosti

Precvičíme automatické vypĺňanie čísel, dátumov a vlastných radov.

<a id="excel-ii-13"></a>
## 13. Vlastné zoznamy

Vytvoríme vlastné poradie hodnôt pre opakované používanie.

<a id="excel-ii-14"></a>
## 14. Efektívne vkladanie dát

Precvičíme `Ctrl + Enter`, kopírovanie a prácu s výberom.

<a id="excel-ii-15"></a>
## 15. Pokročilá selekcia a Prejsť na

Použijeme `Ctrl + G`, `F5` a Prejsť na špeciálne.

<a id="excel-ii-16"></a>
## 16. Dátumy a časy

```excel
=TODAY()
=NOW()
=YEAR(A2)
=MONTH(A2)
=DAY(A2)
```

<a id="excel-ii-17"></a>
## 17. Automatické filtre a zástupné znaky

Precvičíme `*` a `?` pri textových filtroch.

<a id="excel-ii-18"></a>
## 18. Vlastné zoradenie

Zoradíme dáta podľa viacerých stĺpcov a vlastného poradia.

<a id="excel-ii-19"></a>
## 19. Prilepiť špeciálne

Precvičíme hodnoty, vzorce, formáty a transpozíciu.

<a id="excel-ii-20"></a>
## 20. Pokročilé grafy

Porovnáme stĺpcový, čiarový a koláčový graf.

<a id="excel-ii-21"></a>
## 21. Tlač a zlomy strán

Nastavíme orientáciu, tlačovú oblasť a zlomy strán.

<a id="excel-iii-pokrocily"></a>
# 🔵 Excel III. Pokročilý

<a id="excel-iii-22"></a>
## 22. Limity pracovného hárka

Excel podporuje 1 048 576 riadkov a 16 384 stĺpcov.

<a id="excel-iii-23"></a>
## 23. Volatilné funkcie

Vysvetlíme dopad funkcií `NOW`, `TODAY`, `RAND`, `OFFSET` a `INDIRECT` na prepočet.

<a id="excel-iii-24"></a>
## 24. VLOOKUP

```excel
=VLOOKUP(A2,$H$2:$J$100,3,FALSE)
```

<a id="excel-iii-25"></a>
## 25. HLOOKUP

Precvičíme horizontálne vyhľadávanie.

<a id="excel-iii-26"></a>
## 26. MATCH a INDEX

```excel
=INDEX(C2:C100,MATCH(A2,A2:A100,0))
```

<a id="excel-iii-27"></a>
## 27. XLOOKUP

XLOOKUP uvádzame ako modernú alternatívu pre aktuálne verzie Excelu.

<a id="excel-iii-28"></a>
## 28. CHOOSE

```excel
=CHOOSE(B2,"Nízka","Stredná","Vysoká","Kritická")
```

<a id="excel-iii-29"></a>
## 29. Hľadanie riešenia

Použijeme `Dáta > Analýza hypotéz > Hľadanie riešenia`.

<a id="excel-iii-30"></a>
## 30. SUMPRODUCT

```excel
=SUMPRODUCT(B2:B100,C2:C100)
```

<a id="excel-iii-31"></a>
## 31. Ošetrenie chýb

```excel
=IFERROR(VLOOKUP(A2,$H$2:$J$100,3,FALSE),"Nenájdené")
```

<a id="excel-iii-32"></a>
## 32. Projekt – absencie zamestnancov

Zadanie: [`projekty/absencie-zamestnancov/zadanie.md`](projekty/absencie-zamestnancov/zadanie.md)

<a id="excel-iii-33"></a>
## 33. Projekt – firemný rozpočet

Zadanie: [`projekty/rozpocet/zadanie.md`](projekty/rozpocet/zadanie.md)

<a id="excel-iii-34"></a>
## 34. Kontingenčná tabuľka

Zdroj: [`data/predaj.csv`](data/predaj.csv)

<a id="excel-iii-35"></a>
## 35. Unikátne hodnoty

```excel
=UNIQUE(A2:A1000)
```

<a id="excel-iii-36"></a>
## 36. Rýchla analýza

Precvičíme nástroj Rýchla analýza.

<a id="excel-iii-37"></a>
## 37. Import CSV

Sledujeme oddeľovač, kódovanie, desatinný oddeľovač, dátumový formát a dátové typy.

<a id="excel-iii-38"></a>
## 38. Informačné funkcie

```excel
=INFO("osversion")
=SHEET()
=SHEETS()
=ISTEXT(A2)
=TYPE(A2)
=CELL("address",A2)
```

<a id="excel-iv-expert-vba"></a>
# 🔴 Excel IV. Expert – Programovanie vo VBA

<a id="excel-iv-39"></a>
## 39. Prvé makro

```vb
Sub AhojExcel()
    MsgBox "Ahoj Excel!"
End Sub
```

Zdroj: [`01-prve-makro.bas`](vba-ukazky/01-prve-makro.bas)

<a id="excel-iv-40"></a>
## 40. Range

Precvičíme explicitnú prácu s objektom `Range`.

<a id="excel-iv-41"></a>
## 41. Cells

Precvičíme adresovanie `Cells(riadok, stlpec)`.

<a id="excel-iv-42"></a>
## 42. Práca s hárkami

Precvičíme `Worksheets.Add`, pomenovanie a počítanie hárkov.

<a id="excel-iv-43"></a>
## 43. Automatická šírka a výška

Použijeme `Columns.AutoFit` a `Rows.AutoFit`.

<a id="excel-iv-44"></a>
## 44. Premenné a Option Explicit

Použijeme explicitnú deklaráciu premenných.

<a id="excel-iv-45"></a>
## 45. Podmienka If

Precvičíme rozhodovanie podľa hodnoty bunky.

<a id="excel-iv-46"></a>
## 46. Cyklus For

Precvičíme opakovanie nad číslovanými riadkami.

<a id="excel-iv-47"></a>
## 47. For Step

Precvičíme zmenu kroku cyklu.

<a id="excel-iv-48"></a>
## 48. For Each

Precvičíme prechod cez bunky rozsahu.

<a id="excel-iv-49"></a>
## 49. Do While

Precvičíme cyklus s podmienkou na začiatku.

<a id="excel-iv-50"></a>
## 50. Do Until a Exit Do

Precvičíme ukončenie cyklu pri splnení podmienky.

<a id="excel-iv-51"></a>
## 51. MsgBox a InputBox

Precvičíme komunikáciu s používateľom.

<a id="excel-iv-52"></a>
## 52. Nahradenie textu

Použijeme metódu `Replace`.

<a id="excel-iv-53"></a>
## 53. Kopírovanie a prilepenie

Pri kopírovaní hodnôt preferujeme priamy zápis medzi rozsahmi.

<a id="excel-iv-54"></a>
## 54. Vymazanie obsahu a riadkov

Použijeme `ClearContents`, `Clear` a `Delete`.

<a id="excel-iv-55"></a>
## 55. Používateľská funkcia UDF

Vytvoríme vlastnú funkciu na výpočet ceny s DPH.

<a id="excel-iv-56"></a>
## 56. Automatické vytvorenie prehľadu

Zdroj: [`06-automatizacia-prehladu.bas`](vba-ukazky/06-automatizacia-prehladu.bas)

<a id="excel-iv-57"></a>
## 57. Optimalizácia VBA kódu

Použijeme `ScreenUpdating`, `EnableEvents` a `Calculation`.

<a id="excel-iv-58"></a>
## 58. Ovládacie tlačidlo

Priradíme makro tlačidlu formulára.

<a id="prakticke-projekty"></a>
# 🧪 Praktické projekty

1. [Faktúra](projekty/faktura/zadanie.md)
2. [Dochádzka](projekty/dochadzka/zadanie.md)
3. [Firemný rozpočet](projekty/rozpocet/zadanie.md)
4. [Prehľad predaja](projekty/prehlad-predaja/zadanie.md)
5. [Absencie zamestnancov](projekty/absencie-zamestnancov/zadanie.md)

# 🔒 Bezpečnostné poznámky

1. Nespúšťame makrá z nedôveryhodných zdrojov
2. Pred spustením makra skontrolujeme zdrojový kód
3. Produkčné dáta nahrádzame anonymizovanými tréningovými dátami
4. Do repozitára neukladáme heslá, API tokeny ani dôverné firemné exporty
5. Makrá zapisujúce alebo odstraňujúce dáta testujeme na kópii zošita
