# 🧪 Online Kurzy Microsoft Excel (excel-lab)

Praktický repozitár na precvičovanie programu **Microsoft Excel** od úplných základov až po pokročilú analýzu dát, kontingenčné tabuľky, automatizáciu a programovanie vo **VBA (Visual Basic for Applications)**.

Repozitár nadväzuje na štyri online kurzy VITA Academy:

1. [Microsoft Excel I. Začiatočník](https://www.vita.sk/online-kurz-microsoft-excel-i-zaciatocnik/)
2. [Microsoft Excel II. Mierne Pokročilý](https://www.vita.sk/online-kurz-microsoft-excel-ii-mierne-pokrocily/)
3. [Microsoft Excel III. Pokročilý](https://www.vita.sk/online-kurz-microsoft-excel-iii-pokrocily/)
4. [Microsoft Excel IV. Expert: Programovanie vo VBA](https://www.vita.sk/online-kurz-microsoft-excel-iv-expert-programovanie-vo-vba/)

Cieľom repozitára je precvičiť prácu s bunkami, tabuľkami, formátovaním, vzorcami, funkciami, filtrami, grafmi, veľkými dátami, kontingenčnými tabuľkami, importom dát, makrami a objektovým modelom Microsoft Excelu na realistických firemných príkladoch.

> **Poznámka k vzorcom:** technické názvy funkcií ako `SUM`, `AVERAGE`, `VLOOKUP`, `INDEX`, `MATCH` alebo `SUMPRODUCT` uvádzame v medzinárodnom anglickom tvare. Podľa jazykového a regionálneho nastavenia Excelu môže byť oddeľovač argumentov `,` alebo `;`.

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
   6. [XLOOKUP ako moderná alternatíva](#excel-iii-27)
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
7. [Excel ťahák](#excel-tahak)
8. [Časté chyby](#caste-chyby)
9. [Bezpečnostné poznámky](#bezpecnost)

<a id="prakticke-subory"></a>
## 🗂️ Praktické súbory repozitára

```text
excel-lab/
├── README.md
├── dáta/
│   ├── zamestnanci.csv
│   ├── predaj.csv
│   ├── faktúry.csv
│   ├── dochádzka.csv
│   ├── rozpočet.csv
│   ├── produkty.csv
│   └── vyhľadávanie-produktov.csv
├── cvičenia/
│   ├── 01-základy-excelu.md
│   ├── 02-tabuľky-a-formátovanie.md
│   ├── 03-vzorce-a-funkcie.md
│   ├── 04-filtrovanie-a-triedenie.md
│   ├── 05-grafy-a-tlač.md
│   ├── 06-vyhľadávanie-a-analýza.md
│   ├── 07-kontingenčné-tabuľky.md
│   └── 08-vba-automatizácia.md
├── riešenia/
│   ├── excel-01-začiatočník.xlsx
│   ├── excel-02-mierne-pokročilý.xlsx
│   ├── excel-03-pokročilý.xlsx
│   └── excel-04-expert-vba.xlsm
├── vba-ukážky/
│   ├── 01-prvé-makro.bas
│   ├── 02-rozsahy-a-bunky.bas
│   ├── 03-hárky.bas
│   ├── 04-podmienky-a-cykly.bas
│   ├── 05-používateľská-funkcia.bas
│   └── 06-automatizácia-prehľadu.bas
└── projekty/
    ├── faktúra/zadanie.md
    ├── dochádzka/zadanie.md
    ├── rozpočet/zadanie.md
    ├── prehľad-predaja/zadanie.md
    └── absencie-zamestnancov/zadanie.md
```

### Dáta

1. [`zamestnanci.csv`](dáta/zamestnanci.csv) obsahuje zamestnancov, oddelenia, pozície, platy a dátumy nástupu
2. [`predaj.csv`](dáta/predaj.csv) obsahuje obchodné transakcie vhodné na filtrovanie, grafy a kontingenčné tabuľky
3. [`faktúry.csv`](dáta/faktúry.csv) obsahuje faktúry, dátumy splatnosti a stavy úhrad
4. [`dochádzka.csv`](dáta/dochádzka.csv) obsahuje príchody, odchody, prestávky a formu práce
5. [`rozpočet.csv`](dáta/rozpočet.csv) obsahuje plánované a skutočné náklady podľa oddelení
6. [`produkty.csv`](dáta/produkty.csv) obsahuje produktový katalóg a ceny
7. [`vyhľadávanie-produktov.csv`](dáta/vyhľadávanie-produktov.csv) slúži na precvičovanie vyhľadávacích funkcií

### Cvičenia

1. [`01-základy-excelu.md`](cvičenia/01-základy-excelu.md)
2. [`02-tabuľky-a-formátovanie.md`](cvičenia/02-tabuľky-a-formátovanie.md)
3. [`03-vzorce-a-funkcie.md`](cvičenia/03-vzorce-a-funkcie.md)
4. [`04-filtrovanie-a-triedenie.md`](cvičenia/04-filtrovanie-a-triedenie.md)
5. [`05-grafy-a-tlač.md`](cvičenia/05-grafy-a-tlač.md)
6. [`06-vyhľadávanie-a-analýza.md`](cvičenia/06-vyhľadávanie-a-analýza.md)
7. [`07-kontingenčné-tabuľky.md`](cvičenia/07-kontingenčné-tabuľky.md)
8. [`08-vba-automatizácia.md`](cvičenia/08-vba-automatizácia.md)

### Riešenia

1. [`excel-01-začiatočník.xlsx`](riešenia/excel-01-začiatočník.xlsx)
2. [`excel-02-mierne-pokročilý.xlsx`](riešenia/excel-02-mierne-pokročilý.xlsx)
3. [`excel-03-pokročilý.xlsx`](riešenia/excel-03-pokročilý.xlsx)
4. [`excel-04-expert-vba.xlsm`](riešenia/excel-04-expert-vba.xlsm)

### VBA ukážky

1. [`01-prvé-makro.bas`](vba-ukážky/01-prvé-makro.bas)
2. [`02-rozsahy-a-bunky.bas`](vba-ukážky/02-rozsahy-a-bunky.bas)
3. [`03-hárky.bas`](vba-ukážky/03-hárky.bas)
4. [`04-podmienky-a-cykly.bas`](vba-ukážky/04-podmienky-a-cykly.bas)
5. [`05-používateľská-funkcia.bas`](vba-ukážky/05-používateľská-funkcia.bas)
6. [`06-automatizácia-prehľadu.bas`](vba-ukážky/06-automatizácia-prehľadu.bas)

<a id="excel-i-zaciatocnik"></a>
# 🟢 Excel I. Začiatočník

Kurz I. pokrýva používateľské prostredie, pracovné zošity a hárky, prácu s bunkami, tabuľkami, filtrovanie, triedenie, formátovanie, základné vzorce, funkcie, grafy a praktické firemné dokumenty.

<a id="excel-i-01"></a>
## 1. Navigácia a práca s bunkami

Príklad tabuľky:

| Produkt | Cena | Množstvo | Tržba |
|---|---:|---:|---:|
| Notebook | 850 | 2 | |
| Monitor | 240 | 4 | |
| Myš | 35 | 10 | |

Do bunky `D2` zadáme:

```excel
=B2*C2
```

<a id="excel-i-02"></a>
## 2. Relatívne a absolútne odkazy

Ak je sadzba DPH v bunke `F1`, použijeme:

```excel
=D2*$F$1
```

| Typ odkazu | Príklad |
|---|---|
| Relatívny | `D2` |
| Absolútny | `$D$2` |
| Fixovaný riadok | `D$2` |
| Fixovaný stĺpec | `$D2` |

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

Označíme dáta a použijeme `Ctrl + T`. Tabuľku môžeme pomenovať napríklad `tblPredaj`.

```excel
=[@Cena]*[@Množstvo]
```

<a id="excel-i-05"></a>
## 5. Odstránenie duplicít

Použijeme `Dáta > Odstrániť duplicity`. Pred zásahom pracujeme s kópiou zdrojových dát.

<a id="excel-i-06"></a>
## 6. Filtrovanie a triedenie

Na súbore [`predaj.csv`](dáta/predaj.csv) precvičíme:

1. filter podľa regiónu
2. filter podľa obchodníka
3. číselný filter podľa tržby
4. zoradenie podľa tržby zostupne
5. viacúrovňové zoradenie podľa regiónu, obchodníka a tržby

<a id="excel-i-07"></a>
## 7. Formáty čísel a percentá

```text
1250,5       → 1 250,50 €
0,235        → 23,50 %
2026-09-12   → 12. 9. 2026
0,5          → 12:00
```

<a id="excel-i-08"></a>
## 8. Podmienené formátovanie

Príklad vlastného pravidla:

```excel
=$D2>2000
```

<a id="excel-i-09"></a>
## 9. Graf predaja

Zo súhrnu mesačných tržieb vytvoríme stĺpcový graf na porovnanie a čiarový graf na sledovanie trendu.

<a id="excel-i-10"></a>
## 10. Projekt – faktúra

Projektové zadanie: [`projekty/faktúra/zadanie.md`](projekty/faktúra/zadanie.md)

<a id="excel-i-11"></a>
## 11. Projekt – evidencia dochádzky

Projektové zadanie: [`projekty/dochádzka/zadanie.md`](projekty/dochádzka/zadanie.md)

<a id="excel-ii-mierne-pokrocily"></a>
# 🟡 Excel II. Mierne Pokročilý

Kurz II. rozvíja efektívnu prácu s prostredím, postupnosťami, vkladaním dát, selekciami, tabuľkami, vlastným triedením, filtrami, grafmi a tlačou.

<a id="excel-ii-12"></a>
## 12. Číselné a dátumové postupnosti

Precvičíme postupnosti čísel, dátumov, pracovných dní a vlastných zoznamov.

```excel
=SEQUENCE(100)
```

<a id="excel-ii-13"></a>
## 13. Vlastné zoznamy

Príklad vlastného poradia:

```text
Junior
Medior
Senior
Vedúci
Manažér
```

<a id="excel-ii-14"></a>
## 14. Efektívne vkladanie dát

Precvičíme `Enter`, `Shift + Enter`, `Tab`, `Shift + Tab`, `Ctrl + Enter` a `Alt + Enter`.

<a id="excel-ii-15"></a>
## 15. Pokročilá selekcia a Prejsť na

Použijeme `Ctrl + G` alebo `F5` a príkaz Prejsť na špeciálne pre prázdne bunky, vzorce, konštanty a viditeľné bunky.

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

```text
*    ľubovoľný počet znakov
?    jeden ľubovoľný znak
~    potlačenie špeciálneho významu znaku
```

<a id="excel-ii-18"></a>
## 18. Vlastné zoradenie

Vytvoríme napríklad vlastné poradie priorít `Kritická`, `Vysoká`, `Stredná`, `Nízka`.

<a id="excel-ii-19"></a>
## 19. Prilepiť špeciálne

Precvičíme hodnoty, vzorce, formáty, transpozíciu a matematické operácie pri prilepení.

<a id="excel-ii-20"></a>
## 20. Pokročilé grafy

Porovnáme stĺpcový, čiarový a koláčový graf a zvolíme typ podľa analytického cieľa.

<a id="excel-ii-21"></a>
## 21. Tlač a zlomy strán

Precvičíme oblasť tlače, orientáciu, okraje, mierku, náhľad zlomov strán, hlavičku a pätu.

<a id="excel-iii-pokrocily"></a>
# 🟠 Excel III. Pokročilý

Kurz III. sa zameriava na veľké dátové súbory, vyhľadávacie funkcie, Hľadanie riešenia, `SUMPRODUCT`, rozpočty, personálne dáta, kontingenčné tabuľky, importy a informačné funkcie.

<a id="excel-iii-22"></a>
## 22. Limity pracovného hárka

```text
1 048 576 riadkov
16 384 stĺpcov
posledný stĺpec XFD
```

<a id="excel-iii-23"></a>
## 23. Volatilné funkcie

Medzi funkcie, ktoré môžu spôsobovať častejšie prepočítavanie zošita, patria napríklad `NOW`, `TODAY`, `RAND`, `RANDBETWEEN`, `OFFSET` a `INDIRECT`.

<a id="excel-iii-24"></a>
## 24. VLOOKUP

```excel
=VLOOKUP(A2,$H$2:$J$100,3,FALSE)
```

<a id="excel-iii-25"></a>
## 25. HLOOKUP

```excel
=HLOOKUP(B1,$B$10:$M$12,3,FALSE)
```

<a id="excel-iii-26"></a>
## 26. MATCH a INDEX

```excel
=MATCH(A2,$H$2:$H$100,0)
=INDEX($J$2:$J$100,MATCH(A2,$H$2:$H$100,0))
```

<a id="excel-iii-27"></a>
## 27. XLOOKUP ako moderná alternatíva

```excel
=XLOOKUP(A2,$H$2:$H$100,$J$2:$J$100,"Nenájdené")
```

<a id="excel-iii-28"></a>
## 28. CHOOSE

```excel
=CHOOSE(B2,"Nízka","Stredná","Vysoká","Kritická")
```

<a id="excel-iii-29"></a>
## 29. Hľadanie riešenia

Použijeme `Dáta > Analýza hypotéz > Hľadanie riešenia` a vypočítame vstupnú hodnotu potrebnú na dosiahnutie cieľového výsledku.

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

Projektové zadanie: [`projekty/absencie-zamestnancov/zadanie.md`](projekty/absencie-zamestnancov/zadanie.md)

<a id="excel-iii-33"></a>
## 33. Projekt – firemný rozpočet

Projektové zadanie: [`projekty/rozpočet/zadanie.md`](projekty/rozpočet/zadanie.md)

<a id="excel-iii-34"></a>
## 34. Kontingenčná tabuľka

Zo súboru [`predaj.csv`](dáta/predaj.csv) vytvoríme prehľad podľa regiónu, mesiaca a tržby.

<a id="excel-iii-35"></a>
## 35. Unikátne hodnoty

```excel
=UNIQUE(A2:A1000)
```

<a id="excel-iii-36"></a>
## 36. Rýchla analýza

Označíme rozsah a použijeme nástroj Rýchla analýza pre formátovanie, grafy, súčty, tabuľky a minigrafy.

<a id="excel-iii-37"></a>
## 37. Import CSV

Pri importe sledujeme oddeľovač, kódovanie, desatinný oddeľovač, dátumový formát, dátové typy a vedúce nuly.

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

Kurz IV. sa zameriava na automatizáciu, záznam makier, objektový model Excelu, objekty `Range`, `Cells` a `Worksheets`, používateľské funkcie, premenné, podmienky, cykly, ovládacie prvky a optimalizáciu.

<a id="excel-iv-39"></a>
## 39. Prvé makro

```vb
Option Explicit

Sub AhojExcel()
    MsgBox "Ahoj Excel!"
End Sub
```

Zdroj: [`01-prvé-makro.bas`](vba-ukážky/01-prvé-makro.bas)

<a id="excel-iv-40"></a>
## 40. Range

```vb
Sub ZapisDoRozsahu()
    Worksheets("Dáta").Range("A1").Value = "VITA Academy"
End Sub
```

<a id="excel-iv-41"></a>
## 41. Cells

```vb
Sub ZapisPomocouCells()
    Worksheets("Dáta").Cells(1, 1).Value = "A1"
    Worksheets("Dáta").Cells(2, 3).Value = "C2"
End Sub
```

<a id="excel-iv-42"></a>
## 42. Práca s hárkami

```vb
Sub VytvorHarokPrehlad()
    Dim harok As Worksheet
    Set harok = Worksheets.Add
    harok.Name = "Prehľad"
End Sub
```

<a id="excel-iv-43"></a>
## 43. Automatická šírka a výška

```vb
With Worksheets("Prehľad")
    .Columns.AutoFit
    .Rows.AutoFit
End With
```

<a id="excel-iv-44"></a>
## 44. Premenné a Option Explicit

`Option Explicit` vyžaduje deklaráciu premenných a pomáha odhaliť preklepy v názvoch.

<a id="excel-iv-45"></a>
## 45. Podmienka If

```vb
If trzba >= 10000 Then
    Range("B1").Value = "Cieľ splnený"
Else
    Range("B1").Value = "Cieľ nesplnený"
End If
```

<a id="excel-iv-46"></a>
## 46. Cyklus For

```vb
For i = 2 To 101
    Cells(i, 1).Value = i - 1
Next i
```

<a id="excel-iv-47"></a>
## 47. For Step

```vb
For i = 2 To 100 Step 2
    Cells(i, 1).Value = "Párny riadok"
Next i
```

<a id="excel-iv-48"></a>
## 48. For Each

```vb
For Each bunka In Range("D2:D100")
    If IsNumeric(bunka.Value) And bunka.Value < 0 Then
        bunka.Font.Bold = True
    End If
Next bunka
```

<a id="excel-iv-49"></a>
## 49. Do While

```vb
Do While Cells(i, 1).Value <> ""
    Cells(i, 2).Value = UCase(Cells(i, 1).Value)
    i = i + 1
Loop
```

<a id="excel-iv-50"></a>
## 50. Do Until a Exit Do

```vb
Do Until Cells(i, 1).Value = ""
    If Cells(i, 1).Value = "STOP" Then Exit Do
    i = i + 1
Loop
```

<a id="excel-iv-51"></a>
## 51. MsgBox a InputBox

```vb
meno = InputBox("Zadajte meno používateľa:")
If meno <> "" Then MsgBox "Ahoj " & meno
```

<a id="excel-iv-52"></a>
## 52. Nahradenie textu

Použijeme metódu `Replace` nad explicitne zadaným hárkom a rozsahom.

<a id="excel-iv-53"></a>
## 53. Kopírovanie a prilepenie

Pri kopírovaní iba hodnôt preferujeme priamy zápis medzi rozsahmi namiesto zbytočného používania schránky.

<a id="excel-iv-54"></a>
## 54. Vymazanie obsahu a riadkov

```vb
Range("A2:D100").ClearContents
Rows(10).Delete
```

<a id="excel-iv-55"></a>
## 55. Používateľská funkcia UDF

```vb
Function CenaSDPH(cenaBezDPH As Double, sadzbaDPH As Double) As Double
    CenaSDPH = cenaBezDPH * (1 + sadzbaDPH)
End Function
```

<a id="excel-iv-56"></a>
## 56. Automatické vytvorenie prehľadu

Zdrojový modul: [`06-automatizácia-prehľadu.bas`](vba-ukážky/06-automatizácia-prehľadu.bas)

Makro pracuje s hárkom `Dáta`, vytvorí alebo vyčistí hárok `Prehľad`, skopíruje aktuálny rozsah a prispôsobí šírku stĺpcov.

<a id="excel-iv-57"></a>
## 57. Optimalizácia VBA kódu

Pri väčších operáciách môžeme dočasne vypnúť prekresľovanie obrazovky, udalosti a automatický prepočet. Na konci ich vždy obnovíme aj pri vzniku chyby.

<a id="excel-iv-58"></a>
## 58. Ovládacie tlačidlo

Vložíme tlačidlo formulára cez kartu Vývojár a priradíme mu makro na vytvorenie prehľadu.

<a id="excel-tahak"></a>
# 🧠 Excel ťahák

## Najčastejšie klávesové skratky

| Akcia | Skratka |
|---|---|
| Uložiť zošit | `Ctrl + S` |
| Kopírovať | `Ctrl + C` |
| Vystrihnúť | `Ctrl + X` |
| Prilepiť | `Ctrl + V` |
| Späť | `Ctrl + Z` |
| Upraviť bunku | `F2` |
| Prepnúť typ odkazu vo vzorci | `F4` |
| Prejsť na | `Ctrl + G` alebo `F5` |
| Vytvoriť tabuľku | `Ctrl + T` |
| Zapnúť alebo vypnúť filter | `Ctrl + Shift + L` |
| Nový riadok v bunke | `Alt + Enter` |

## Základné funkcie

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
| Dnešný dátum | `TODAY` |
| Dátum a čas | `NOW` |

<a id="prakticke-projekty"></a>
# 🧪 Praktické projekty

1. [Faktúra](projekty/faktúra/zadanie.md) – DPH, medzisúčty, formátovanie a tlač
2. [Dochádzka](projekty/dochádzka/zadanie.md) – dátumy, časy, prestávky a odpracované hodiny
3. [Firemný rozpočet](projekty/rozpočet/zadanie.md) – rozpočet, skutočnosť, odchýlka a Hľadanie riešenia
4. [Prehľad predaja](projekty/prehľad-predaja/zadanie.md) – filtre, vyhľadávacie funkcie, grafy a kontingenčná tabuľka
5. [Absencie zamestnancov](projekty/absencie-zamestnancov/zadanie.md) – personálne dáta, dátumy, súhrny a kontingenčné tabuľky

<a id="caste-chyby"></a>
# ⚠️ Časté chyby

## Číslo uložené ako text

Overíme:

```excel
=ISNUMBER(A2)
```

## Nesprávne fixovaný rozsah

Menej vhodné:

```excel
=VLOOKUP(A2,H2:J100,3,FALSE)
```

Vhodné pri kopírovaní vzorca:

```excel
=VLOOKUP(A2,$H$2:$J$100,3,FALSE)
```

## Zlúčené bunky v dátovej tabuľke

Zlúčené bunky v dátových oblastiach komplikujú filtrovanie, triedenie, importy aj automatizáciu.

## Hodnoty zapísané priamo vo vzorcoch

Menej vhodné:

```excel
=A2*1.23
```

Lepšie:

```excel
=A2*(1+$F$1)
```

## VBA zošit uložený ako XLSX

Zošit s VBA ukladáme ako `Excel Macro-Enabled Workbook (*.xlsm)`.

<a id="bezpecnost"></a>
# 🔒 Bezpečnostné poznámky

1. nespúšťame makrá z neznámych alebo nedôveryhodných zdrojov
2. pred spustením makra skontrolujeme zdrojový kód
3. produkčné dáta nahrádzame anonymizovanými tréningovými dátami
4. do repozitára neukladáme heslá, API tokeny, osobné údaje ani dôverné firemné exporty
5. makrá zapisujúce alebo odstraňujúce dáta testujeme najprv na kópii zošita
6. pri automatizácii používame explicitné názvy zošitov, hárkov a rozsahov

# 📚 Užitočné odkazy

1. [Microsoft Excel I. Začiatočník](https://www.vita.sk/online-kurz-microsoft-excel-i-zaciatocnik/)
2. [Microsoft Excel II. Mierne Pokročilý](https://www.vita.sk/online-kurz-microsoft-excel-ii-mierne-pokrocily/)
3. [Microsoft Excel III. Pokročilý](https://www.vita.sk/online-kurz-microsoft-excel-iii-pokrocily/)
4. [Microsoft Excel IV. Expert: Programovanie vo VBA](https://www.vita.sk/online-kurz-microsoft-excel-iv-expert-programovanie-vo-vba/)
5. [Excel specifications and limits](https://support.microsoft.com/en-us/excel/excel-specifications-and-limits)
6. [Klávesové skratky v Exceli](https://support.microsoft.com/sk-sk/accessibility/excel/keyboard-shortcuts-in-excel)
7. [VBA Range object](https://learn.microsoft.com/sk-sk/office/vba/api/excel.range(object))

# ✅ Cieľ repozitára

Po prejdení príkladov by sme mali vedieť:

1. samostatne vytvárať a upravovať Excel zošity
2. efektívne pracovať s bunkami, rozsahmi a tabuľkami
3. filtrovať, triediť a formátovať dáta
4. používať vzorce, funkcie a rôzne typy odkazov
5. vytvárať grafy a tlačové prehľady
6. analyzovať väčšie dátové súbory
7. používať vyhľadávacie funkcie a `SUMPRODUCT`
8. vytvárať a analyzovať kontingenčné tabuľky
9. importovať CSV a ďalšie dátové formáty
10. nahrávať a upravovať makrá
11. programovať praktické VBA procedúry
12. automatizovať opakované úlohy v Microsoft Exceli
