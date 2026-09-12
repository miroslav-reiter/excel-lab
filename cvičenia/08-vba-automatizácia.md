# Cvičenie 08 – Automatizácia pomocou VBA

## Cieľ

Precvičíme makrá, objektový model Excelu, podmienky, cykly, používateľské funkcie a automatické vytvorenie prehľadu.

## Súbory

Použijeme `riešenia/excel-04-expert-vba.xlsm` a moduly z priečinka `vba-ukážky/`.

## Zadania

1. Otvoríme `excel-04-expert-vba.xlsm` v desktopovej aplikácii Excel
2. Aktivujeme kartu Vývojár
3. Otvoríme editor jazyka VBA pomocou `Alt + F11`
4. Importujeme modul `01-prvé-makro.bas`
5. Spustíme procedúru `AhojExcel`
6. Importujeme `02-rozsahy-a-bunky.bas` a otestujeme prácu s objektmi `Range` a `Cells`
7. Importujeme `03-hárky.bas` a vytvoríme hárok `Prehľad`
8. Importujeme `04-podmienky-a-cykly.bas` a spustíme príklady podmienok a cyklov
9. Importujeme `05-používateľská-funkcia.bas` a použijeme vlastnú funkciu v bunke
10. Importujeme `06-automatizácia-prehľadu.bas` a vygenerujeme prehľad z hárka `Dáta`
11. Priradíme makro `VytvorPrehladPredaja` tlačidlu formulára
12. Uložíme zošit ako `.xlsm`

## Bezpečnosť

1. Makrá spúšťame iba z dôveryhodných zdrojov
2. Pred spustením si prečítame zdrojový kód
3. Automatizáciu testujeme na kópii dát
4. Produkčné heslá a tokeny do VBA kódu nevkladáme

## Kontrola

1. Editor VBA nesmie hlásiť chybu pri kompilácii
2. Prehľad sa musí vytvoriť opakovane bez vytvárania duplicitných hárkov
3. Po skončení makra musia byť obnovené nastavenia `ScreenUpdating`, `EnableEvents` a `Calculation`
