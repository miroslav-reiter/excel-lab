# Cvičenie 08 – VBA automatizácia

## Cieľ

Precvičíme makrá, objektový model Excelu, podmienky, cykly, UDF a automatické vytvorenie reportu.

## Súbory

Použijeme `solutions/excel-04-vba.xlsm` a moduly z priečinka `vba/`.

## Zadania

1. Otvoríme `excel-04-vba.xlsm` v desktopovom Exceli
2. Aktivujeme kartu Developer
3. Otvoríme VBA Editor pomocou `Alt + F11`
4. Importujeme modul `01-first-macro.bas`
5. Spustíme procedúru `HelloExcel`
6. Importujeme `02-range-cells.bas` a otestujeme prácu s `Range` a `Cells`
7. Importujeme `03-worksheets.bas` a vytvoríme hárok `Report`
8. Importujeme `04-conditions-loops.bas` a spustíme príklady podmienok a cyklov
9. Importujeme `05-udf.bas` a použijeme vlastnú funkciu v bunke
10. Importujeme `06-report-automation.bas` a vygenerujeme report z hárka `Data`
11. Priradíme makro `CreateSalesReport` tlačidlu formulára
12. Uložíme zošit ako `.xlsm`

## Bezpečnosť

1. Makrá spúšťame iba z dôveryhodných zdrojov
2. Pred spustením si prečítame zdrojový kód
3. Automatizáciu testujeme na kópii dát
4. Produkčné heslá a tokeny do VBA kódu nevkladáme

## Kontrola

1. VBE nesmie hlásiť compile error
2. Report sa musí vytvoriť opakovane bez vytvárania duplicitných hárkov
3. Po skončení makra musia byť obnovené nastavenia `ScreenUpdating`, `EnableEvents` a `Calculation`
