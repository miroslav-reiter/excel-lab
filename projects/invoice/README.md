# Projekt – Faktúra

## Cieľ

Vytvoríme praktickú faktúru s automatickými výpočtami a tlačovým výstupom.

## Zdrojové dáta

Použijeme `../../data/invoices.csv` a `../../data/products.csv`.

## Požiadavky

1. Vytvoríme hlavičku faktúry s číslom, dátumom vystavenia a dátumom splatnosti
2. Pridáme odberateľa a základné identifikačné údaje
3. Vytvoríme položky `Kód`, `Produkt`, `Množstvo`, `Cena bez DPH`, `DPH`, `Cena s DPH`
4. Názov a cenu produktu načítame podľa produktového kódu
5. Vypočítame základ dane, DPH a celkovú sumu
6. Použijeme absolútny odkaz na sadzbu DPH alebo sadzbu načítame z produktových dát
7. Nastavíme formát EUR
8. Pripravíme tlač na jednu stranu A4
9. Overíme správanie pri neexistujúcom produktovom kóde
10. Uložíme výsledok ako `.xlsx`
