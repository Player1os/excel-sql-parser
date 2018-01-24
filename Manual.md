<link rel="stylesheet" href="C:\\Users\\osama.hassanein\\.vscode\\markdown-pdf.css" />
<style>
</style>

<header class="title">
	<h1 class="main-title">Užívateľský návod pre Excel SQL Parser</h1>
	<h2 class="sub-title">Makro umožňujúce vykonať SQL dotazy nad dátami v Exceli</h2>
</header>

# Ako funguje SQL v Exceli

Excel využíva rovnakú verziu SQL jazyka ako aktuálne nainštalovaná verzia aplikácie *Microsoft Access*. Všetky dáta, nad ktorými budeme volať SQL dotazy sa musia nachádzať v hárkoch súboru `SQLParser.xlsm`.

Ak už s SQL viete pracovať jediné, čo je potrebné si ujasniť, je ako Excel vníma tabuľky. Za tabuľku môžeme považovať ľubovoľný obdĺžnikový výsek buniek v jednom hárku (napr. pole `A3:C25` v hárku `Sheet2`). Excel automaticky pomenováva stĺpce takejto tabuľky podľa hodnôt obsiahnutých v prvom riadku. Zvyšné riadky tvoria jednotlivé dátové položky danej tabuľky.

V SQL dotazoch potrebujeme spôsob na pomenovanie tabuliek. Nasleduje syntax, ktorou označujeme pole buniek ako tabuľku:
`[Názov_hárku$Adresa_ľavej_hornej_bunky:Adresa_pravej_doľnej_bunky]`.

Napríklad ak chceme v dotaze označiť tabuľku, ktorej obsah tvorí pole v hárku `Sheet3`, pričom jeho ľavá horná bunka je `B4` a jeho pravá dolná bunka je `E31`, musíme napísať `[Sheet3$B4:E31]`. V takejto tabuľke bude riadok `B4:E4` obsahovať názvy stĺpcov a riadky v poli `B5:E31` budú obsahovať jednotlivé dátové položky tabuľky.

## Príklady SQL dotazov

Nasledujúci dotaz jednoducho skopíruje obsah poľa na nové miesto.

```sql
SELECT
	*
FROM
	[Main$A2:C20]
;
```

Tento dotaz využíva názvy stĺpcov na vytvorenie zložitejšieho dotazu. Tieto názvy sa musia nachádzať niekde v riadku `A2:C2`, inak je dotaz neplatný.

```sql
SELECT
	cena,
	výkon
FROM
	[Sheet1$A2:C20]
WHERE
	cena > 0
;
```

<div style="page-break-before: always"></div>

Tento dotaz ukazuje, že je možne sa odkazovať na viacero tabuliek naraz a pracovať s ich obsahom.

```sql
SELECT
	tab1.meno
FROM
	[Sheet1$A2:C20] as tab1,
	[Sheet3$B3:D21] as tab2
WHERE
	tab1.cena > tab2.poplatok
	AND tab1.meno > tab2.meno
;
```

## Umiestnenie výsledku

Na výsledok ľubovoľného SQL dotazu sa môžeme pozerať taktiež ako na tabuľku, ktorú vieme zobraziť v Exceli do obdĺžnikového poľa buniek. V našom makre je potrebné pred vykonaním dotazu zvoliť bunku, ktorá bude tvoriť ľavý horný okraj tohto poľa. Je potrebné zvoliť túto bunku tak, aby sme mali dostatok miesta pre výsledok dotazu smerom napravo a dole od nej.
 
# Postup na vykonanie SQL dotazu

1.	Do súboru `SQLParser.xlsm` nakopírujeme všetky dáta, nad ktorými chceme vykonať dotazy (tu môžeme v súbore vytvoriť nové zošity).
1.	Keď sú všetky dáta nakopírované a máme vyhradené miesto pre výsledok dotazu, stlačíme klávesovú skratku `Ctrl + Q`, čím aktivujeme makro.
1.	V zobrazenom okne napíšeme najprv SQL dotaz, ktorý chceme vykonať a potom zvolíme ľavý horný okraj poľa, kam sa zobrazí výsledok dotazu.
1.	Stlačíme tlačidlo **Vykonaj**. A ak je dotaz korektný, makro zobrazí hlášku s výsledkom dotazu.

Ak má dotaz pracovať s viacerými tabuľkami, je výhodnejšie umiestniť obsah každej tabuľky do samostatného hárku, tak aby tabuľka vždy začínala naľavo hore v bunke `A1`.

Z podobného dôvodu je jednoduchšie si vyhradiť jeden hárok (napr. hárok `Main`) pre výsledok dotazu tak, aby sa zobrazil v poli začínajúcom naľavo hore v bunke `A1`.

Súbor `SQLParser.xlsm` je nastavený tak aby sa otváral v **ReadOnly** móde, nie je možné ho preto prepísať a po zavretí sa zmeny neuložia.
