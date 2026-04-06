# Obsazování

## Stažení a instalace

**Požadavky:** Python 3.10+, macOS

```bash
# Klonuj repozitář
git clone https://github.com/Kaliend/dubbing-casting.git
cd dubbing-casting

# Nainstaluj závislosti
python3 -m pip install -r requirements.txt

# Spusť aplikaci
python3 main.py
```

Desktop aplikace v `PySide6` pro lokální workflow obsazování dabingu:

1. vložit vstupy po dílech jako `POSTAVA / TC / TEXT` nebo `POSTAVA / VSTUPY / REPLIKY`
2. spočítat `VSTUPY` a `REPLIKY`, kde `REPLIKY = ceil(počet_slov / 8)`
3. přiřadit dabéry na úrovni `KOMPLET`
4. uložit pracovní stav do lokálního `.json`
5. vyexportovat hotový `.xlsx` podle šablony `SABLONA_OBSAZENI(01-06).xlsx`

## Architektura

- `obsazovani/core.py`: parsing a agregace vstupů
- `obsazovani/exporter.py`: generování výsledného `.xlsx`
- `obsazovani/app_state.py`: pracovní stav aplikace a orchestrace přepočtu/exportu
- `obsazovani/project_store.py`: JSON persistence a načítání textových vstupů
- `obsazovani/desktop/`: PySide6 UI vrstva bez serveru a bez embedded browseru
- `main.py`: desktop entry point

## Spuštění

Nainstaluj dependency:

```bash
python3 -m pip install -r requirements.txt
```

Spusť desktop aplikaci:

```bash
python3 main.py
```

## Co aplikace umí

- `1 až 6` děl v samostatných tabech, s možností díla přidávat, přejmenovávat a odebírat
- ruční vložení nebo načtení `.txt/.tsv/.csv/.xlsx` pro aktivní dílo
- hromadný import děl z jednoho `.xlsx` workbooku, více souborů nebo celé složky
- import specializovaných zdrojů:
  - Netflix `.xlsx`: první list, sloupce `SOURCE`, `DIALOGUE`, `IN-TIMECODE`
  - obecné `.xlsx`: výběr listu s patternem `POSTAVA / TC / TEXT` nebo `POSTAVA / VSTUPY / REPLIKY`
  - pokud zdroj obsahuje `DABÉR` nebo `POZNÁMKA`, import je převezme do projektového obsazení
  - IYUNO `.doc` a `.docx`: dialogová tabulka `Character / TC / Note / TEXT`
  - klasický `.docx`: řádky `POSTAVA / TC / TEXT`
- průběžný přepočet `KOMPLET` tabulky
- přímou editaci sloupců `Dabér` a `Poznámka`
- souhrn dabérů a stav neobsazených postav
- uložení a načtení projektu přes jednoduchý `.json`
- export do Excel šablony včetně vlastních názvů děl v `KOMPLET` a na aktivních listech děl
- volitelný export `HERCI` po jednotlivých dílech přes checkbox v desktop UI

## Poznámky

- Export zůstává bez externích excelových knihoven, používá existující XML/XLSX vrstvu.
- Šablona i UI momentálně cílí na maximálně 6 děl.
- Import `IYUNO .doc` je teď řešen přes `textutil` na macOS; `.xlsx` a `.docx` import běží čistě v Pythonu.
- `server.py` a složka `web/` zůstávají v repu jako původní prototyp, ale hlavní UI je desktop aplikace.
