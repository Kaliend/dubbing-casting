from __future__ import annotations

import threading

_local = threading.local()

TRANSLATIONS: dict[str, dict[str, str]] = {
    "cs": {
        # --- actions ---
        "action.new_project": "Nový projekt",
        "action.open_project": "Otevřít projekt...",
        "action.bulk_import": "Hromadný import děl...",
        "action.save_project": "Uložit projekt",
        "action.save_as": "Uložit projekt jako...",
        "action.recalculate": "Přepočítat",
        "action.export_xlsx": "Export XLSX...",
        "action.komplet_focus": "Režim Komplet",
        "action.komplet_focus.tip": "Skryje vedlejší panely a zvětší sekci Komplet.",
        "action.quit": "Konec",
        "action.lang_cs": "Čeština",
        "action.lang_en": "English",
        # --- menus ---
        "menu.file": "Soubor",
        "menu.tools": "Nástroje",
        "menu.language": "Jazyk / Language",
        "toolbar.main": "Hlavní panel",
        # --- window / project ---
        "project.default_title": "Obsazení projektu",
        "project.unsaved": "Neuložený projekt",
        "app.name": "Obsazování",
        # --- labels ---
        "label.project_name": "Název projektu",
        "label.search": "Hledat",
        "label.show": "Zobrazit",
        # --- group boxes ---
        "group.episodes": "Díla",
        "group.komplet": "Komplet",
        "group.actor_summary": "Souhrn dabérů",
        "group.validation": "Validace",
        # --- metric cards ---
        "metric.characters": "Postavy",
        "metric.inputs": "Vstupy",
        "metric.replicas": "Repliky",
        "metric.missing": "Neobsazeno",
        # --- episode buttons ---
        "btn.add_episode": "Přidat dílo",
        "btn.rename_episode": "Přejmenovat dílo",
        "btn.remove_episode": "Odebrat dílo",
        # --- filter combo ---
        "filter.all": "Všechny postavy",
        "filter.unassigned": "Jen neobsazené",
        "filter.assigned": "Jen obsazené",
        # --- placeholders ---
        "search.placeholder": "Hledat v postavách, dabérech a poznámkách",
        "title.placeholder": "Název projektu",
        # --- checkboxes ---
        "checkbox.herci_by_episode": "V exportu rozdělit HERCI po dílech",
        "checkbox.herci_by_episode.tip": "Přidá do listu HERCI dvojice sloupců VSTUPY/REPLIKY pro každé aktivní dílo.",
        # --- validation action button ---
        "btn.unify_name": "Sjednotit název...",
        "btn.unify_name.tip": "Dostupné jen pro validace typu možná jde o stejného dabéra.",
        # --- status bar messages ---
        "status.komplet_on": "Režim Komplet je zapnutý.",
        "status.komplet_off": "Původní rozvržení je obnovené.",
        "status.analysis_current": "Přehled je aktuální.",
        "status.new_project": "Vytvořen nový projekt.",
        "status.project_loaded": "Projekt načten.",
        "status.project_saved": "Projekt uložen.",
        "status.project_saved_to": "Projekt uložen do {name}.",
        "status.episode_added": "Přidáno dílo {label}.",
        "status.episode_renamed": "Dílo bylo přejmenováno na {label}.",
        "status.episode_removed": "Dílo {label} bylo odebráno.",
        "status.episode_cleared": "Dílo {label} bylo vyčištěno.",
        "status.episode_loaded": "Načteno dílo {label}.",
        "status.episode_loaded_sheet": "Načteno dílo {label} z listu {sheet}.",
        "status.bulk_imported": "Hromadně naimportováno {count} děl.",
        "status.actor_unified": "Sjednocen název dabéra na {name}.",
        "status.export_done": "Export hotový: {name}",
        # --- dialogs ---
        "dialog.open_project": "Otevřít projekt",
        "dialog.save_as": "Uložit projekt jako",
        "dialog.export_xlsx": "Export XLSX",
        "dialog.load_input": "Načíst vstup pro dílo {label}",
        "dialog.select_sheet": "Vybrat list pro import",
        "dialog.select_sheet.msg": "Excel obsahuje více použitelných listů. Vyber list pro aktivní dílo:",
        "dialog.bulk_import": "Hromadný import děl",
        "dialog.bulk_import.msg": "Vyber zdroj hromadného importu:",
        "dialog.bulk_import.confirm_title": "Potvrdit hromadný import",
        "dialog.bulk_import.confirm_msg": "Naimportovat vybrané zdroje do projektu?",
        "dialog.bulk_import.btn": "Importovat",
        "dialog.select_workbook": "Vybrat workbook pro hromadný import",
        "dialog.select_files": "Vybrat soubory pro hromadný import",
        "dialog.select_dir": "Vybrat složku pro hromadný import",
        # --- error dialog titles ---
        "error.open_project": "Načtení projektu selhalo",
        "error.save_project": "Uložení projektu selhalo",
        "error.export": "Export selhal",
        "error.add_episode": "Nelze přidat dílo",
        "error.rename_episode": "Přejmenování selhalo",
        "error.remove_episode": "Odebrání selhalo",
        "error.load_input": "Načtení vstupu selhalo",
        "error.bulk_import": "Hromadný import selhal",
        "error.bulk_import_warning": "Hromadný import nelze provést",
        # --- file filters ---
        "filter.project_file": "Projekt Obsazování (*.json);;JSON (*.json)",
        "filter.input_files": "Podporované vstupy (*.txt *.tsv *.csv *.xlsx *.doc *.docx);;Všechny soubory (*)",
        "filter.excel": "Excel Workbook (*.xlsx)",
        # --- confirm / remove episode ---
        "confirm.remove_episode.title": "Odebrat dílo",
        "confirm.remove_episode": "Opravdu chceš odebrat dílo {label}? Jeho obsah bude smazán.",
        # --- unsaved changes ---
        "unsaved.title": "Neuložené změny",
        "unsaved.text": "Projekt obsahuje neuložené změny.",
        "unsaved.info": "Chceš je před pokračováním uložit?",
        # --- unify actor name ---
        "unify.title": "Sjednotit název dabéra",
        "unify.label": "Vyber cílový název dabéra:",
        "unify.custom": "Vlastní název...",
        "unify.custom.title": "Vlastní název dabéra",
        "unify.custom.label": "Zadej sjednocený název dabéra:",
        "unify.confirm.title": "Potvrdit sjednocení názvu",
        "unify.confirm.msg": 'Opravdu chceš sjednotit varianty dabéra {variants} na \u201e{name}\u201c?',
        "unify.fail.title": "Sjednocení selhalo",
        "unify.fail.empty": "Cílový název dabéra nesmí být prázdný.",
        "unify.info.title": "Sjednocení názvu",
        "unify.info.not_enough": "Vybraná validace neobsahuje dost variant pro sjednocení.",
        "unify.info.none_found": "V projektu nebyla nalezena žádná odpovídající obsazení.",
        # --- rename episode ---
        "rename.title": "Přejmenovat dílo",
        "rename.label": "Název díla:",
        # --- bulk import plan ---
        "bulk.source.workbook": "Jeden workbook s více listy",
        "bulk.source.files": "Více souborů",
        "bulk.source.dir": "Složka se soubory",
        "bulk.plan.starts": "Import začne od díla {label} a zpracuje {count} zdrojů.",
        "bulk.plan.overwrites": "Přepíše obsah {count} existujících děl.",
        "bulk.plan.creates": "Přidá {count} nových děl.",
        "bulk.plan.episode": "Dílo {num}: {label}",
        "bulk.plan.overwrites_suffix": "přepíše obsah",
        "bulk.plan.creates_suffix": "nové dílo",
        # --- summary status ---
        "summary.missing": "Chybí obsadit {chars} postav, {inputs} vstupů a {replicas} replik.",
        "summary.missing.also": " Další validace: {warnings} upozornění",
        "summary.missing.also.info": ", {info} informací.",
        "summary.missing.also.no_info": ".",
        "summary.validation.warnings": "Validace našla {warnings} upozornění{suffix}.",
        "summary.validation.info_only": "Validace našla {info} informací.",
        "summary.ok": "Všechny postavy mají přiřazeného dabéra.",
        "summary.validation.warnings.suffix": " a {info} informací",
        # --- validation status ---
        "validation.status.warnings": "Validace hlásí {warnings} upozornění{suffix}.",
        "validation.status.info_only": "Validace hlásí {info} informací.",
        "validation.status.ok": "Další rizika nebyla nalezena.",
        "validation.status.warnings.suffix": " a {info} informací",
        # --- table column headers ---
        "col.character": "Postava",
        "col.inputs": "Vstupy",
        "col.replicas": "Repliky",
        "col.actor": "Dabér",
        "col.note": "Poznámka",
        "col.level": "Úroveň",
        "col.area": "Oblast",
        "col.detail": "Detail",
        # --- validation severity ---
        "severity.warning": "Upozornění",
        "severity.info": "Info",
        # --- tooltip ---
        "tooltip.unassigned": "Postava zatím nemá přiřazeného dabéra.",
        # --- episode editor ---
        "editor.intro": "Podporované formáty: TXT/CSV/TSV/XLSX s poli POSTAVA / TC / TEXT nebo POSTAVA / VSTUPY / REPLIKY.",
        "editor.load": "Načíst soubor",
        "editor.clear": "Vyčistit dílo",
        "editor.placeholder": "Dílo {label}: vlož TSV/CSV/TXT nebo importuj XLSX s hlavičkami POSTAVA, TC, TEXT nebo POSTAVA, VSTUPY, REPLIKY.",
        # --- core errors ---
        "core.error.bad_input_cols": "Vstup musí obsahovat aspoň sloupec POSTAVA a TEXT nebo VSTUPY/REPLIKY.",
        "core.error.unrecognized_input": "Nepodařilo se rozpoznat vstup. Očekávám hlavičky POSTAVA/TC/TEXT nebo POSTAVA/VSTUPY/REPLIKY.",
        # --- project_store errors ---
        "store.error.workbook_not_xlsx": "Hromadný import workbooku podporuje jen .xlsx soubory.",
        "store.error.no_files": "Nebyly vybrány žádné soubory pro hromadný import.",
        "store.error.multi_sheet": 'Soubor {name} obsahuje více použitelných listů. Pro takový workbook použij režim \u201eJeden workbook s více listy\u201c.',
        "store.error.no_inputs": "Ve vybraných souborech nebyl nalezen žádný podporovaný vstup.",
        "store.error.not_dir": "{name} není platná složka.",
        "store.error.empty_dir": "Ve vybrané složce nebyly nalezeny podporované soubory.",
        "store.error.invalid_json": "Soubor {name} neobsahuje platný JSON projekt.",
        "store.error.invalid_json_structure": "Projektový JSON musí obsahovat objekt s title, episodes a assignments.",
        # --- app_state errors ---
        "state.error.no_path": "Chybí cílová cesta pro uložení projektu.",
        "state.error.no_sources": "Chybí zdroje pro hromadný import.",
        "state.error.bad_position": "Neplatná cílová pozice pro import.",
        "state.error.no_slots": "Od vybraného díla je k dispozici jen {available} slotů, ale import vyžaduje {count} děl.",
        "state.error.no_plan": "Chybí plán hromadného importu.",
        "state.error.max_episodes": "Projekt může obsahovat maximálně {max} děl.",
        "state.error.episode_not_found": "Dílo neexistuje.",
        "state.error.empty_label": "Název díla nesmí být prázdný.",
        "state.error.duplicate_label": 'Dílo s názvem \u201e{label}\u201c už existuje.',
        "state.error.min_episodes": "Projekt musí obsahovat alespoň jedno dílo.",
        "state.error.empty_actor": "Cílový název dabéra nesmí být prázdný.",
        "state.error.no_variants": "Chybí varianty dabéra ke sjednocení.",
        # --- validation messages ---
        "val.unassigned": "Postava {character} zatím nemá dabéra ({inputs} vstupů, {replicas} replik{works}).",
        "val.unassigned.works": "; díla: {works}",
        "val.generic_bucket": "Postava {character} má {replicas} replik. Zkontroluj, jestli pod ní není víc konkrétních rolí.",
        "val.char_variants": "Možná jde o stejnou postavu: {variants}.",
        "val.actor_variants": "Možná jde o stejného dabéra: {variants}.",
        "val.actor_high_load": "Dabér {actor} má vysokou celkovou zátěž ({inputs} vstupů, {replicas} replik).",
        "val.actor_many_chars": "Dabér {actor} má přiřazeno {chars} postav ({inputs} vstupů, {replicas} replik).",
        "val.actor_episode_high_load": "Dabér {actor} má v díle {label} vysokou zátěž ({inputs} vstupů, {replicas} replik).",
        "val.actor_episode_many_chars": "Dabér {actor} má v díle {label} {chars} postavy ({inputs} vstupů, {replicas} replik).",
        # --- validation categories ---
        "cat.unassigned": "Neobsazeno",
        "cat.characters": "Postavy",
        "cat.names": "Jména",
        "cat.load": "Zátěž",
        "cat.casting": "Obsazení",
        # --- server ---
        "server.invalid_json": "Neplatný JSON payload.",
        # --- exporter status cell ---
        "export.missing_status": "JEŠTĚ CHYBÍ DOPLNIT {chars} POSTAV, {inputs} VSTUPŮ A {replicas} REPLIK.",
        "export.complete_status": "VŠECHNY POSTAVY OBSAZENY.",
        # --- main tabs ---
        "tab.prehled": "Přehled",
        "tab.herci": "Herci",
        "tab.komplet": "Komplet",
        "tab.daberi": "Dabéři",
        # --- panel toggles ---
        "toggle.panels": "Panely:",
        "toggle.dila": "Díla",
        "toggle.komplet": "Komplet",
        "toggle.summary": "Souhrn dabérů",
        "toggle.validation": "Validace",
    },
    "en": {
        # --- actions ---
        "action.new_project": "New project",
        "action.open_project": "Open project...",
        "action.bulk_import": "Bulk import episodes...",
        "action.save_project": "Save project",
        "action.save_as": "Save project as...",
        "action.recalculate": "Recalculate",
        "action.export_xlsx": "Export XLSX...",
        "action.komplet_focus": "Full cast mode",
        "action.komplet_focus.tip": "Hides side panels and expands the Full cast section.",
        "action.quit": "Quit",
        "action.lang_cs": "Čeština",
        "action.lang_en": "English",
        # --- menus ---
        "menu.file": "File",
        "menu.tools": "Tools",
        "menu.language": "Language / Jazyk",
        "toolbar.main": "Main toolbar",
        # --- window / project ---
        "project.default_title": "Project casting",
        "project.unsaved": "Unsaved project",
        "app.name": "Casting",
        # --- labels ---
        "label.project_name": "Project name",
        "label.search": "Search",
        "label.show": "Show",
        # --- group boxes ---
        "group.episodes": "Episodes",
        "group.komplet": "Full cast",
        "group.actor_summary": "Voice actor summary",
        "group.validation": "Validation",
        # --- metric cards ---
        "metric.characters": "Characters",
        "metric.inputs": "Inputs",
        "metric.replicas": "Replicas",
        "metric.missing": "Unassigned",
        # --- episode buttons ---
        "btn.add_episode": "Add episode",
        "btn.rename_episode": "Rename episode",
        "btn.remove_episode": "Remove episode",
        # --- filter combo ---
        "filter.all": "All characters",
        "filter.unassigned": "Unassigned only",
        "filter.assigned": "Assigned only",
        # --- placeholders ---
        "search.placeholder": "Search characters, voice actors and notes",
        "title.placeholder": "Project name",
        # --- checkboxes ---
        "checkbox.herci_by_episode": "Split ACTORS by episode in export",
        "checkbox.herci_by_episode.tip": "Adds INPUTS/REPLICAS column pairs per active episode to the ACTORS sheet.",
        # --- validation action button ---
        "btn.unify_name": "Unify name...",
        "btn.unify_name.tip": "Available only for \"possible same voice actor\" validations.",
        # --- status bar messages ---
        "status.komplet_on": "Full cast mode is on.",
        "status.komplet_off": "Original layout restored.",
        "status.analysis_current": "View is up to date.",
        "status.new_project": "New project created.",
        "status.project_loaded": "Project loaded.",
        "status.project_saved": "Project saved.",
        "status.project_saved_to": "Project saved to {name}.",
        "status.episode_added": "Episode {label} added.",
        "status.episode_renamed": "Episode renamed to {label}.",
        "status.episode_removed": "Episode {label} removed.",
        "status.episode_cleared": "Episode {label} cleared.",
        "status.episode_loaded": "Episode {label} loaded.",
        "status.episode_loaded_sheet": "Episode {label} loaded from sheet {sheet}.",
        "status.bulk_imported": "Bulk imported {count} episodes.",
        "status.actor_unified": "Voice actor name unified to {name}.",
        "status.export_done": "Export done: {name}",
        # --- dialogs ---
        "dialog.open_project": "Open project",
        "dialog.save_as": "Save project as",
        "dialog.export_xlsx": "Export XLSX",
        "dialog.load_input": "Load input for episode {label}",
        "dialog.select_sheet": "Select sheet for import",
        "dialog.select_sheet.msg": "The workbook contains multiple usable sheets. Select a sheet for the active episode:",
        "dialog.bulk_import": "Bulk import episodes",
        "dialog.bulk_import.msg": "Select bulk import source:",
        "dialog.bulk_import.confirm_title": "Confirm bulk import",
        "dialog.bulk_import.confirm_msg": "Import selected sources into the project?",
        "dialog.bulk_import.btn": "Import",
        "dialog.select_workbook": "Select workbook for bulk import",
        "dialog.select_files": "Select files for bulk import",
        "dialog.select_dir": "Select folder for bulk import",
        # --- error dialog titles ---
        "error.open_project": "Failed to open project",
        "error.save_project": "Failed to save project",
        "error.export": "Export failed",
        "error.add_episode": "Cannot add episode",
        "error.rename_episode": "Rename failed",
        "error.remove_episode": "Remove failed",
        "error.load_input": "Failed to load input",
        "error.bulk_import": "Bulk import failed",
        "error.bulk_import_warning": "Bulk import cannot proceed",
        # --- file filters ---
        "filter.project_file": "Casting Project (*.json);;JSON (*.json)",
        "filter.input_files": "Supported inputs (*.txt *.tsv *.csv *.xlsx *.doc *.docx);;All files (*)",
        "filter.excel": "Excel Workbook (*.xlsx)",
        # --- confirm / remove episode ---
        "confirm.remove_episode.title": "Remove episode",
        "confirm.remove_episode": "Really remove episode {label}? Its content will be deleted.",
        # --- unsaved changes ---
        "unsaved.title": "Unsaved changes",
        "unsaved.text": "The project has unsaved changes.",
        "unsaved.info": "Do you want to save before continuing?",
        # --- unify actor name ---
        "unify.title": "Unify voice actor name",
        "unify.label": "Select target voice actor name:",
        "unify.custom": "Custom name...",
        "unify.custom.title": "Custom voice actor name",
        "unify.custom.label": "Enter unified voice actor name:",
        "unify.confirm.title": "Confirm name unification",
        "unify.confirm.msg": "Really unify voice actor variants {variants} to \"{name}\"?",
        "unify.fail.title": "Unification failed",
        "unify.fail.empty": "Target voice actor name must not be empty.",
        "unify.info.title": "Unify name",
        "unify.info.not_enough": "The selected validation does not contain enough variants for unification.",
        "unify.info.none_found": "No matching assignments found in the project.",
        # --- rename episode ---
        "rename.title": "Rename episode",
        "rename.label": "Episode name:",
        # --- bulk import plan ---
        "bulk.source.workbook": "Single workbook with multiple sheets",
        "bulk.source.files": "Multiple files",
        "bulk.source.dir": "Folder with files",
        "bulk.plan.starts": "Import will start from episode {label} and process {count} sources.",
        "bulk.plan.overwrites": "Will overwrite content of {count} existing episodes.",
        "bulk.plan.creates": "Will add {count} new episodes.",
        "bulk.plan.episode": "Episode {num}: {label}",
        "bulk.plan.overwrites_suffix": "overwrites content",
        "bulk.plan.creates_suffix": "new episode",
        # --- summary status ---
        "summary.missing": "Missing {chars} characters, {inputs} inputs and {replicas} replicas.",
        "summary.missing.also": " Also: {warnings} warnings",
        "summary.missing.also.info": ", {info} info items.",
        "summary.missing.also.no_info": ".",
        "summary.validation.warnings": "Validation found {warnings} warnings{suffix}.",
        "summary.validation.info_only": "Validation found {info} info items.",
        "summary.ok": "All characters have an assigned voice actor.",
        "summary.validation.warnings.suffix": " and {info} info items",
        # --- validation status ---
        "validation.status.warnings": "Validation reports {warnings} warnings{suffix}.",
        "validation.status.info_only": "Validation reports {info} info items.",
        "validation.status.ok": "No further risks found.",
        "validation.status.warnings.suffix": " and {info} info items",
        # --- table column headers ---
        "col.character": "Character",
        "col.inputs": "Inputs",
        "col.replicas": "Replicas",
        "col.actor": "Voice actor",
        "col.note": "Note",
        "col.level": "Level",
        "col.area": "Area",
        "col.detail": "Detail",
        # --- validation severity ---
        "severity.warning": "Warning",
        "severity.info": "Info",
        # --- tooltip ---
        "tooltip.unassigned": "Character has no assigned voice actor yet.",
        # --- episode editor ---
        "editor.intro": "Supported formats: TXT/CSV/TSV/XLSX with fields CHARACTER / TC / TEXT or CHARACTER / INPUTS / REPLICAS.",
        "editor.load": "Load file",
        "editor.clear": "Clear episode",
        "editor.placeholder": "Episode {label}: paste TSV/CSV/TXT or import XLSX with headers CHARACTER, TC, TEXT or CHARACTER, INPUTS, REPLICAS.",
        # --- core errors ---
        "core.error.bad_input_cols": "Input must contain at least a CHARACTER and TEXT column or INPUTS/REPLICAS.",
        "core.error.unrecognized_input": "Could not recognize input format. Expected headers CHARACTER/TC/TEXT or CHARACTER/INPUTS/REPLICAS.",
        # --- project_store errors ---
        "store.error.workbook_not_xlsx": "Bulk workbook import only supports .xlsx files.",
        "store.error.no_files": "No files selected for bulk import.",
        "store.error.multi_sheet": "File {name} contains multiple usable sheets. Use the \"Single workbook with multiple sheets\" mode for this workbook.",
        "store.error.no_inputs": "No supported inputs found in the selected files.",
        "store.error.not_dir": "{name} is not a valid folder.",
        "store.error.empty_dir": "No supported files found in the selected folder.",
        "store.error.invalid_json": "File {name} does not contain a valid JSON project.",
        "store.error.invalid_json_structure": "Project JSON must contain an object with title, episodes and assignments.",
        # --- app_state errors ---
        "state.error.no_path": "No target path for saving the project.",
        "state.error.no_sources": "No sources for bulk import.",
        "state.error.bad_position": "Invalid target position for import.",
        "state.error.no_slots": "Only {available} slots available from the selected episode, but import requires {count} episodes.",
        "state.error.no_plan": "No bulk import plan provided.",
        "state.error.max_episodes": "Project can contain at most {max} episodes.",
        "state.error.episode_not_found": "Episode does not exist.",
        "state.error.empty_label": "Episode name must not be empty.",
        "state.error.duplicate_label": "Episode named \"{label}\" already exists.",
        "state.error.min_episodes": "Project must contain at least one episode.",
        "state.error.empty_actor": "Target voice actor name must not be empty.",
        "state.error.no_variants": "No voice actor variants to unify.",
        # --- validation messages ---
        "val.unassigned": "Character {character} has no voice actor yet ({inputs} inputs, {replicas} replicas{works}).",
        "val.unassigned.works": "; episodes: {works}",
        "val.generic_bucket": "Character {character} has {replicas} replicas. Check if it groups multiple specific roles.",
        "val.char_variants": "Possibly the same character: {variants}.",
        "val.actor_variants": "Possibly the same voice actor: {variants}.",
        "val.actor_high_load": "Voice actor {actor} has high total load ({inputs} inputs, {replicas} replicas).",
        "val.actor_many_chars": "Voice actor {actor} is assigned to {chars} characters ({inputs} inputs, {replicas} replicas).",
        "val.actor_episode_high_load": "Voice actor {actor} has high load in episode {label} ({inputs} inputs, {replicas} replicas).",
        "val.actor_episode_many_chars": "Voice actor {actor} has {chars} characters in episode {label} ({inputs} inputs, {replicas} replicas).",
        # --- validation categories ---
        "cat.unassigned": "Unassigned",
        "cat.characters": "Characters",
        "cat.names": "Names",
        "cat.load": "Load",
        "cat.casting": "Casting",
        # --- server ---
        "server.invalid_json": "Invalid JSON payload.",
        # --- exporter status cell ---
        "export.missing_status": "STILL MISSING {chars} CHARACTERS, {inputs} INPUTS AND {replicas} REPLICAS.",
        "export.complete_status": "ALL CHARACTERS ASSIGNED.",
        # --- main tabs ---
        "tab.prehled": "Overview",
        "tab.herci": "Actors",
        "tab.komplet": "Full Cast",
        "tab.daberi": "Voice Actors",
        # --- panel toggles ---
        "toggle.panels": "Panels:",
        "toggle.dila": "Episodes",
        "toggle.komplet": "Full Cast",
        "toggle.summary": "Actor Summary",
        "toggle.validation": "Validation",
    },
}


def set_language(lang: str) -> None:
    if lang in TRANSLATIONS:
        _local.lang = lang


def get_language() -> str:
    return getattr(_local, "lang", "cs")


def t(key: str, **kwargs: object) -> str:
    lang = get_language()
    lang_dict = TRANSLATIONS.get(lang, TRANSLATIONS["cs"])
    template = lang_dict.get(key) or TRANSLATIONS["cs"].get(key, key)
    if kwargs:
        try:
            return template.format(**kwargs)
        except (KeyError, IndexError):
            return template
    return template
