from __future__ import annotations

import re
import shutil
import sys
from dataclasses import dataclass
from pathlib import Path

import pythoncom
import win32com.client as win32


CONFIG_FILE_NAME = "configurazione_schede.txt"

SECTION_BY_CHOICE = {
    "1": "PASTA_CORTA",
    "2": "PASTA_LUNGA",
    "3": "SILO",
}

XL_UP = -4162
XL_PASTE_FORMATS = -4122
XL_NONE = -4142
XL_INSIDE_HORIZONTAL = 12
XL_PATTERN_NONE = -4142
XL_COLORINDEX_NONE = -4142
COLOR_WHITE = 16777215
XL_CONTINUOUS = 1
XL_THIN = 2
XL_MEDIUM = -4138
XL_EDGE_LEFT = 7
XL_EDGE_TOP = 8
XL_EDGE_BOTTOM = 9
XL_EDGE_RIGHT = 10
XL_INSIDE_VERTICAL = 11

SPLIT_RULES_BY_SECTION = {
    "PASTA_CORTA": {
        "label": "Tagliapasta",
        "keywords": ["TAGLIAPASTA"],
        "only_sheet": None,
    },
    "PASTA_LUNGA": {
        "label": "Stenditrice",
        "keywords": ["STENDITRICE", "PAREGGIATORE", "TRITURATORE"],
        "only_sheet": "B",
    },
    "SILO": {
        "label": "Tagliapasta",
        "keywords": ["TAGLIAPASTA"],
        "only_sheet": None,
    },
}


@dataclass
class ConfigRow:
    title: str
    sheet_name: str
    split_tagliapasta: bool


def log(message: str) -> None:
    print(message, flush=True)


def pause_and_exit(code: int) -> None:
    try:
        input("\nPremi INVIO per chiudere...")
    except EOFError:
        pass
    raise SystemExit(code)


def parse_measure(raw_value: object, expected_unit: str) -> float:
    if raw_value is None:
        return 0.0

    if isinstance(raw_value, (int, float)):
        return float(raw_value)

    text = str(raw_value).strip()
    if not text:
        return 0.0

    compact = text.replace(" ", "").replace(",", ".")
    match = re.match(r"^(-?\d+(?:\.\d+)?)([A-Za-z]*)$", compact)
    if not match:
        return 0.0

    number_part, unit_part = match.groups()
    if unit_part and unit_part.lower() != expected_unit.lower():
        return 0.0

    return float(number_part)


def normalize_text(value: object) -> str:
    if value is None:
        return ""
    return str(value).upper()


def normalize_output_number(value: float) -> float:
    return float(f"{value:.2f}")


def force_remove_fill(sheet, from_row: int, to_row: int, from_col: int, to_col: int) -> None:
    """
    Rimozione hard dello sfondo cella-per-cella.
    Necessario perché in alcuni template .xls lo stile copiato può mantenere fill verde
    anche dopo operazioni su range interi.
    """
    for row in range(from_row, to_row + 1):
        for col in range(from_col, to_col + 1):
            cell = sheet.Cells(row, col)
            # In alcuni .xls il "no fill" viene reinterpretato con un ColorIndex di palette.
            # Forziamo quindi sfondo pieno bianco per evitare qualsiasi tinta verde.
            cell.Interior.Pattern = 1
            cell.Interior.Color = COLOR_WHITE
            cell.Interior.ColorIndex = XL_COLORINDEX_NONE


def read_config_for_section(config_path: Path, section: str) -> list[ConfigRow]:
    if not config_path.exists():
        raise FileNotFoundError(f"File configurazione non trovato: {config_path}")

    rows: list[ConfigRow] = []
    current_section: str | None = None

    for raw_line in config_path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line:
            continue

        if line.startswith("#"):
            current_section = line[1:].strip().upper()
            continue

        if current_section != section:
            continue

        match = re.match(r'^"([^"]+)";"([^"]+)";([01])$', line)
        if not match:
            log(f"[WARN] Riga configurazione ignorata (formato non valido): {line}")
            continue

        title, sheet_name, split_raw = match.groups()
        rows.append(
            ConfigRow(
                title=title.strip(),
                sheet_name=sheet_name.strip(),
                split_tagliapasta=(split_raw == "1"),
            )
        )

    return rows


def choose_section() -> str:
    log("Seleziona il tipo impianto:")
    log("  1) Pasta Corta")
    log("  2) Pasta Lunga")
    log("  3) Silo")

    choice = input("Scelta [1/2/3]: ").strip()
    section = SECTION_BY_CHOICE.get(choice)
    if not section:
        raise ValueError("Scelta non valida.")
    return section


def find_input_workbook(base_dir: Path) -> Path:
    candidates_all = sorted(
        p
        for p in base_dir.glob("*.xlsx")
        if not p.name.startswith("~$")
    )
    candidates = [
        p for p in candidates_all
        if "esempio" not in p.stem.lower()
        and not p.stem.lower().endswith(" - output")
    ]

    if not candidates:
        raise FileNotFoundError("Nessun file .xlsx di input trovato nella cartella dell'eseguibile.")
    if len(candidates) > 1:
        names = ", ".join(p.name for p in candidates)
        raise RuntimeError(
            "Trovati più file .xlsx. Lascia solo il file di input nella cartella. "
            f"File trovati: {names}"
        )
    return candidates[0]


def derive_summary_output_name(input_file: Path) -> str:
    return f"{input_file.stem} - Smistamento potenza.xls"


def derive_input_output_name(input_file: Path) -> str:
    return f"{input_file.stem}_Output{input_file.suffix}"


def sheet_exists(workbook, name: str) -> bool:
    for sh in workbook.Worksheets:
        if str(sh.Name).strip().lower() == name.strip().lower():
            return True
    return False


def get_sheet(workbook, name: str):
    for sh in workbook.Worksheets:
        if str(sh.Name).strip().lower() == name.strip().lower():
            return sh
    return None


def compute_totals_for_sheet(
    sheet,
    split_enabled: bool,
    sheet_name: str,
    split_keywords: list[str],
    split_only_sheet: str | None,
) -> tuple[float, float, float, float]:
    used_range = sheet.UsedRange
    first_row = used_range.Row
    last_row = first_row + used_range.Rows.Count - 1
    if last_row < 8:
        last_row = 8

    kw_no_tag = 0.0
    kw_tag = 0.0
    amp_no_tag = 0.0
    amp_tag = 0.0

    for row in range(8, last_row + 1):
        col_a = normalize_text(sheet.Cells(row, "A").Value)
        if col_a.startswith("TOTALE"):
            continue

        kw_val = parse_measure(sheet.Cells(row, "D").Value, "kW")
        amp_val = parse_measure(sheet.Cells(row, "F").Value, "A")
        if kw_val == 0.0 and amp_val == 0.0:
            continue

        desc = normalize_text(sheet.Cells(row, "B").Value) + " " + normalize_text(sheet.Cells(row, "C").Value)
        is_split_bucket = any(keyword in desc for keyword in split_keywords)
        if split_only_sheet and sheet_name.strip().upper() != split_only_sheet.strip().upper():
            is_split_bucket = False

        if split_enabled and is_split_bucket:
            kw_tag += kw_val
            amp_tag += amp_val
        else:
            kw_no_tag += kw_val
            amp_no_tag += amp_val

    return kw_no_tag, kw_tag, amp_no_tag, amp_tag


def insert_totals_rows(
    sheet,
    split_enabled: bool,
    sheet_name: str,
    split_label: str,
    split_keywords: list[str],
    split_only_sheet: str | None,
) -> tuple[float, float, float, float]:
    has_existing_totals = normalize_text(sheet.Range("A4").Value).startswith("TOTALE")
    expected_split_total = f"TOTALE {split_label}".upper()
    has_existing_tag_totals = normalize_text(sheet.Range("A5").Value).startswith(expected_split_total)

    if not (has_existing_totals and has_existing_tag_totals):
        sheet.Rows("4:6").Insert()
    sheet.Range("A4").Value = "Totale"
    sheet.Range("A5").Value = f"Totale {split_label}"
    sheet.Range("D7").Value = "kW"
    sheet.Range("F7").Value = "A"

    # Copia stile da una riga dati reale (dopo lo shift)
    sheet.Rows(7).Copy()
    sheet.Rows(4).PasteSpecial(XL_PASTE_FORMATS)
    sheet.Rows(5).PasteSpecial(XL_PASTE_FORMATS)

    kw_no_tag, kw_tag, amp_no_tag, amp_tag = compute_totals_for_sheet(
        sheet,
        split_enabled,
        sheet_name,
        split_keywords,
        split_only_sheet,
    )

    sheet.Range("D4").Value = normalize_output_number(kw_no_tag)
    sheet.Range("F4").Value = normalize_output_number(amp_no_tag)

    if split_enabled:
        sheet.Range("D5").Value = normalize_output_number(kw_tag)
        sheet.Range("F5").Value = normalize_output_number(amp_tag)
    else:
        sheet.Range("D5").Value = ""
        sheet.Range("F5").Value = ""

    return kw_no_tag, kw_tag, amp_no_tag, amp_tag


def create_summary_file(
    excel_app,
    output_path: Path,
    summary_rows: list[dict],
) -> None:
    wb = excel_app.Workbooks.Add()
    try:
        ws = wb.Worksheets(1)
        ws.Name = "Smistamento potenza"
        start_row = 11

        # Layout richiesto: lavoro solo su colonne A..F
        for row in range(1, 500):
            for col in range(1, 7):
                ws.Cells(row, col).Value = None

        # Imposta larghezze colonne principali.
        ws.Columns("A").ColumnWidth = 37.0
        ws.Columns("B").ColumnWidth = 6.71
        ws.Columns("C").ColumnWidth = 7.43
        ws.Columns("D").ColumnWidth = 4.71
        ws.Columns("E").ColumnWidth = 4.57
        ws.Columns("F").ColumnWidth = 28.29

        # Font richiesto
        ws.Range("A1:F500").Font.Name = "Arial"
        ws.Range("A1:F500").Font.Size = 14

        # Righe 1 e 2 unificate A..F e vuote
        ws.Range("A1:F1").Merge()
        ws.Range("A2:F2").Merge()

        # Righe 3 e 4 vuote (già pulite)

        # Riga 5: griglia completa e intestazioni
        ws.Cells(5, 2).Value = "kW"
        ws.Cells(5, 3).Value = "A"
        ws.Cells(5, 4).Value = "sez."

        # Righe 5,6,7: griglia completa A..F (bordi medi)
        for row in (5, 6, 7):
            rng = ws.Range(ws.Cells(row, 1), ws.Cells(row, 6))
            rng.Borders.LineStyle = XL_CONTINUOUS
            rng.Borders.Weight = XL_MEDIUM

        # Righe 8..10: solo bordi esterni A..F (bordi medi)
        outer = ws.Range("A8:F10")
        outer.Borders(XL_EDGE_LEFT).LineStyle = XL_CONTINUOUS
        outer.Borders(XL_EDGE_LEFT).Weight = XL_MEDIUM
        outer.Borders(XL_EDGE_TOP).LineStyle = XL_CONTINUOUS
        outer.Borders(XL_EDGE_TOP).Weight = XL_MEDIUM
        outer.Borders(XL_EDGE_BOTTOM).LineStyle = XL_CONTINUOUS
        outer.Borders(XL_EDGE_BOTTOM).Weight = XL_MEDIUM
        outer.Borders(XL_EDGE_RIGHT).LineStyle = XL_CONTINUOUS
        outer.Borders(XL_EDGE_RIGHT).Weight = XL_MEDIUM

        row = start_row
        prev_title = ""
        written_rows: list[dict] = []
        for item in summary_rows:
            title = item["title"]
            kw = item["kw"]
            amp = item["amp"]
            rif = item["rif"]
            is_tag = item["is_tagliapasta"]
            group_key = item["group_key"]

            # Mantiene la riga vuota separatrice prima del blocco "Sezionatore..."
            if prev_title and not prev_title.upper().startswith("SEZIONATORE") and title.upper().startswith("SEZIONATORE"):
                row += 1

            ws.Cells(row, 1).Value = title
            ws.Cells(row, 2).Value = float(f"{kw:.2f}")
            ws.Cells(row, 3).Value = float(f"{amp:.2f}")
            ws.Cells(row, 4).Value = rif  # RIF = nome foglio
            ws.Cells(row, 5).Value = ""
            ws.Cells(row, 6).Value = ""

            if is_tag:
                ws.Range(ws.Cells(row, 1), ws.Cells(row, 4)).Font.Color = 255  # rosso

            written_rows.append({"row": row, "group_key": group_key})
            prev_title = title
            row += 1

        # Bordi tabella dati A..F (riga 11..ultima scritta), bordo medio.
        if written_rows:
            last_written_row = written_rows[-1]["row"]
        else:
            last_written_row = start_row
        table_range = ws.Range(ws.Cells(start_row, 1), ws.Cells(last_written_row, 6))
        table_range.Borders.LineStyle = XL_CONTINUOUS
        table_range.Borders.Weight = XL_MEDIUM
        table_range.VerticalAlignment = -4107
        table_range.HorizontalAlignment = 1  # left

        # Rimuove il bordo interno orizzontale tra righe dello stesso foglio (split pressa/tagliapasta)
        for idx in range(1, len(written_rows)):
            prev = written_rows[idx - 1]
            cur = written_rows[idx]
            if prev["group_key"] and prev["group_key"] == cur["group_key"]:
                rng = ws.Range(ws.Cells(prev["row"], 1), ws.Cells(cur["row"], 6))
                rng.Borders(XL_INSIDE_HORIZONTAL).LineStyle = XL_NONE

        # Riga totale in basso, come nel formato validato.
        total_row = max(24, last_written_row + 3)
        ws.Cells(total_row, 1).Value = "TOTALE"
        ws.Cells(total_row, 2).Value = round(sum(item["kw"] for item in summary_rows), 2)
        ws.Cells(total_row, 3).Value = round(sum(item["amp"] for item in summary_rows), 2)
        ws.Cells(total_row, 4).Value = ""
        ws.Cells(total_row, 5).Value = ""
        ws.Cells(total_row, 6).Value = ""
        ws.Range(ws.Cells(total_row, 1), ws.Cells(total_row, 6)).Borders.LineStyle = XL_CONTINUOUS
        ws.Range(ws.Cells(total_row, 1), ws.Cells(total_row, 6)).Borders.Weight = XL_MEDIUM

        # Bordo destro colonna F dalla riga 1 fino a due righe sotto il totale.
        end_row = total_row + 2
        fcol = ws.Range(ws.Cells(1, 6), ws.Cells(end_row, 6))
        fcol.Borders(XL_EDGE_RIGHT).LineStyle = XL_CONTINUOUS
        fcol.Borders(XL_EDGE_RIGHT).Weight = XL_MEDIUM

        # Riga due righe sotto il totale: bordo basso da A a F.
        bottom_line = ws.Range(ws.Cells(end_row, 1), ws.Cells(end_row, 6))
        bottom_line.Borders(XL_EDGE_BOTTOM).LineStyle = XL_CONTINUOUS
        bottom_line.Borders(XL_EDGE_BOTTOM).Weight = XL_MEDIUM

        # Richiesta utente: colonna A sufficientemente larga per mostrare tutto
        # il testo (righe dalla 6 in avanti).
        ws.Columns("A").AutoFit()
        if ws.Columns("A").ColumnWidth < 37:
            ws.Columns("A").ColumnWidth = 37

        # Regola richiesta: nessuna cella con sfondo colorato nel file di output.
        used = ws.UsedRange
        # Rimuove eventuali formattazioni condizionali che potrebbero colorare B/C in verde.
        used.FormatConditions.Delete()
        from_row = used.Row
        from_col = used.Column
        to_row = from_row + used.Rows.Count - 1
        to_col = from_col + used.Columns.Count - 1
        force_remove_fill(ws, from_row, to_row, from_col, to_col)

        wb.SaveAs(str(output_path), FileFormat=56)  # 56 = .xls
    finally:
        wb.Close(SaveChanges=False)


def main() -> int:
    if getattr(sys, "frozen", False):
        base_dir = Path(sys.executable).resolve().parent
    else:
        base_dir = Path.cwd().resolve()
    config_path = base_dir / CONFIG_FILE_NAME

    log("=== Script Potenze ===")
    log(f"Cartella lavoro: {base_dir}")

    try:
        section = choose_section()
        split_rules = SPLIT_RULES_BY_SECTION.get(section, SPLIT_RULES_BY_SECTION["PASTA_CORTA"])
        split_label = split_rules["label"]
        split_keywords = split_rules["keywords"]
        split_only_sheet = split_rules["only_sheet"]
        config_rows = read_config_for_section(config_path, section)
        if not config_rows:
            raise RuntimeError(f"Nessuna riga configurata nella sezione #{section}.")

        input_file = find_input_workbook(base_dir)
        output_input_name = derive_input_output_name(input_file)
        output_input_path = base_dir / output_input_name
        shutil.copy2(input_file, output_input_path)

        output_summary_name = derive_summary_output_name(input_file)
        output_summary_path = base_dir / output_summary_name

        log(f"File input trovato: {input_file.name}")
        log(f"Creo copia output input: {output_input_path.name}")
        log(f"Sezione configurazione: #{section}")

        pythoncom.CoInitialize()
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        summary_rows: list[tuple[str, float, float]] = []
        try:
            wb_input = excel.Workbooks.Open(str(output_input_path))
            try:
                for cfg in config_rows:
                    sheet = get_sheet(wb_input, cfg.sheet_name)
                    if sheet is None:
                        log(f"[INFO] Foglio '{cfg.sheet_name}' non presente: salto.")
                        continue

                    log(f"[INFO] Elaboro foglio '{cfg.sheet_name}' ({cfg.title})...")
                    kw_no_tag, kw_tag, amp_no_tag, amp_tag = insert_totals_rows(
                        sheet,
                        cfg.split_tagliapasta,
                        cfg.sheet_name,
                        split_label,
                        split_keywords,
                        split_only_sheet,
                    )
                    summary_rows.append(
                        {
                            "title": cfg.title,
                            "kw": kw_no_tag,
                            "amp": amp_no_tag,
                            "rif": cfg.sheet_name,
                            "is_tagliapasta": False,
                            "group_key": cfg.sheet_name if cfg.split_tagliapasta else None,
                        }
                    )

                    if cfg.split_tagliapasta:
                        if "Pressa" in cfg.title:
                            split_title = cfg.title.replace("Pressa", split_label)
                        else:
                            split_title = f"{cfg.title} {split_label}"
                        summary_rows.append(
                            {
                                "title": split_title,
                                "kw": kw_tag,
                                "amp": amp_tag,
                                "rif": cfg.sheet_name,
                                "is_tagliapasta": True,
                                "group_key": cfg.sheet_name,
                            }
                        )

                wb_input.Save()
                log(f"[OK] File output input aggiornato: {output_input_path.name}")
            finally:
                wb_input.Close(SaveChanges=False)

            create_summary_file(excel, output_summary_path, summary_rows)
            log(f"[OK] File smistamento generato: {output_summary_path.name}")
        finally:
            excel.CutCopyMode = False
            excel.Quit()
            pythoncom.CoUninitialize()

        log("\nCompletato con successo.")
        return 0
    except Exception as exc:
        log(f"\n[ERRORE] {exc}")
        return 1


if __name__ == "__main__":
    exit_code = main()
    if exit_code != 0:
        pause_and_exit(exit_code)
    pause_and_exit(0)
