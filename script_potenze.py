from __future__ import annotations

import re
import shutil
import sys
from dataclasses import dataclass
from pathlib import Path

import pythoncom
import win32com.client as win32


CONFIG_FILE_NAME = "configurazione_schede.txt"

# Foglio "V": righe con testo tipo "COMANDO VITE 1" in colonna A (case insensitive).
_COMANDO_VITE_RE = re.compile(r"COMANDO\s+VITE\s*(\d+)", re.IGNORECASE)
_MAX_VITE_ROWS_SMISTAMENTO = 4
_SMISTAMENTO_VITE_START_ROW = 7
_SMISTAMENTO_MAIN_TABLE_ROW = 11

# Righe #SOLO_IDENTIFICATORE (es. #PASTA_LUNGA). Altre righe che iniziano con # restano commenti senza cambiare sezione.
_CONFIG_SECTION_HEADER_RE = re.compile(r"^#([A-Za-z_][A-Za-z0-9_]*)\s*$")


def _config_section_name_from_line(line: str) -> str | None:
    m = _CONFIG_SECTION_HEADER_RE.match(line.strip())
    return m.group(1).upper() if m else None


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

@dataclass
class ConfigRow:
    sheet_name: str


@dataclass
class LineTypeOption:
    key: str
    label: str


VALID_GROUP_KEYS = frozenset(
    {"pressa", "secondario", "recupero_polveri", "movimenti_linea", "movimenti_selezione_sili"}
)


@dataclass
class LineDisplayConfig:
    """Override etichette in tabella smistamento (per tipologia linea)."""

    label_overrides: dict[str, str]
    static_summary_rows: list[tuple[str | None, str, str]]
    primary_range_sheet: str | None
    offsheet_pressa_label: str | None
    offsheet_secondario_label: str | None
    sheet_group_labels: dict[tuple[str, str], str]


@dataclass
class SectionSplitRules:
    split_label: str
    pressa_label: str
    recupero_polveri_label: str
    movimenti_linea_label: str
    movimenti_selezione_sili_label: str
    pressa_range: tuple[int, int]
    secondary_range: tuple[int, int]
    recupero_polveri_range: tuple[int, int] | None
    movimenti_linea_range: tuple[int, int] | None
    movimenti_selezione_sili_range: tuple[int, int] | None


@dataclass
class UnmatchedRow:
    row: int
    raw_code: str
    parsed_code: int | None
    kw: float
    amp: float


def log(message: str) -> None:
    print(message, flush=True)


def pause_and_exit(code: int) -> None:
    try:
        input("\nPremi INVIO per chiudere...")
    except EOFError:
        pass
    raise SystemExit(code)


def parse_numeric_loose(raw_value: object) -> float:
    """Estrae il primo numero da una cella (per righe statiche smistamento se formato non ha suffisso kW/A)."""
    if raw_value is None:
        return 0.0
    if isinstance(raw_value, (int, float)):
        return float(raw_value)
    text = str(raw_value).strip()
    if not text:
        return 0.0
    compact = text.replace(" ", "").replace(",", ".")
    match = re.match(r"^(-?\d+(?:\.\d+)?)", compact)
    if match:
        return float(match.group(1))
    match = re.search(r"-?\d+(?:\.\d+)?", compact)
    return float(match.group(0)) if match else 0.0


def parse_grand_total_kw(raw_value: object) -> float:
    """
    Valore in kW per somme foglio senza gruppi: numeri Excel come kW; testo tipo 18.5kW;
    watt puri (es. 20W, senza 'k' prima di W) convertiti in kW. Evita di trattare 20W come 20 kW.
    """
    if raw_value is None:
        return 0.0
    if isinstance(raw_value, (int, float)):
        return float(raw_value)
    text = str(raw_value).strip()
    if not text:
        return 0.0
    compact = text.replace(" ", "").replace(",", ".")
    upper = compact.upper()
    mnum = re.match(r"^(-?\d+(?:\.\d+)?)", compact)
    if not mnum:
        return 0.0
    val = float(mnum.group(1))
    if upper.endswith("KW"):
        return val
    # Watt: termina con W ma non con KW (es. 20W, 500W)
    if upper.endswith("W") and not upper.endswith("KW"):
        return val / 1000.0
    return val


def parse_grand_total_amp(raw_value: object) -> float:
    """
    Corrente in A per somme foglio senza gruppi: numeri Excel come A; testo tipo 73.24A;
    milliampere (mA in cella, es. 50mA -> 0.05 A); kiloampere (kA -> A).
    Ordine suffissi: kA, mA, poi A (come per kW/W).
    """
    if raw_value is None:
        return 0.0
    if isinstance(raw_value, (int, float)):
        return float(raw_value)
    text = str(raw_value).strip()
    if not text:
        return 0.0
    compact = text.replace(" ", "").replace(",", ".")
    upper = compact.upper()
    mnum = re.match(r"^(-?\d+(?:\.\d+)?)", compact)
    if not mnum:
        return 0.0
    val = float(mnum.group(1))
    if upper.endswith("KA"):
        return val * 1000.0
    if upper.endswith("MA"):
        return val / 1000.0
    if upper.endswith("A"):
        return val
    return val


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


def _parse_range(raw: str) -> tuple[int, int]:
    match = re.match(r"^\s*(\d+)\s*-\s*(\d+)\s*$", raw)
    if not match:
        raise ValueError(f"Range non valido: {raw!r}. Usa formato start-stop, es. 165-350.")
    start, stop = int(match.group(1)), int(match.group(2))
    if start > stop:
        raise ValueError(f"Range non valido: start > stop ({start}-{stop}).")
    return start, stop


def _parse_optional_range(raw: str) -> tuple[int, int] | None:
    value = raw.strip().upper()
    if value in {"", "NONE", "N/A", "-"}:
        return None
    return _parse_range(raw)


def parse_excel_cell_ref(ref: str) -> tuple[int, int]:
    """
    Converte un riferimento stile Excel (es. C2, AA10) in (riga, colonna) 1-based.
    """
    text = ref.strip().upper()
    match = re.match(r"^([A-Z]+)(\d+)$", text)
    if not match:
        raise ValueError(f"Riferimento cella non valido: {ref!r}. Usa formato come C2 o AA10.")
    letters, row_s = match.groups()
    row = int(row_s)
    col = 0
    for ch in letters:
        col = col * 26 + (ord(ch) - ord("A") + 1)
    return row, col


def row_is_comando_vite(col_a_value: object) -> bool:
    text = str(col_a_value or "").strip()
    return bool(_COMANDO_VITE_RE.search(text))


def extract_comando_vite_summary_rows(sheet) -> list[dict]:
    """
    Cerca righe con 'COMANDO VITE x' in colonna A; kW e A da col. D e F.
    Ordina per x; massimo _MAX_VITE_ROWS_SMISTAMENTO righe (slot righe 7-10 nello smistamento).
    """
    used_range = sheet.UsedRange
    first_row = used_range.Row
    last_row = first_row + used_range.Rows.Count - 1
    found: dict[int, tuple[float, float]] = {}

    for row in range(first_row, last_row + 1):
        col_a = sheet.Cells(row, "A").Value
        text = str(col_a or "").strip()
        m = _COMANDO_VITE_RE.search(text)
        if not m:
            continue
        num = int(m.group(1))
        if num in found:
            continue
        kw_val = parse_measure(sheet.Cells(row, "D").Value, "kW")
        amp_val = parse_measure(sheet.Cells(row, "F").Value, "A")
        if kw_val == 0.0 and amp_val == 0.0:
            kw_val = parse_numeric_loose(sheet.Cells(row, "D").Value)
            amp_val = parse_numeric_loose(sheet.Cells(row, "F").Value)
        found[num] = (kw_val, amp_val)

    ordered_nums = sorted(found.keys())[:_MAX_VITE_ROWS_SMISTAMENTO]
    out: list[dict] = []
    for num in ordered_nums:
        kw, amp = found[num]
        out.append(
            {
                "title": f"Motore vite {num}",
                "kw": normalize_output_number(kw),
                "amp": normalize_output_number(amp),
                "rif": "V",
                "is_tagliapasta": False,
                "group_key": f"V:vite:{num}",
            }
        )
    return out


def read_static_row_measures(sheet, cell_ref: str) -> tuple[float, float]:
    """kW e A sulla stessa riga del riferimento: colonne D e F (come righe impianto)."""
    row, _col = parse_excel_cell_ref(cell_ref)
    d_val = sheet.Cells(row, "D").Value
    f_val = sheet.Cells(row, "F").Value
    kw = parse_measure(d_val, "kW")
    amp = parse_measure(f_val, "A")
    if abs(kw) <= 1e-9 and abs(amp) <= 1e-9:
        kw = parse_numeric_loose(d_val)
        amp = parse_numeric_loose(f_val)
    return normalize_output_number(kw), normalize_output_number(amp)


def read_unified_line_config(config_path: Path, line_type_key: str) -> tuple[list[ConfigRow], SectionSplitRules, LineDisplayConfig]:
    """
    Profilo unico per la scelta menu: sezione #LINE_DISPLAY_<CHIAVE> (CHIAVE = prima colonna di #LINE_TYPES).

    Contiene tutto cio che serve:
      - Impianto: @SPLIT_LABEL, @LABEL_*, @RANGE_*
      - Fogli da elaborare:
          * se c'e almeno una @MAP_SHEET_GROUP -> solo i fogli citati nella MAP (ordine di prima apparizione);
          * altrimenti -> righe "Titolo";"NomeFoglio"
      - Smistamento: @OVERRIDE_*, @MAP_SHEET_GROUP, @PRIMARY_RANGE_SHEET, @SUMMARY_ROW*, ...
    """
    if not config_path.exists():
        raise FileNotFoundError(f"File configurazione non trovato: {config_path}")

    target = f"LINE_DISPLAY_{line_type_key.strip().upper()}"

    quoted_sheet_rows: list[ConfigRow] = []
    split_label = "Stenditrice"
    pressa_label = "Movimenti Pressa"
    recupero_polveri_label = "Movimenti Recupero Polveri"
    movimenti_linea_label = "Movimenti Linea"
    movimenti_selezione_sili_label = "Movimenti selezione prodotto e sili"
    pressa_range = (165, 350)
    secondary_range = (351, 449)
    recupero_polveri_range: tuple[int, int] | None = None
    movimenti_linea_range: tuple[int, int] | None = (450, 860)
    movimenti_selezione_sili_range: tuple[int, int] | None = None

    label_overrides: dict[str, str] = {}
    static_summary_rows: list[tuple[str | None, str, str]] = []
    primary_range_sheet: str | None = None
    offsheet_pressa_label: str | None = None
    offsheet_secondario_label: str | None = None
    sheet_group_labels: dict[tuple[str, str], str] = {}
    map_sheet_names_ordered: list[str] = []
    map_sheet_names_seen: set[str] = set()

    current_section: str | None = None
    for raw_line in config_path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line:
            continue

        if line.startswith("#"):
            sec = _config_section_name_from_line(line)
            if sec is not None:
                current_section = sec
            continue

        if current_section != target:
            continue

        if line.startswith("@MAP_SHEET_GROUP="):
            payload = line[len("@MAP_SHEET_GROUP=") :].strip()
            parts = payload.split(";", 2)
            if len(parts) < 3 or not parts[0] or not parts[1] or not parts[2]:
                log(f"[WARN] @MAP_SHEET_GROUP ignorata (servono foglio;gruppo;titolo): {line}")
                continue
            sheet_name_raw, gkey_raw, title = parts[0].strip(), parts[1].strip().lower(), parts[2].strip()
            if gkey_raw not in VALID_GROUP_KEYS:
                log(
                    f"[WARN] @MAP_SHEET_GROUP gruppo non valido {gkey_raw!r} "
                    f"(attesi {sorted(VALID_GROUP_KEYS)}): {line}"
                )
                continue
            sheet_group_labels[(sheet_name_raw.upper(), gkey_raw)] = title
            sn_norm = sheet_name_raw.strip()
            key_u = sn_norm.upper()
            if key_u not in map_sheet_names_seen:
                map_sheet_names_seen.add(key_u)
                map_sheet_names_ordered.append(sn_norm)
            continue

        if line.startswith("@SUMMARY_ROW_SHEET="):
            payload = line[len("@SUMMARY_ROW_SHEET=") :].strip()
            parts = [p.strip() for p in payload.split(";", 2)]
            if len(parts) < 3 or not parts[0] or not parts[1] or not parts[2]:
                log(f"[WARN] @SUMMARY_ROW_SHEET ignorata (servono foglio;cella;titolo): {line}")
                continue
            sheet_filter, ref, title = parts[0], parts[1], parts[2]
            static_summary_rows.append((sheet_filter, ref, title))
            continue

        if line.startswith("@SUMMARY_ROW="):
            payload = line[len("@SUMMARY_ROW=") :].strip()
            ref_part, sep, title_part = payload.partition(";")
            if not sep:
                log(f"[WARN] @SUMMARY_ROW ignorata (manca ';' e titolo): {line}")
                continue
            ref = ref_part.strip()
            title = title_part.strip()
            if not ref or not title:
                log(f"[WARN] @SUMMARY_ROW ignorata (riferimento o titolo vuoto): {line}")
                continue
            static_summary_rows.append((None, ref, title))
            continue

        match_new = re.match(r'^"([^"]+)";"([^"]+)"$', line)
        match_old = re.match(r'^"([^"]+)";"([^"]+)";([01])$', line)
        if match_new or match_old:
            if match_new:
                _title, sheet_name = match_new.groups()
            else:
                _title, sheet_name, _legacy = match_old.groups()
            quoted_sheet_rows.append(ConfigRow(sheet_name=sheet_name.strip()))
            continue

        if not line.startswith("@"):
            log(f"[WARN] Riga ignorata in #{target}: {line}")
            continue

        key, sep, value = line.partition("=")
        if not sep:
            log(f"[WARN] Direttiva ignorata (formato non valido): {line}")
            continue
        key_u = key.strip().upper()
        value_stripped = value.strip().strip('"').strip("'")

        if key_u == "@SPLIT_LABEL":
            if value_stripped:
                split_label = value_stripped
            continue
        if key_u == "@LABEL_PRESSA":
            if value_stripped:
                pressa_label = value_stripped
            continue
        if key_u == "@LABEL_RECUPERO_POLVERI":
            if value_stripped:
                recupero_polveri_label = value_stripped
            continue
        if key_u == "@LABEL_MOVIMENTI_LINEA":
            if value_stripped:
                movimenti_linea_label = value_stripped
            continue
        if key_u == "@LABEL_MOVIMENTI_SELEZIONE_SILI":
            if value_stripped:
                movimenti_selezione_sili_label = value_stripped
            continue
        if key_u == "@RANGE_PRESSA":
            pressa_range = _parse_range(value_stripped)
            continue
        if key_u == "@RANGE_FORMATRICE":
            secondary_range = _parse_range(value_stripped)
            continue
        if key_u == "@RANGE_RECUPERO_POLVERI":
            recupero_polveri_range = _parse_optional_range(value_stripped)
            continue
        if key_u == "@RANGE_MOVIMENTI_LINEA":
            movimenti_linea_range = _parse_optional_range(value_stripped)
            continue
        if key_u == "@RANGE_MOVIMENTI_SELEZIONE_SILI":
            movimenti_selezione_sili_range = _parse_optional_range(value_stripped)
            continue

        if key_u == "@PRIMARY_RANGE_SHEET":
            primary_range_sheet = value_stripped if value_stripped else None
            continue
        if key_u == "@OFFSHEET_PRESSA_LABEL":
            offsheet_pressa_label = value_stripped if value_stripped else None
            continue
        if key_u == "@OFFSHEET_SECONDARIO_LABEL":
            offsheet_secondario_label = value_stripped if value_stripped else None
            continue

        if key_u == "@OVERRIDE_PRESSA":
            if value_stripped:
                label_overrides["pressa"] = value_stripped
            continue
        if key_u == "@OVERRIDE_RECUPERO_POLVERI":
            if value_stripped:
                label_overrides["recupero_polveri"] = value_stripped
            continue
        if key_u == "@OVERRIDE_SECONDARIO":
            if value_stripped:
                label_overrides["secondario"] = value_stripped
            continue
        if key_u == "@OVERRIDE_MOVIMENTI_LINEA":
            if value_stripped:
                label_overrides["movimenti_linea"] = value_stripped
            continue
        if key_u == "@OVERRIDE_MOVIMENTI_SELEZIONE_SILI":
            if value_stripped:
                label_overrides["movimenti_selezione_sili"] = value_stripped
            continue

        log(f"[WARN] Direttiva sconosciuta in #{target} ignorata: {line}")

    if map_sheet_names_ordered:
        if quoted_sheet_rows:
            log(
                "[INFO] Sezione con @MAP_SHEET_GROUP: le righe \"...\";\"NomeFoglio\" "
                "non definiscono fogli extra (usa solo la MAP per l'elenco fogli)."
            )
        rows = [ConfigRow(sheet_name=s) for s in map_sheet_names_ordered]
    else:
        deduped: list[ConfigRow] = []
        dedupe_seen: set[str] = set()
        for r in quoted_sheet_rows:
            u = r.sheet_name.strip().upper()
            if u not in dedupe_seen:
                dedupe_seen.add(u)
                deduped.append(r)
        rows = deduped

    rules = SectionSplitRules(
        split_label=split_label,
        pressa_label=pressa_label,
        recupero_polveri_label=recupero_polveri_label,
        movimenti_linea_label=movimenti_linea_label,
        movimenti_selezione_sili_label=movimenti_selezione_sili_label,
        pressa_range=pressa_range,
        secondary_range=secondary_range,
        recupero_polveri_range=recupero_polveri_range,
        movimenti_linea_range=movimenti_linea_range,
        movimenti_selezione_sili_range=movimenti_selezione_sili_range,
    )
    display_cfg = LineDisplayConfig(
        label_overrides=label_overrides,
        static_summary_rows=static_summary_rows,
        primary_range_sheet=primary_range_sheet,
        offsheet_pressa_label=offsheet_pressa_label,
        offsheet_secondario_label=offsheet_secondario_label,
        sheet_group_labels=sheet_group_labels,
    )
    return rows, rules, display_cfg


def summary_label_for_group(
    sheet_name: str,
    group_key: str,
    default_label: str,
    display_cfg: LineDisplayConfig,
) -> str:
    # Fuori dal foglio "primario" (@PRIMARY_RANGE_SHEET), Pressa/Stenditrice da range usano solo
    # @OFFSHEET_* (cosi @OVERRIDE_SECONDARIO sul primario non sovrascrive Sezionatore ... altrove).
    if display_cfg.primary_range_sheet:
        primary = display_cfg.primary_range_sheet.strip().upper()
        if sheet_name.strip().upper() != primary:
            if group_key == "pressa" and display_cfg.offsheet_pressa_label:
                return display_cfg.offsheet_pressa_label
            if group_key == "secondario" and display_cfg.offsheet_secondario_label:
                return display_cfg.offsheet_secondario_label
    return display_cfg.label_overrides.get(group_key, default_label)


def title_for_sheet_group(
    sheet_name: str,
    group_key: str,
    default_label: str,
    display_cfg: LineDisplayConfig,
) -> str:
    """Titolo riga smistamento: priorità @MAP_SHEET_GROUP, poi logica @PRIMARY_RANGE_SHEET / @OVERRIDE_*."""
    sn = sheet_name.strip().upper()
    mapped = display_cfg.sheet_group_labels.get((sn, group_key))
    if mapped:
        return mapped
    return summary_label_for_group(sheet_name, group_key, default_label, display_cfg)


def read_line_type_options(config_path: Path) -> list[LineTypeOption]:
    if not config_path.exists():
        raise FileNotFoundError(f"File configurazione non trovato: {config_path}")

    options: list[LineTypeOption] = []
    current_section: str | None = None
    for raw_line in config_path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line:
            continue

        if line.startswith("#"):
            sec = _config_section_name_from_line(line)
            if sec is not None:
                current_section = sec
            continue

        if current_section != "LINE_TYPES":
            continue

        match = re.match(r'^"([^"]+)";"([^"]+)"(?:;"([^"]+)")?\s*$', line)
        if not match:
            log(f"[WARN] Riga tipologia linea ignorata (formato non valido): {line}")
            continue

        key, label, _legacy_section = match.groups()
        options.append(
            LineTypeOption(
                key=key.strip().upper(),
                label=label.strip(),
            )
        )

    if options:
        return options

    # Fallback compatibilita: sezione non presente nel file configurazione.
    return [
        LineTypeOption("PASTA_LUNGA_4_VITI", "Pasta Lunga 4 Viti"),
        LineTypeOption("PASTA_LUNGA_2_VITI", "Pasta Lunga 2 Viti"),
        LineTypeOption("PASTA_CORTA_2_VITI", "Pasta Corta 2 Viti"),
        LineTypeOption("CTA", "CTA"),
    ]


def choose_section(config_path: Path) -> LineTypeOption:
    options = read_line_type_options(config_path)
    log("Seleziona la tipologia linea:")
    for idx, option in enumerate(options, start=1):
        log(f"  {idx}) {option.label}")

    valid_choices = "/".join(str(i) for i in range(1, len(options) + 1))
    choice = input(f"Scelta [{valid_choices}]: ").strip()
    try:
        selected = options[int(choice) - 1]
    except (ValueError, IndexError):
        raise ValueError("Scelta non valida.")
    return selected


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


_SHEET_LETTERS_THEN_DIGITS_RE = re.compile(r"^([A-Za-z]+)(\d+)$")


def sheet_uses_range_grouping(sheet_name: str) -> bool:
    """
    True -> insert_totals_rows sul foglio Excel + smistamento da gruppi M/R (eventuale filtro MAP).

    False per nome <Lettere><numero> con numero > 1 (es. B2, C4): non si applicano mai i gruppi/range M/R;
    una riga smistamento con somma delle colonne D e F sul UsedRange (titolo dal primo @MAP_SHEET_GROUP).
    """
    s = sheet_name.strip()
    if re.fullmatch(r"[A-Za-z]+", s):
        return True
    m = _SHEET_LETTERS_THEN_DIGITS_RE.fullmatch(s)
    if m:
        return int(m.group(2)) <= 1
    return True


def compute_sheet_grand_total(sheet) -> tuple[float, float]:
    """
    Somma colonna D (kW) e F (A) sul UsedRange senza classificazione M/R.
    Esclude righe con A che inizia per TOTALE e COMANDO VITE.
    Righe senza alcun contributo numerico in D ne in F vengono saltate (riduce rumore da UsedRange esteso).
    """
    used_range = sheet.UsedRange
    first_row = used_range.Row
    last_row = first_row + used_range.Rows.Count - 1
    total_kw = 0.0
    total_amp = 0.0
    for row in range(first_row, last_row + 1):
        col_a = normalize_text(sheet.Cells(row, "A").Value)
        if col_a.startswith("TOTALE"):
            continue
        if row_is_comando_vite(sheet.Cells(row, "A").Value):
            continue
        raw_d = sheet.Cells(row, "D").Value
        raw_f = sheet.Cells(row, "F").Value
        d_blank = raw_d is None or str(raw_d).strip() == ""
        f_blank = raw_f is None or str(raw_f).strip() == ""
        if d_blank and f_blank:
            continue
        total_kw += parse_grand_total_kw(raw_d)
        total_amp += parse_grand_total_amp(raw_f)
    return normalize_output_number(total_kw), normalize_output_number(total_amp)


def first_map_title_for_sheet(display_cfg: LineDisplayConfig, sheet_name: str) -> str | None:
    """Primo titolo @MAP_SHEET_GROUP per quel nome foglio (ordine nel file di config)."""
    sn = sheet_name.strip().upper()
    for (sh, _gk), title in display_cfg.sheet_group_labels.items():
        if sh == sn:
            return title
    return None


def map_group_keys_for_sheet(display_cfg: LineDisplayConfig, sheet_name: str) -> set[str] | None:
    """
    Per fogli il cui nome e' solo lettere (es. B, V): None -> si emettono tutti i gruppi con totale
    non zero; @MAP_SHEET_GROUP serve solo ai titoli (title_for_sheet_group), non a limitare le righe.

    Per fogli con suffisso numerico nel nome (es. B1): se esiste @MAP_SHEET_GROUP per quel foglio,
    restituisce i soli group_key da elencare (es. B1 -> solo pressa). Se non c'e MAP per quel nome ->
    None (tutti i gruppi non zero).
    """
    raw = sheet_name.strip()
    if re.fullmatch(r"[A-Za-z]+", raw):
        return None

    sn = sheet_name.strip().upper()
    keys: set[str] = set()
    for (sh, gk), _title in display_cfg.sheet_group_labels.items():
        if sh == sn:
            keys.add(gk)
    if not keys:
        return None
    return keys


def extract_three_digit_code(value: object) -> int | None:
    """
    Estrae il codice a tre cifre dal formato:
    <lettera><codice_tre_cifre><codice_una_cifra>
    Esempio: M3651 -> 365
    """
    if value is None:
        return None
    text = str(value).strip().upper()
    match = re.match(r"^[MR](\d{3})\d$", text)
    if not match:
        return None
    return int(match.group(1))


def in_range(code: int | None, bounds: tuple[int, int]) -> bool:
    if code is None:
        return False
    start, stop = bounds
    return start <= code <= stop


def group_definitions(split_cfg: SectionSplitRules) -> list[dict]:
    defs: list[dict] = [
        # Priorita: recupero prima di pressa (range sovrapposti possibili, es. Cous Cous 191-199)
        {"key": "recupero_polveri", "label": split_cfg.recupero_polveri_label, "bounds": split_cfg.recupero_polveri_range},
        {"key": "pressa", "label": split_cfg.pressa_label, "bounds": split_cfg.pressa_range},
        {"key": "secondario", "label": split_cfg.split_label, "bounds": split_cfg.secondary_range},
        {"key": "movimenti_linea", "label": split_cfg.movimenti_linea_label, "bounds": split_cfg.movimenti_linea_range},
        {
            "key": "movimenti_selezione_sili",
            "label": split_cfg.movimenti_selezione_sili_label,
            "bounds": split_cfg.movimenti_selezione_sili_range,
        },
    ]
    return [entry for entry in defs if entry["bounds"] is not None]


def compute_totals_for_sheet(
    sheet,
    split_cfg: SectionSplitRules,
) -> tuple[dict[str, dict[str, float]], list[UnmatchedRow]]:
    used_range = sheet.UsedRange
    first_row = used_range.Row
    last_row = first_row + used_range.Rows.Count - 1
    if last_row < 8:
        last_row = 8

    totals: dict[str, dict[str, float]] = {}
    for entry in group_definitions(split_cfg):
        totals[entry["key"]] = {"kw": 0.0, "amp": 0.0}
    unmatched_rows: list[UnmatchedRow] = []

    for row in range(8, last_row + 1):
        col_a = normalize_text(sheet.Cells(row, "A").Value)
        if col_a.startswith("TOTALE"):
            continue

        if row_is_comando_vite(sheet.Cells(row, "A").Value):
            continue

        kw_val = parse_measure(sheet.Cells(row, "D").Value, "kW")
        amp_val = parse_measure(sheet.Cells(row, "F").Value, "A")
        if kw_val == 0.0 and amp_val == 0.0:
            continue

        code = extract_three_digit_code(sheet.Cells(row, "A").Value)
        matched_key: str | None = None
        for entry in group_definitions(split_cfg):
            if in_range(code, entry["bounds"]):
                matched_key = entry["key"]
                break

        if matched_key is None:
            unmatched_rows.append(
                UnmatchedRow(
                    row=row,
                    raw_code=str(sheet.Cells(row, "A").Value or "").strip(),
                    parsed_code=code,
                    kw=kw_val,
                    amp=amp_val,
                )
            )
            continue

        totals[matched_key]["kw"] += kw_val
        totals[matched_key]["amp"] += amp_val

    return totals, unmatched_rows


def insert_totals_rows(
    sheet,
    split_cfg: SectionSplitRules,
) -> tuple[dict[str, dict[str, float]], list[UnmatchedRow]]:
    split_label = split_cfg.split_label
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

    totals, unmatched_rows = compute_totals_for_sheet(sheet, split_cfg)
    pressa = totals.get("pressa", {"kw": 0.0, "amp": 0.0})
    secondario = totals.get("secondario", {"kw": 0.0, "amp": 0.0})

    sheet.Range("D4").Value = normalize_output_number(pressa["kw"])
    sheet.Range("F4").Value = normalize_output_number(pressa["amp"])

    sheet.Range("D5").Value = normalize_output_number(secondario["kw"])
    sheet.Range("F5").Value = normalize_output_number(secondario["amp"])

    return totals, unmatched_rows


def create_summary_file(
    excel_app,
    output_path: Path,
    summary_rows: list[dict],
    vite_rows: list[dict] | None = None,
) -> None:
    wb = excel_app.Workbooks.Add()
    try:
        ws = wb.Worksheets(1)
        ws.Name = "Smistamento potenza"
        start_row = _SMISTAMENTO_MAIN_TABLE_ROW
        vite_rows = vite_rows or []

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

        # Riga 5: intestazioni tabella
        ws.Cells(5, 2).Value = "kW"
        ws.Cells(5, 3).Value = "A"
        ws.Cells(5, 4).Value = "sez."

        # Righe 5-6: griglia intestazione
        for row in (5, 6):
            rng = ws.Range(ws.Cells(row, 1), ws.Cells(row, 6))
            rng.Borders.LineStyle = XL_CONTINUOUS
            rng.Borders.Weight = XL_MEDIUM

        # Righe 7-10: Motore vite 1..4 dal foglio V (testo COMANDO VITE x)
        vite_last_row = _SMISTAMENTO_VITE_START_ROW - 1
        if vite_rows:
            r = _SMISTAMENTO_VITE_START_ROW
            for item in vite_rows[:_MAX_VITE_ROWS_SMISTAMENTO]:
                ws.Cells(r, 1).Value = item["title"]
                ws.Cells(r, 2).Value = normalize_output_number(item["kw"])
                ws.Cells(r, 3).Value = normalize_output_number(item["amp"])
                ws.Cells(r, 4).Value = item["rif"]
                ws.Cells(r, 5).Value = ""
                ws.Cells(r, 6).Value = ""
                vite_last_row = r
                r += 1
            vite_rng = ws.Range(
                ws.Cells(_SMISTAMENTO_VITE_START_ROW, 1),
                ws.Cells(vite_last_row, 6),
            )
            vite_rng.Borders.LineStyle = XL_CONTINUOUS
            vite_rng.Borders.Weight = XL_MEDIUM
            vite_rng.VerticalAlignment = -4107
            vite_rng.HorizontalAlignment = 1

        # Righe vuote 7-10 se non ci sono viti: solo contorno come prima (area riservata)
        if not vite_rows:
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
            ws.Cells(row, 2).Value = normalize_output_number(kw)
            ws.Cells(row, 3).Value = normalize_output_number(amp)
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
        total_kw = sum(item["kw"] for item in summary_rows) + sum(item["kw"] for item in vite_rows)
        total_amp = sum(item["amp"] for item in summary_rows) + sum(item["amp"] for item in vite_rows)
        total_row = max(24, last_written_row + 3)
        ws.Cells(total_row, 1).Value = "TOTALE"
        ws.Cells(total_row, 2).Value = normalize_output_number(total_kw)
        ws.Cells(total_row, 3).Value = normalize_output_number(total_amp)
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
        selected_line_type = choose_section(config_path)
        config_rows, split_cfg, display_cfg = read_unified_line_config(config_path, selected_line_type.key)
        if not config_rows:
            raise RuntimeError(
                f"Nessuna riga foglio in #LINE_DISPLAY_{selected_line_type.key}: "
                'aggiungi righe "Titolo";"NomeFoglio".'
            )

        input_file = find_input_workbook(base_dir)
        output_input_name = derive_input_output_name(input_file)
        output_input_path = base_dir / output_input_name
        shutil.copy2(input_file, output_input_path)

        output_summary_name = derive_summary_output_name(input_file)
        output_summary_path = base_dir / output_summary_name

        log(f"File input trovato: {input_file.name}")
        log(f"Creo copia output input: {output_input_path.name}")
        log(f"Tipologia linea: {selected_line_type.label} ({selected_line_type.key})")
        log(f"Profilo configurazione: #LINE_DISPLAY_{selected_line_type.key}")

        pythoncom.CoInitialize()
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        summary_rows: list[dict] = []
        vite_summary_rows: list[dict] = []
        try:
            wb_input = excel.Workbooks.Open(str(output_input_path))
            try:
                for cfg in config_rows:
                    sheet = get_sheet(wb_input, cfg.sheet_name)
                    if sheet is None:
                        log(f"[INFO] Foglio '{cfg.sheet_name}' non presente: salto.")
                        continue

                    log(f"[INFO] Elaboro foglio '{cfg.sheet_name}'...")
                    use_groups = sheet_uses_range_grouping(cfg.sheet_name)
                    unmatched_rows: list[UnmatchedRow] = []

                    if use_groups:
                        totals_by_group, unmatched_rows = insert_totals_rows(sheet, split_cfg)
                        if cfg.sheet_name.strip().lower() == "v":
                            vite_summary_rows = extract_comando_vite_summary_rows(sheet)
                            if vite_summary_rows:
                                log(
                                    f"[INFO] Foglio V: {len(vite_summary_rows)} righe COMANDO VITE -> smistamento righe "
                                    f"{_SMISTAMENTO_VITE_START_ROW}-{_SMISTAMENTO_VITE_START_ROW + len(vite_summary_rows) - 1}."
                                )

                        map_only_groups = map_group_keys_for_sheet(display_cfg, cfg.sheet_name)
                        for entry in group_definitions(split_cfg):
                            key = entry["key"]
                            if map_only_groups is not None and key not in map_only_groups:
                                continue
                            title = title_for_sheet_group(cfg.sheet_name, key, entry["label"], display_cfg)
                            kw = totals_by_group.get(key, {}).get("kw", 0.0)
                            amp = totals_by_group.get(key, {}).get("amp", 0.0)
                            if abs(kw) <= 1e-9 and abs(amp) <= 1e-9:
                                continue
                            summary_rows.append(
                                {
                                    "title": title,
                                    "kw": normalize_output_number(kw),
                                    "amp": normalize_output_number(amp),
                                    "rif": cfg.sheet_name,
                                    "is_tagliapasta": key == "secondario",
                                    "group_key": f"{cfg.sheet_name}:{key}",
                                }
                            )
                    else:
                        unmatched_rows = []
                        log(
                            f"[INFO] Foglio '{cfg.sheet_name}': suffisso numerico > 1 -> "
                            "nessun raggruppamento M/R; somma colonne D e F sul UsedRange; titolo da MAP."
                        )
                        kw_g, amp_g = compute_sheet_grand_total(sheet)
                        title_u = first_map_title_for_sheet(display_cfg, cfg.sheet_name)
                        if not title_u:
                            title_u = cfg.sheet_name
                            log(
                                f"[WARN] Foglio '{cfg.sheet_name}' senza @MAP_SHEET_GROUP: "
                                "titolo smistamento = nome foglio."
                            )
                        summary_rows.append(
                            {
                                "title": title_u,
                                "kw": normalize_output_number(kw_g),
                                "amp": normalize_output_number(amp_g),
                                "rif": cfg.sheet_name,
                                "is_tagliapasta": False,
                                "group_key": f"{cfg.sheet_name}:sheet_total",
                            }
                        )

                    for sheet_filter, cell_ref, static_title in display_cfg.static_summary_rows:
                        if sheet_filter is not None:
                            if sheet_filter.strip().lower() != cfg.sheet_name.strip().lower():
                                continue
                        try:
                            kw_s, amp_s = read_static_row_measures(sheet, cell_ref)
                        except ValueError as exc:
                            log(f"[WARN] Foglio '{cfg.sheet_name}': {exc}")
                            continue
                        summary_rows.append(
                            {
                                "title": static_title,
                                "kw": normalize_output_number(kw_s),
                                "amp": normalize_output_number(amp_s),
                                "rif": cfg.sheet_name,
                                "is_tagliapasta": False,
                                "group_key": f"{cfg.sheet_name}:cell:{cell_ref.upper()}",
                            }
                        )

                    if use_groups and unmatched_rows:
                        log(
                            f"[WARN] Foglio '{cfg.sheet_name}': {len(unmatched_rows)} righe con valori non classificate "
                            "perche fuori da tutti i range configurati."
                        )
                        for item in unmatched_rows:
                            if row_is_comando_vite(item.raw_code):
                                continue
                            log(
                                f"       - riga {item.row}: codice='{item.raw_code}', parsed={item.parsed_code}, "
                                f"kW={item.kw:.2f}, A={item.amp:.2f}"
                            )

                if not vite_summary_rows:
                    v_only = get_sheet(wb_input, "V")
                    if v_only is not None:
                        vite_summary_rows = extract_comando_vite_summary_rows(v_only)
                        if vite_summary_rows:
                            log("[INFO] Foglio V presente: COMANDO VITE estratti anche senza 'V' nella lista fogli del profilo.")

                wb_input.Save()
                log(f"[OK] File output input aggiornato: {output_input_path.name}")
            finally:
                wb_input.Close(SaveChanges=False)

            create_summary_file(excel, output_summary_path, summary_rows, vite_rows=vite_summary_rows)
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
