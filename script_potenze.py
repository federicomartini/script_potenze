from __future__ import annotations

import re
import shutil
import sys
from dataclasses import dataclass
from pathlib import Path

import pythoncom
import win32com.client as win32


CONFIG_FILE_NAME = "configurazione_schede.txt"

# Foglio "V": testo tipo "COMANDO VITE 1" in colonna A oppure B (case insensitive).
_COMANDO_VITE_RE = re.compile(r"COMANDO\s+VITE\s*(\d+)", re.IGNORECASE)
# Foglio "TR" (Cous Cous): COMANDO TRITURATORE PRODOTTO FINE / GROSSO in A o B.
_TRITURATORE_FINE_RE = re.compile(
    r"COMANDO\s+TRITURATORE\s+PRODOTTO\s+FINE\b",
    re.IGNORECASE,
)
_TRITURATORE_GROSSO_RE = re.compile(
    r"COMANDO\s+TRITURATORE\s+PRODOTTO\s+GROSSO\b",
    re.IGNORECASE,
)
_LINE_TYPE_VITE_COUNT_RE = re.compile(r"_(\d+)_VITI$")
_SMISTAMENTO_VITE_START_ROW = 7
_SMISTAMENTO_MAIN_TABLE_ROW = 11
# File output (*_Output.xlsx): colonna con gruppo M/R assegnato dalla classificazione a range.
_MR_GROUP_OUTPUT_COL = "P"
# Tabella totali M/R: riga 4 = intestazione gialla (Totali / kW / A); righe successive = un totale per gruppo;
# poi riga vuota; poi intestazione tabella dati originale (D/F con kW e A); poi i dati.
_MR_TOTALS_BLOCK_FIRST_ROW = 4
# Solo colonne A..F per i totali (il resto non serve).
_MR_TOTALS_LAST_COL = 6
# Foglio con un solo totale: righe inserite sopra l'intestazione (Totali + Totale + vuoto), senza riga target assoluta.
_GRAND_TOTAL_BLOCK_ROWS = 3

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
# Giallo foglio Excel (BGR 00FFFF) per la prima riga della tabella totali.
COLOR_YELLOW_EXCEL = 65535
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
    cous_cous_mode: bool = False


@dataclass
class SectionSplitRules:
    split_label: str
    pressa_label: str
    recupero_polveri_label: str
    movimenti_linea_label: str
    movimenti_selezione_sili_label: str
    pressa_range: tuple[int, int]
    secondary_range: tuple[int, int] | None
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


def row_is_comando_vite(col_a_value: object, col_b_value: object | None = None) -> bool:
    """True se COMANDO VITE x e in colonna A oppure in colonna B."""
    if _COMANDO_VITE_RE.search(str(col_a_value or "").strip()):
        return True
    if col_b_value is not None and _COMANDO_VITE_RE.search(str(col_b_value or "").strip()):
        return True
    return False


def row_is_comando_trituratore(col_a_value: object, col_b_value: object | None = None) -> bool:
    """True se COMANDO TRITURATORE PRODOTTO FINE/GROSSO in colonna A o B (Cous Cous, foglio TR)."""
    text_a = str(col_a_value or "").strip()
    text_b = str(col_b_value or "").strip() if col_b_value is not None else ""
    for t in (text_a, text_b):
        if _TRITURATORE_FINE_RE.search(t) or _TRITURATORE_GROSSO_RE.search(t):
            return True
    return False


def row_skips_mr_data_row(col_a_value: object, col_b_value: object | None = None) -> bool:
    """Righe comando vite o trituratore: escluse da somme M/R e da totali foglio."""
    return row_is_comando_vite(col_a_value, col_b_value) or row_is_comando_trituratore(
        col_a_value, col_b_value
    )


def line_type_is_cous_cous(line_key: str) -> bool:
    """Linee Cous Cous: foglio TR, non il modello COMANDO VITE su V."""
    k = line_key.strip().upper()
    return "COUS_COU" in k or k.startswith("COUS_COUS")


_COUS_COUS_PRESSA_WORD_RE = re.compile(r"(?i)\bPressa\b")


def cous_cous_replace_pressa_in_text(text: str) -> str:
    """Cous Cous: ovunque compare la parola 'Pressa' (titoli, MAP, override) -> 'Impastatrici'."""
    return _COUS_COUS_PRESSA_WORD_RE.sub("Impastatrici", text)


def _cous_cous_finalize_label(display_cfg: LineDisplayConfig, text: str) -> str:
    if display_cfg.cous_cous_mode:
        return cous_cous_replace_pressa_in_text(text)
    return text


def expected_vite_count_for_line_type(line_key: str) -> int | None:
    """
    Numero viti atteso dal nome tipologia (es. PASTA_LUNGA_4_VITI -> 4).
    CTA: una sola vite (COMANDO VITE 1) anche senza suffisso _N_VITI nel nome.
    None per Cous Cous, ecc.
    """
    if line_type_is_cous_cous(line_key):
        return None
    k = line_key.strip().upper()
    if k == "CTA" or k.startswith("CTA_"):
        return 1
    m = _LINE_TYPE_VITE_COUNT_RE.search(k)
    return int(m.group(1)) if m else None


def collect_comando_vite_rows(sheet) -> dict[int, tuple[float, float]]:
    """
    Tutte le righe COMANDO VITE x nel foglio (testo in colonna A o B); kW/A da col. D e F.
    Chiave duplicata -> errore.
    """
    used_range = sheet.UsedRange
    first_row = used_range.Row
    last_row = first_row + used_range.Rows.Count - 1
    found: dict[int, tuple[float, float]] = {}

    for row in range(first_row, last_row + 1):
        col_a = sheet.Cells(row, "A").Value
        col_b = sheet.Cells(row, "B").Value
        text_a = str(col_a or "").strip()
        text_b = str(col_b or "").strip()
        m = _COMANDO_VITE_RE.search(text_b) or _COMANDO_VITE_RE.search(text_a)
        if not m:
            continue
        num = int(m.group(1))
        if num in found:
            raise RuntimeError(
                f"Foglio V: COMANDO VITE {num} presente piu di una volta (righe duplicate per lo stesso numero)."
            )
        kw_val = parse_grand_total_kw(sheet.Cells(row, "D").Value)
        amp_val = parse_grand_total_amp(sheet.Cells(row, "F").Value)
        if kw_val == 0.0 and amp_val == 0.0:
            kw_val = parse_measure(sheet.Cells(row, "D").Value, "kW")
            amp_val = parse_measure(sheet.Cells(row, "F").Value, "A")
            if kw_val == 0.0 and amp_val == 0.0:
                kw_val = parse_numeric_loose(sheet.Cells(row, "D").Value)
                amp_val = parse_numeric_loose(sheet.Cells(row, "F").Value)
        found[num] = (kw_val, amp_val)

    return found


def validate_vite_sheet(expected_n: int, found: dict[int, tuple[float, float]]) -> None:
    """
    Fallisce se il foglio V non contiene esattamente `expected_n` viti, come da tipologia linea,
    con COMANDO VITE 1 … COMANDO VITE expected_n (una riga per numero, nessun altro indice).
    `expected_n` deriva dal nome tipologia (es. PASTA_LUNGA_4_VITI -> 4).
    """
    n_file = len(found)
    if n_file != expected_n:
        raise RuntimeError(
            f"Foglio V: servono esattamente {expected_n} viti (tipologia linea scelta), "
            f"nel file risultano {n_file} righe COMANDO VITE numerate distinte. "
            f"Numeri presenti: {sorted(found.keys())}."
        )
    required = set(range(1, expected_n + 1))
    got = set(found.keys())
    if got != required:
        raise RuntimeError(
            f"Foglio V: con {expected_n} viti i numeri COMANDO VITE devono essere esattamente "
            f"da 1 a {expected_n} (una riga per ogni numero). "
            f"Trovati invece: {sorted(got)}."
        )


def vite_dict_to_summary_rows(found: dict[int, tuple[float, float]]) -> list[dict]:
    out: list[dict] = []
    for num in sorted(found.keys()):
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


def collect_comando_trituratore_rows(sheet) -> dict[str, tuple[float, float]]:
    """
    Foglio TR (Cous Cous): righe COMANDO TRITURATORE PRODOTTO FINE e GROSSO (testo in A o B); kW/A in D e F.
    Chiavi 'fine' e 'grosso'; duplicati -> errore.
    """
    used_range = sheet.UsedRange
    first_row = used_range.Row
    last_row = first_row + used_range.Rows.Count - 1
    found: dict[str, tuple[float, float]] = {}

    for row in range(first_row, last_row + 1):
        col_a = sheet.Cells(row, "A").Value
        col_b = sheet.Cells(row, "B").Value
        text_a = str(col_a or "").strip()
        text_b = str(col_b or "").strip()
        kind: str | None = None
        if _TRITURATORE_FINE_RE.search(text_b) or _TRITURATORE_FINE_RE.search(text_a):
            kind = "fine"
        elif _TRITURATORE_GROSSO_RE.search(text_b) or _TRITURATORE_GROSSO_RE.search(text_a):
            kind = "grosso"
        if not kind:
            continue
        if kind in found:
            raise RuntimeError(
                f"Foglio TR: COMANDO TRITURATORE PRODOTTO {kind.upper()} presente piu di una volta."
            )
        kw_val = parse_grand_total_kw(sheet.Cells(row, "D").Value)
        amp_val = parse_grand_total_amp(sheet.Cells(row, "F").Value)
        if kw_val == 0.0 and amp_val == 0.0:
            kw_val = parse_measure(sheet.Cells(row, "D").Value, "kW")
            amp_val = parse_measure(sheet.Cells(row, "F").Value, "A")
            if kw_val == 0.0 and amp_val == 0.0:
                kw_val = parse_numeric_loose(sheet.Cells(row, "D").Value)
                amp_val = parse_numeric_loose(sheet.Cells(row, "F").Value)
        found[kind] = (kw_val, amp_val)

    return found


def validate_trituratore_sheet(found: dict[str, tuple[float, float]]) -> None:
    """Esattamente una riga FINE e una GROSSO sul foglio TR."""
    need = {"fine", "grosso"}
    got = set(found.keys())
    if got != need:
        raise RuntimeError(
            "Foglio TR (Cous Cous): servono esattamente due righe, "
            "COMANDO TRITURATORE PRODOTTO FINE e COMANDO TRITURATORE PRODOTTO GROSSO "
            f"(testo in colonna A o B). Trovati tipi: {sorted(got)}."
        )


def trituratore_dict_to_summary_rows(found: dict[str, tuple[float, float]]) -> list[dict]:
    titles = {"fine": "Trituratore prodotto fine", "grosso": "Trituratore prodotto grosso"}
    out: list[dict] = []
    for kind in ("fine", "grosso"):
        kw, amp = found[kind]
        out.append(
            {
                "title": titles[kind],
                "kw": normalize_output_number(kw),
                "amp": normalize_output_number(amp),
                "rif": "TR",
                "is_tagliapasta": False,
                "group_key": f"TR:trituratore:{kind}",
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
      - Impianto: @SPLIT_LABEL, @LABEL_*, @RANGE_* (@RANGE_PRESSA sempre obbligatorio: start-stop).
        @RANGE_RECUPERO_POLVERI opzionale: se interseca @RANGE_PRESSA, i codici nel recupero vanno al
        gruppo recupero (priorità sulla Pressa). @RANGE_CUOCITORE opzionale sul gruppo secondario
        (stessa chiave «secondario» di @RANGE_FORMATRICE; se valorizzato ha precedenza su FORMATRICE).
      - Fogli da elaborare:
          * se c'e almeno una @MAP_SHEET_GROUP o @MAP_SHEET -> solo i fogli citati (ordine di prima apparizione);
          * altrimenti -> righe "Titolo";"NomeFoglio"
      - Smistamento: @OVERRIDE_*, @MAP_SHEET_GROUP, @PRIMARY_RANGE_SHEET, @SUMMARY_ROW*, ...
    """
    if not config_path.exists():
        raise FileNotFoundError(f"File configurazione non trovato: {config_path}")

    target = f"LINE_DISPLAY_{line_type_key.strip().upper()}"

    quoted_sheet_rows: list[ConfigRow] = []
    split_label = "Stenditrice"
    if line_type_is_cous_cous(line_type_key):
        pressa_label = "Movimenti Impastatrici"
    else:
        pressa_label = "Movimenti Pressa"
    recupero_polveri_label = "Movimenti Recupero Polveri"
    movimenti_linea_label = "Movimenti Linea"
    movimenti_selezione_sili_label = "Movimenti selezione prodotto e sili"
    pressa_range = (165, 350)
    secondary_range = (351, 449)
    recupero_polveri_range = None
    movimenti_linea_range = (450, 860)
    movimenti_selezione_sili_range = None
    cuocitore_range_input = None

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

        if line.startswith("@MAP_SHEET="):
            payload = line[len("@MAP_SHEET=") :].strip()
            sheet_name_raw = payload.split(";")[0].strip()
            if not sheet_name_raw:
                log(f"[WARN] @MAP_SHEET ignorata (nome foglio vuoto): {line}")
                continue
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
            try:
                pressa_range = _parse_range(value_stripped)
            except ValueError as exc:
                raise RuntimeError(
                    f"#{target}: @RANGE_PRESSA e' obbligatorio: intervallo start-stop (es. 165-350); "
                    "NONE non e' ammesso."
                ) from exc
            continue
        if key_u == "@RANGE_FORMATRICE":
            secondary_range = _parse_optional_range(value_stripped)
            continue
        if key_u == "@RANGE_CUOCITORE":
            cuocitore_range_input = _parse_optional_range(value_stripped)
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

    if cuocitore_range_input is not None:
        if secondary_range is not None and secondary_range != cuocitore_range_input:
            log(
                "[WARN] @RANGE_CUOCITORE e @RANGE_FORMATRICE entrambi valorizzati e diversi: "
                "per il gruppo secondario (@SPLIT_LABEL) si applica solo @RANGE_CUOCITORE."
            )
        secondary_range = cuocitore_range_input

    if map_sheet_names_ordered:
        if quoted_sheet_rows:
            log(
                "[INFO] Sezione con @MAP_SHEET_GROUP / @MAP_SHEET: le righe \"...\";\"NomeFoglio\" "
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

    is_cc = line_type_is_cous_cous(line_type_key)
    if is_cc:
        split_label = cous_cous_replace_pressa_in_text(split_label)
        pressa_label = cous_cous_replace_pressa_in_text(pressa_label)
        recupero_polveri_label = cous_cous_replace_pressa_in_text(recupero_polveri_label)
        movimenti_linea_label = cous_cous_replace_pressa_in_text(movimenti_linea_label)
        movimenti_selezione_sili_label = cous_cous_replace_pressa_in_text(movimenti_selezione_sili_label)
        if offsheet_pressa_label:
            offsheet_pressa_label = cous_cous_replace_pressa_in_text(offsheet_pressa_label)
        if offsheet_secondario_label:
            offsheet_secondario_label = cous_cous_replace_pressa_in_text(offsheet_secondario_label)
        label_overrides = {k: cous_cous_replace_pressa_in_text(v) for k, v in label_overrides.items()}
        static_summary_rows = [
            (sf, ref, cous_cous_replace_pressa_in_text(tit)) for sf, ref, tit in static_summary_rows
        ]
        sheet_group_labels = {
            k: cous_cous_replace_pressa_in_text(v) for k, v in sheet_group_labels.items()
        }

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
        cous_cous_mode=is_cc,
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
                return _cous_cous_finalize_label(display_cfg, display_cfg.offsheet_pressa_label)
            if group_key == "secondario" and display_cfg.offsheet_secondario_label:
                return _cous_cous_finalize_label(display_cfg, display_cfg.offsheet_secondario_label)
    return _cous_cous_finalize_label(
        display_cfg, display_cfg.label_overrides.get(group_key, default_label)
    )


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
        title = mapped
    else:
        title = summary_label_for_group(sheet_name, group_key, default_label, display_cfg)
    title = sezionatore_label_on_letter1_sheet(sheet_name, title)
    return _cous_cous_finalize_label(display_cfg, title)


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
        LineTypeOption("PASTA_LUNGA_1_VITI", "Pasta Lunga 1 Vite"),
        LineTypeOption("PASTA_LUNGA_2_VITI", "Pasta Lunga 2 Viti"),
        LineTypeOption("PASTA_CORTA_1_VITI", "Pasta Corta 1 Vite"),
        LineTypeOption("PASTA_CORTA_2_VITI", "Pasta Corta 2 Viti"),
        LineTypeOption("CTA", "CTA"),
        LineTypeOption("COUS_COUS", "Cous Cous"),
    ]


def choose_section(config_path: Path) -> LineTypeOption:
    options = read_line_type_options(config_path)
    log("Seleziona la tipologia linea:")
    for idx, option in enumerate(options, start=1):
        log(f"  {idx}) {option.label}")

    valid_choices = "/".join(str(i) for i in range(1, len(options) + 1))
    n = len(options)
    choice = input(f"Scelta [{valid_choices}]: ").strip()
    if not choice:
        raise ValueError(
            f"Scelta non valida: nessun valore inserito. "
            f"Digita un numero intero da 1 a {n} per la tipologia linea."
        )
    try:
        idx = int(choice)
    except ValueError:
        raise ValueError(
            f"Scelta non valida: «{choice}» non e' un numero intero. "
            f"Usa un valore da 1 a {n}."
        )
    if idx < 1 or idx > n:
        raise ValueError(
            f"Scelta non valida: {idx} e' fuori intervallo. "
            f"Sono definite solo le opzioni da 1 a {n}."
        )
    return options[idx - 1]


def excel_two_decimal_format(decimal_separator: str) -> str:
    """Pattern formato Excel per due decimali visibili (`decimal_separator` = '.' o ',')."""
    return "0,00" if decimal_separator == "," else "0.00"


def choose_decimal_separator() -> str:
    log("Separatore decimale nei file Excel generati (solo visualizzazione):")
    log("  1) Punto (.)")
    log("  2) Virgola (,)")
    choice = input("Scelta [1/2]: ").strip()
    if choice == "2":
        return ","
    if choice == "1" or choice == "":
        return "."
    raise ValueError(
        f"Scelta non valida: «{choice}» non e' ammesso per il separatore decimale. "
        "Digita 1 (punto), 2 (virgola) oppure Invio per usare il punto."
    )


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


def sheet_name_suffix_numeric_equals_one(sheet_name: str) -> bool:
    """True se il nome e' <Lettere><cifre> e la parte numerica e' esattamente 1 (es. B1, AA1; no B11)."""
    m = _SHEET_LETTERS_THEN_DIGITS_RE.fullmatch(sheet_name.strip())
    return bool(m and int(m.group(2)) == 1)


def sezionatore_label_on_letter1_sheet(sheet_name: str, label: str) -> str:
    """
    Su fogli <Lettere>1 (es. A1, B1, AA1; esclusi A11, B2) sostituisce il prefisso «Movimenti » con «Sezionatore »
    (stessa regola dello smistamento). Output Excel: righe «Totale …» e colonna P gruppi.
    """
    if not sheet_name_suffix_numeric_equals_one(sheet_name):
        return label
    if label.startswith("Movimenti "):
        return "Sezionatore " + label[len("Movimenti ") :]
    if label == "Movimenti":
        return "Sezionatore"
    return label


def sheet_is_multi_group_mr_name(sheet_name: str) -> bool:
    """
    Fogli dove tutti i gruppi definiti dai range M/R vanno sempre considerati (nessun filtro su @MAP_SHEET_GROUP):
    - nome con sole lettere (es. A, B, V, TR);
    - nome <Lettere><cifre> con parte numerica ESATTAMENTE 1 (es. B1, AA1).
    Esclusi suffissi diversi da 1 (es. B2, B11, C9): li si tratta come fogli 'numerati' a parte.
    """
    s = sheet_name.strip()
    if re.fullmatch(r"[A-Za-z]+", s):
        return True
    m = _SHEET_LETTERS_THEN_DIGITS_RE.fullmatch(s)
    if m:
        return int(m.group(2)) == 1
    return False


def sheet_uses_range_grouping(sheet_name: str) -> bool:
    """
    True -> insert_totals_rows sul foglio Excel + smistamento da gruppi M/R (eventuale filtro MAP).

    False per nome <Lettere><numero> con numero diverso da 1 (es. B2, B11): somma D/F senza split range;
    titolo smistamento dal primo @MAP_SHEET_GROUP se presente.
    Altri nomi (non matching Lettere+Cifre) restano True come prima.
    """
    s = sheet_name.strip()
    if re.fullmatch(r"[A-Za-z]+", s):
        return True
    m = _SHEET_LETTERS_THEN_DIGITS_RE.fullmatch(s)
    if m:
        return int(m.group(2)) == 1
    return True


def compute_sheet_grand_total(sheet) -> tuple[float, float]:
    """
    Somma colonna D (kW) e F (A) sul UsedRange senza classificazione M/R.
    Esclude righe con A «Totali» (tabella totali), le sole righe «Totale» generate dallo script
    (A uguale a «Totale» o «Totale <gruppo>»), e righe COMANDO VITE / COMANDO TRITURATORE (A o B).
    Righe senza alcun contributo numerico in D ne in F vengono saltate (riduce rumore da UsedRange esteso).
    """
    used_range = sheet.UsedRange
    first_row = used_range.Row
    last_row = first_row + used_range.Rows.Count - 1
    total_kw = 0.0
    total_amp = 0.0
    for row in range(first_row, last_row + 1):
        col_a = normalize_text(sheet.Cells(row, "A").Value)
        if col_a == "TOTALI":
            continue
        if is_script_generated_mr_totale_label(col_a):
            continue
        if row_skips_mr_data_row(sheet.Cells(row, "A").Value, sheet.Cells(row, "B").Value):
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
            return _cous_cous_finalize_label(display_cfg, title)
    return None


def map_group_keys_for_sheet(display_cfg: LineDisplayConfig, sheet_name: str) -> set[str] | None:
    """
    None -> smistamento elenca tutti i gruppi con totale non zero (classificazione dai @RANGE_*).
    @MAP_SHEET_GROUP resta solo per titoli custom (title_for_sheet_group), non per limitare i gruppi,
    su fogli solo-lettera e su fogli <Lettere>1 (es. A, B, B1; esclusi B11, B2).

    Per gli altri nomi foglio (es. C3): se esiste MAP per quel nome, si filtra ai soli group_key indicati;
    se non c'e MAP -> None (tutti i gruppi non zero).
    """
    if sheet_is_multi_group_mr_name(sheet_name):
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
        # Prima recupero, poi pressa: se @RANGE_RECUPERO_POLVERI e' dentro (o interseca) @RANGE_PRESSA,
        # i codici nel recupero devono contare come recupero, non come pressa.
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


def group_display_label_for_sheet(sheet_name: str, group_key: str, base_label: str) -> str:
    """
    Etichetta mostrata per un gruppo su un dato foglio (Output Excel e default smistamento).
    Foglio C vs CC: stesso @RANGE_MOVIMENTI_SELEZIONE_SILI, descrizioni distinte.
    """
    if group_key != "movimenti_selezione_sili":
        out = base_label
    else:
        sn = sheet_name.strip().upper()
        if sn == "CC":
            out = "Movimenti Scarico Sili e Confezionamento"
        elif sn == "C":
            out = "Mov. Selezione Prodotto e Carico Sili"
        else:
            out = base_label
    return sezionatore_label_on_letter1_sheet(sheet_name, out)


def _mr_strict_blank_separator_a_to_f(sheet, row: int) -> bool:
    """Riga separatrice tra tabella totali e tabella dati: celle A..F tutte vuote."""
    for c in range(1, _MR_TOTALS_LAST_COL + 1):
        v = sheet.Cells(row, c).Value
        if v is not None and str(v).strip() != "":
            return False
    return True


def is_script_generated_mr_totale_label(col_a_normalized: str) -> bool:
    """
    True solo per le righe «Totale» inserite dallo script (da rimuovere in strip / da escludere da somme):
    - A esattamente «Totale» → normalizzato «TOTALE»
    - «Totale <gruppo>» → «TOTALE » seguito da testo (es. «TOTALE PRESSA»)

    False per testi di riga originale del foglio tipo «Totale: 12,5» («TOTALE: ...») o altre varianti.
    """
    if not col_a_normalized:
        return False
    if col_a_normalized == "TOTALE":
        return True
    return col_a_normalized.startswith("TOTALE ") and len(col_a_normalized) > len("TOTALE ")


def mr_insert_anchor_row_above_first_data(sheet, first_data_row: int) -> int:
    """
    Prima riga da cui inserire il blocco di righe così che intestazione e dati restino uniti.
    Se tra intestazione e primo dato c'è una riga vuota in A..F, l'inserimento deve partire
    dall'intestazione (non dalla riga vuota), altrimenti l'intestazione non scende e kW/A finiscono sulla riga sbagliata.
    """
    r = first_data_row - 1
    while r >= 1 and _mr_strict_blank_separator_a_to_f(sheet, r):
        r -= 1
    return max(1, r)


def mr_data_header_row_above(
    sheet,
    first_data_row: int,
    *,
    min_row: int,
) -> int:
    """
    Riga di intestazione tabella dati: la prima non vuota in A..F sopra il primo dato, saltando
    solo righe completamente vuote in A..F (stesso gap possibile tra intestazione e dati).
    min_row: non risalire sopra questa riga (subito sotto il blocco totali inserito).
    """
    r = first_data_row - 1
    while r > min_row and _mr_strict_blank_separator_a_to_f(sheet, r):
        r -= 1
    return max(min_row, r)


def _row_is_mr_block_blank_separator(sheet, row: int) -> bool:
    """Riga vuota tra intestazione kW/A e tabella dati (nessun codice M/R, nessun contributo)."""
    col_a = normalize_text(sheet.Cells(row, "A").Value)
    if col_a == "TOTALI":
        return False
    if is_script_generated_mr_totale_label(col_a):
        return False
    if row_skips_mr_data_row(sheet.Cells(row, "A").Value, sheet.Cells(row, "B").Value):
        return False
    d_cell = str(sheet.Cells(row, "D").Value or "").strip().upper()
    if d_cell == "KW":
        return False
    if extract_three_digit_code(sheet.Cells(row, "A").Value) is not None:
        return False
    kw_val = parse_measure(sheet.Cells(row, "D").Value, "kW")
    amp_val = parse_measure(sheet.Cells(row, "F").Value, "A")
    return kw_val == 0.0 and amp_val == 0.0


def replace_dim_potenza_corrente_attr_labels(
    sheet,
    row_top: int,
    row_bottom: int,
    *,
    max_col: int = 24,
    only_columns_d_f: bool = False,
) -> None:
    """
    Sostituisce le etichette potenza/corrente con 'kW' e 'A'.
    - Se only_columns_d_f=True: modifica solo le colonne D ed F (intestazione tabella dati spostata),
      così il resto della riga resta identico al foglio originale.
    - Se False: scansiona fino a max_col (es. più righe sopra la tabella) e applica anche i casi
      schema con (104) / (101) oltre D/F se presenti.
    """
    for r in range(row_top, row_bottom + 1):
        col_range = (4, 6) if only_columns_d_f else range(1, max_col + 1)
        for c in col_range:
            v = sheet.Cells(r, c).Value
            if v is None:
                continue
            t = str(v).strip()
            if not t:
                continue
            low = t.lower()
            if c == 4 and "potenza" in low:
                sheet.Cells(r, c).Value = "kW"
            elif c == 6 and "corrente" in low:
                sheet.Cells(r, c).Value = "A"
            elif not only_columns_d_f:
                if "potenza" in low and "(104)" in low:
                    sheet.Cells(r, c).Value = "kW"
                elif "corrente" in low and "(101)" in low:
                    sheet.Cells(r, c).Value = "A"


def clear_sheet_row_cells(sheet, row: int, *, from_col: int = 1, to_col: int = 16) -> None:
    for c in range(from_col, to_col + 1):
        sheet.Cells(row, c).ClearContents()


def clear_row_contents_right_of_col(sheet, row: int, *, from_col: int) -> None:
    """Svuota celle da from_col in poi (es. oltre F per righe totali)."""
    for c in range(from_col, 50):
        sheet.Cells(row, c).ClearContents()


def normalize_data_kw_amp_numeric_cells(
    sheet,
    first_data_row: int,
    last_row: int,
    *,
    number_format: str,
) -> None:
    """
    Nella tabella dati: in D e F solo numeri (unità solo in intestazione), usando parse_grand_total_*.
    """
    for row in range(first_data_row, last_row + 1):
        col_a = normalize_text(sheet.Cells(row, "A").Value)
        if is_script_generated_mr_totale_label(col_a):
            continue
        if row_skips_mr_data_row(sheet.Cells(row, "A").Value, sheet.Cells(row, "B").Value):
            continue
        raw_d = sheet.Cells(row, "D").Value
        raw_f = sheet.Cells(row, "F").Value
        d_blank = raw_d is None or str(raw_d).strip() == ""
        f_blank = raw_f is None or str(raw_f).strip() == ""
        if d_blank and f_blank:
            continue
        if not d_blank:
            kw = parse_grand_total_kw(raw_d)
            sheet.Cells(row, "D").Value = normalize_output_number(kw)
            sheet.Cells(row, "D").NumberFormat = number_format
        if not f_blank:
            amp = parse_grand_total_amp(raw_f)
            sheet.Cells(row, "F").Value = normalize_output_number(amp)
            sheet.Cells(row, "F").NumberFormat = number_format


def find_first_mr_data_row(sheet) -> int:
    """Prima riga dati impianto M/R (codice in A e/o valori in D/F), esclusi Totali e COMANDO VITE / TRITURATORE."""
    used_range = sheet.UsedRange
    first_row = used_range.Row
    last_row = first_row + used_range.Rows.Count - 1
    start_scan = max(4, first_row)
    for row in range(start_scan, last_row + 1):
        col_a = normalize_text(sheet.Cells(row, "A").Value)
        if col_a == "TOTALI":
            continue
        if is_script_generated_mr_totale_label(col_a):
            continue
        if row_skips_mr_data_row(sheet.Cells(row, "A").Value, sheet.Cells(row, "B").Value):
            continue
        raw_d_hdr = str(sheet.Cells(row, "D").Value or "")
        raw_f_hdr = str(sheet.Cells(row, "F").Value or "")
        if "attributo a scelta del simbolo" in raw_d_hdr.lower() or "attributo a scelta del simbolo" in raw_f_hdr.lower():
            continue
        d_cell = str(sheet.Cells(row, "D").Value or "").strip().upper()
        if d_cell == "KW":
            continue
        if extract_three_digit_code(sheet.Cells(row, "A").Value) is not None:
            return row
        kw_val = parse_measure(sheet.Cells(row, "D").Value, "kW")
        amp_val = parse_measure(sheet.Cells(row, "F").Value, "A")
        if kw_val != 0.0 or amp_val != 0.0:
            return row
    return 8


def strip_existing_mr_totals_block(sheet) -> None:
    """
    Rimuove la tabella totali inserita dallo script: eventuale riga gialla 'Totali', le righe
    'Totale ...', e una riga vuota A..F subito sotto, senza toccare l'intestazione tabella dati.
    La riga 'Totali' può non essere alla riga 4 (fogli con totale unico).
    """
    try:
        ur = sheet.UsedRange
    except Exception:
        return
    last = min(ur.Row + ur.Rows.Count - 1, ur.Row + 400)

    r_tot: int | None = None
    for rr in range(4, last + 1):
        if normalize_text(sheet.Cells(rr, "A").Value) == "TOTALI":
            r_tot = rr
            break

    if r_tot is not None:
        r = r_tot
        sheet.Rows(r).Delete()
        last -= 1
        while r <= last:
            if is_script_generated_mr_totale_label(normalize_text(sheet.Cells(r, "A").Value)):
                sheet.Rows(r).Delete()
                last -= 1
                continue
            break
        if r <= last and _mr_strict_blank_separator_a_to_f(sheet, r):
            sheet.Rows(r).Delete()
        return

    r = 4
    if r > last:
        return
    while r <= last:
        if is_script_generated_mr_totale_label(normalize_text(sheet.Cells(r, "A").Value)):
            sheet.Rows(r).Delete()
            last -= 1
            continue
        break
    if r <= last and _mr_strict_blank_separator_a_to_f(sheet, r):
        sheet.Rows(r).Delete()


def find_first_grand_data_row(sheet) -> int:
    """
    Prima riga dati motore (fogli senza split gruppi): stessa regola di find_first_mr_data_row.
    Non usare «prima cella D/F non vuota»: l'intestazione tabella ha spesso 0.00 in D/F e verrebbe
    scambiata per dato, spostando l'inserimento totali sulle righe titolo sopra l'intestazione vera.
    """
    return find_first_mr_data_row(sheet)


def insert_grand_total_row(sheet, *, number_format: str = "0.00") -> None:
    """
    Foglio senza split per gruppi: inserisce esattamente tre righe vuote sopra l'intestazione della
    tabella dati (ancora allineata al primo dato M/R, non alle celle 0.00 dell'intestazione). Scrive
    Totali (giallo), Totale con somme, riga vuota; sulla riga intestazione tabella (ricavata come per
    i fogli multi-gruppo) imposta D/F = kW e A in grassetto. Nessuna riga target fissa (niente shift cumulativi).
    """
    strip_existing_mr_totals_block(sheet)
    fd_before = find_first_grand_data_row(sheet)
    kw_g, amp_g = compute_sheet_grand_total(sheet)
    shift_top = mr_insert_anchor_row_above_first_data(sheet, fd_before)

    block = _GRAND_TOTAL_BLOCK_ROWS
    sheet.Rows(f"{shift_top}:{shift_top + block - 1}").Insert()

    first_data_row = find_first_grand_data_row(sheet)
    min_header = shift_top + block
    header_row = mr_data_header_row_above(sheet, first_data_row, min_row=min_header)
    tr = shift_top + 1

    try:
        sheet.Rows(first_data_row).Copy()
        sheet.Rows(tr).PasteSpecial(XL_PASTE_FORMATS)
    except Exception:
        pass

    y = shift_top
    sheet.Cells(y, "A").Value = "Totali"
    sheet.Cells(y, "D").Value = "kW"
    sheet.Cells(y, "F").Value = "A"
    sheet.Cells(y, "B").ClearContents()
    sheet.Cells(y, "C").ClearContents()
    sheet.Cells(y, "E").ClearContents()
    y_rng = sheet.Range(sheet.Cells(y, 1), sheet.Cells(y, _MR_TOTALS_LAST_COL))
    y_rng.Interior.Pattern = 1
    y_rng.Interior.Color = COLOR_YELLOW_EXCEL
    y_rng.Font.Bold = True
    clear_row_contents_right_of_col(sheet, y, from_col=_MR_TOTALS_LAST_COL + 1)

    sheet.Cells(tr, "A").Value = "Totale"
    sheet.Cells(tr, "D").Value = normalize_output_number(kw_g)
    sheet.Cells(tr, "F").Value = normalize_output_number(amp_g)
    sheet.Cells(tr, "D").NumberFormat = number_format
    sheet.Cells(tr, "F").NumberFormat = number_format
    clear_row_contents_right_of_col(sheet, tr, from_col=_MR_TOTALS_LAST_COL + 1)

    clear_sheet_row_cells(sheet, shift_top + 2)

    replace_dim_potenza_corrente_attr_labels(
        sheet, header_row, header_row, only_columns_d_f=True
    )
    if str(sheet.Cells(header_row, "D").Value or "").strip().lower() != "kw":
        sheet.Cells(header_row, "D").Value = "kW"
    if str(sheet.Cells(header_row, "F").Value or "").strip().lower() != "a":
        sheet.Cells(header_row, "F").Value = "A"
    sheet.Cells(header_row, "D").Font.Bold = True
    sheet.Cells(header_row, "F").Font.Bold = True

    try:
        ur = sheet.UsedRange
        last_mr = ur.Row + ur.Rows.Count - 1
    except Exception:
        last_mr = first_data_row + 500
    normalize_data_kw_amp_numeric_cells(
        sheet, first_data_row, last_mr, number_format=number_format
    )


def remove_sheets_not_in_config(workbook, config_rows: list[ConfigRow], line_type_key: str) -> None:
    """Elimina dal workbook i fogli il cui nome non compare nella configurazione della linea scelta.
    Il foglio V non viene mai rimosso (COMANDO VITE). Il foglio TR non viene rimosso per linee Cous Cous."""
    allowed = {str(c.sheet_name).strip().upper() for c in config_rows}
    idx = workbook.Worksheets.Count
    while idx >= 1:
        ws = workbook.Worksheets(idx)
        name_u = str(ws.Name).strip().upper()
        if name_u == "V":
            idx -= 1
            continue
        if name_u == "TR" and line_type_is_cous_cous(line_type_key):
            idx -= 1
            continue
        if name_u not in allowed:
            log(f"[INFO] Rimosso foglio non in configurazione: {ws.Name}")
            try:
                ws.Delete()
            except Exception as exc:
                log(f"[WARN] Impossibile eliminare il foglio '{ws.Name}': {exc}")
        idx -= 1


def mr_data_row_totals_classification(
    sheet,
    row: int,
    split_cfg: SectionSplitRules,
    groups: list[dict],
) -> tuple[bool, str | None, float, float]:
    """
    (include, matched_group_key, kw, amp).
    include=False: la riga non entra nei totali.
    include=True e matched_group_key None: contributo non assegnato ad alcun @RANGE_*.
    """
    col_a = normalize_text(sheet.Cells(row, "A").Value)
    if col_a == "TOTALI":
        return False, None, 0.0, 0.0
    if is_script_generated_mr_totale_label(col_a):
        return False, None, 0.0, 0.0
    if row_skips_mr_data_row(sheet.Cells(row, "A").Value, sheet.Cells(row, "B").Value):
        return False, None, 0.0, 0.0

    kw_val = parse_measure(sheet.Cells(row, "D").Value, "kW")
    amp_val = parse_measure(sheet.Cells(row, "F").Value, "A")
    if kw_val == 0.0 and amp_val == 0.0:
        return False, None, 0.0, 0.0

    code = extract_three_digit_code(sheet.Cells(row, "A").Value)
    for entry in groups:
        if in_range(code, entry["bounds"]):
            return True, entry["key"], kw_val, amp_val
    return True, None, kw_val, amp_val


def compute_totals_for_sheet(
    sheet,
    split_cfg: SectionSplitRules,
) -> tuple[dict[str, dict[str, float]], list[UnmatchedRow]]:
    used_range = sheet.UsedRange
    first_row = used_range.Row
    last_row = first_row + used_range.Rows.Count - 1
    first_data_row = find_first_mr_data_row(sheet)
    if last_row < first_data_row:
        last_row = first_data_row

    groups = group_definitions(split_cfg)
    totals: dict[str, dict[str, float]] = {}
    for entry in groups:
        totals[entry["key"]] = {"kw": 0.0, "amp": 0.0}
    unmatched_rows: list[UnmatchedRow] = []

    for row in range(first_data_row, last_row + 1):
        incl, matched_key, kw_val, amp_val = mr_data_row_totals_classification(sheet, row, split_cfg, groups)
        if not incl:
            continue

        if matched_key is None:
            code = extract_three_digit_code(sheet.Cells(row, "A").Value)
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


def clear_mr_group_annotation_column(sheet) -> None:
    """Pulisce la colonna P nel UsedRange (annotazioni gruppo M/R)."""
    try:
        ur = sheet.UsedRange
    except Exception:
        return
    if ur is None:
        return
    top = ur.Row
    last = ur.Row + ur.Rows.Count - 1
    sheet.Range(sheet.Cells(top, _MR_GROUP_OUTPUT_COL), sheet.Cells(last, _MR_GROUP_OUTPUT_COL)).ClearContents()


def annotate_mr_group_column_p(
    sheet,
    split_cfg: SectionSplitRules,
    *,
    sheet_name: str,
    cous_cous_line: bool,
) -> None:
    """
    Colonna P: etichetta gruppo per ogni riga dati che concorre ai totali (anche con un solo gruppo a range).
    Usata sia sui fogli con righe «Totale …» per gruppo sia su fogli tipo B2 (totale unico), usando gli stessi @RANGE_*.
    """
    clear_mr_group_annotation_column(sheet)
    groups = group_definitions(split_cfg)
    if not groups:
        return

    label_by_key: dict[str, str] = {}
    for entry in groups:
        lbl = group_display_label_for_sheet(sheet_name, entry["key"], entry["label"])
        if cous_cous_line:
            lbl = cous_cous_replace_pressa_in_text(lbl)
        label_by_key[entry["key"]] = lbl

    used_range = sheet.UsedRange
    last_row = used_range.Row + used_range.Rows.Count - 1
    first_data_row = find_first_mr_data_row(sheet)

    for row in range(first_data_row, last_row + 1):
        incl, matched_key, _, _ = mr_data_row_totals_classification(sheet, row, split_cfg, groups)
        if not incl:
            continue
        cell = sheet.Cells(row, _MR_GROUP_OUTPUT_COL)
        if matched_key is not None:
            cell.Value = label_by_key[matched_key]
        else:
            cell.Value = "Non classificato"


def insert_totals_rows(
    sheet,
    split_cfg: SectionSplitRules,
    *,
    sheet_name: str,
    number_format: str = "0.00",
    cous_cous_line: bool = False,
) -> tuple[dict[str, dict[str, float]], list[UnmatchedRow], int]:
    """
    Inserisce sopra la tabella dati: (1) riga gialla con Totali / kW / A in A,D,F; (2) una riga
    per gruppo con descrizione in A e totali in D,F; (3) riga vuota A..F; (4) intestazione
    originale della tabella (solo D/F sostituiti con kW e A); poi i dati con numeri puri in D e F.
    Terzo valore: shift righe rispetto al layout prima dell'operazione (per log).
    """
    strip_existing_mr_totals_block(sheet)
    fd_before = find_first_mr_data_row(sheet)
    totals, unmatched_rows = compute_totals_for_sheet(sheet, split_cfg)
    entries = group_definitions(split_cfg)
    if not entries:
        clear_mr_group_annotation_column(sheet)
        return totals, unmatched_rows, 0

    n = len(entries)
    t0 = _MR_TOTALS_BLOCK_FIRST_ROW
    # t0 = intestazione gialla; t0+1..t0+n = totali; t0+n+1 = vuoto; t0+n+2 = intestazione dati; dati da t0+n+3
    target_data = t0 + n + 3

    # Inserire sopra il primo dato sposta solo i dati, non l'intestazione: la tabella (intestazione+dati)
    # va spostata inserendo righe a partire dall'intestazione (salta righe vuote A..F tra intestazione e dato).
    shift_top = mr_insert_anchor_row_above_first_data(sheet, fd_before)

    if fd_before < target_data:
        ins = target_data - fd_before
        if ins > 0:
            sheet.Rows(f"{shift_top}:{shift_top + ins - 1}").Insert()
    elif fd_before > target_data:
        dels = fd_before - target_data
        if dels > 0:
            sheet.Rows(f"{target_data}:{fd_before - 1}").Delete()

    fd_check = find_first_mr_data_row(sheet)
    if fd_check != target_data:
        log(
            f"[WARN] Layout totali: prima riga dati attesa {target_data}, trovata {fd_check}; "
            "verificare righe sopra la tabella dati."
        )

    first_data_row = fd_check
    min_header = t0 + n + 2
    header_row = mr_data_header_row_above(sheet, first_data_row, min_row=min_header)
    blank_row = t0 + n + 1

    try:
        sheet.Rows(first_data_row).Copy()
        for i in range(n):
            sheet.Rows(t0 + 1 + i).PasteSpecial(XL_PASTE_FORMATS)
    except Exception:
        pass

    y = t0
    sheet.Cells(y, "A").Value = "Totali"
    sheet.Cells(y, "D").Value = "kW"
    sheet.Cells(y, "F").Value = "A"
    sheet.Cells(y, "B").ClearContents()
    sheet.Cells(y, "C").ClearContents()
    sheet.Cells(y, "E").ClearContents()
    y_rng = sheet.Range(sheet.Cells(y, 1), sheet.Cells(y, _MR_TOTALS_LAST_COL))
    y_rng.Interior.Pattern = 1
    y_rng.Interior.Color = COLOR_YELLOW_EXCEL
    y_rng.Font.Bold = True
    clear_row_contents_right_of_col(sheet, y, from_col=_MR_TOTALS_LAST_COL + 1)

    for i, entry in enumerate(entries):
        r = t0 + 1 + i
        tw = totals.get(entry["key"], {"kw": 0.0, "amp": 0.0})
        row_label = group_display_label_for_sheet(sheet_name, entry["key"], entry["label"])
        label_a = f"Totale {row_label}"
        if cous_cous_line:
            label_a = cous_cous_replace_pressa_in_text(label_a)
        sheet.Cells(r, "A").Value = label_a
        sheet.Cells(r, "D").Value = normalize_output_number(tw["kw"])
        sheet.Cells(r, "F").Value = normalize_output_number(tw["amp"])
        sheet.Cells(r, "D").NumberFormat = number_format
        sheet.Cells(r, "F").NumberFormat = number_format
        clear_row_contents_right_of_col(sheet, r, from_col=_MR_TOTALS_LAST_COL + 1)

    replace_dim_potenza_corrente_attr_labels(
        sheet, header_row, header_row, only_columns_d_f=True
    )
    if str(sheet.Cells(header_row, "D").Value or "").strip().lower() != "kw":
        sheet.Cells(header_row, "D").Value = "kW"
    if str(sheet.Cells(header_row, "F").Value or "").strip().lower() != "a":
        sheet.Cells(header_row, "F").Value = "A"

    clear_sheet_row_cells(sheet, blank_row)

    try:
        ur = sheet.UsedRange
        last_mr = ur.Row + ur.Rows.Count - 1
    except Exception:
        last_mr = first_data_row + 500
    normalize_data_kw_amp_numeric_cells(
        sheet, first_data_row, last_mr, number_format=number_format
    )

    row_shift = find_first_mr_data_row(sheet) - fd_before

    annotate_mr_group_column_p(
        sheet,
        split_cfg,
        sheet_name=sheet_name,
        cous_cous_line=cous_cous_line,
    )

    return totals, unmatched_rows, row_shift


def create_summary_file(
    excel_app,
    output_path: Path,
    summary_rows: list[dict],
    vite_rows: list[dict] | None = None,
    *,
    number_format: str = "0.00",
) -> None:
    wb = excel_app.Workbooks.Add()
    try:
        ws = wb.Worksheets(1)
        ws.Name = "Smistamento potenza"
        vite_rows = vite_rows or []
        if vite_rows:
            start_row = _SMISTAMENTO_VITE_START_ROW + len(vite_rows) + 1
        else:
            start_row = _SMISTAMENTO_MAIN_TABLE_ROW

        # Layout richiesto: lavoro solo su colonne A..F
        for row in range(1, 500):
            for col in range(1, 7):
                ws.Cells(row, col).Value = None

        # Colonne D/F larghezze guida; A,B,C,E si adattano al contenuto a fine elaborazione.
        ws.Columns("D").ColumnWidth = 4.71
        ws.Columns("F").ColumnWidth = 28.29

        # Font richiesto
        ws.Range("A1:F500").Font.Name = "Arial"
        ws.Range("A1:F500").Font.Size = 14

        # Righe 1 e 2 unificate A..F e vuote
        ws.Range("A1:F1").Merge()
        ws.Range("A2:F2").Merge()

        # Righe 3 e 4 vuote (già pulite)

        # Riga 5: intestazioni tabella (nome foglio originale in colonna E)
        ws.Cells(5, 2).Value = "kW"
        ws.Cells(5, 3).Value = "A"
        ws.Cells(5, 5).Value = "Rif."
        ws.Range("A5:F5").Font.Bold = True

        # Righe 5-6: griglia intestazione
        for row in (5, 6):
            rng = ws.Range(ws.Cells(row, 1), ws.Cells(row, 6))
            rng.Borders.LineStyle = XL_CONTINUOUS
            rng.Borders.Weight = XL_MEDIUM

        # Righe da 7 in basso: Motore vite 1..N dal foglio V (testo COMANDO VITE x)
        vite_last_row = _SMISTAMENTO_VITE_START_ROW - 1
        if vite_rows:
            r = _SMISTAMENTO_VITE_START_ROW
            for item in vite_rows:
                ws.Cells(r, 1).Value = item["title"]
                ws.Cells(r, 2).Value = normalize_output_number(item["kw"])
                ws.Cells(r, 3).Value = normalize_output_number(item["amp"])
                ws.Cells(r, 4).Value = ""
                ws.Cells(r, 5).Value = item["rif"]
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
            ws.Cells(row, 4).Value = ""
            ws.Cells(row, 5).Value = rif  # nome foglio Excel originale
            ws.Cells(row, 6).Value = ""

            if is_tag:
                ws.Range(ws.Cells(row, 1), ws.Cells(row, 5)).Font.Color = 255  # rosso

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

        # Riga totale in basso: formule SUM su B/C così si aggiornano se il progettista modifica i valori sopra.
        total_row = max(24, last_written_row + 3)
        ws.Cells(total_row, 1).Value = "TOTALE"

        # Range SUM: sempre dalla riga dove possono comparire le viti (7), fino all'ultima riga dati tabella.
        first_sum_row = _SMISTAMENTO_VITE_START_ROW
        last_sum_row: int | None = None
        if written_rows:
            last_sum_row = written_rows[-1]["row"]
        elif vite_rows:
            last_sum_row = vite_last_row
        else:
            first_sum_row = None

        if (
            first_sum_row is not None
            and last_sum_row is not None
            and last_sum_row >= first_sum_row
        ):
            ws.Cells(total_row, 2).Formula = f"=SUM(B{first_sum_row}:B{last_sum_row})"
            ws.Cells(total_row, 3).Formula = f"=SUM(C{first_sum_row}:C{last_sum_row})"
        else:
            total_kw = sum(item["kw"] for item in summary_rows) + sum(item["kw"] for item in vite_rows)
            total_amp = sum(item["amp"] for item in summary_rows) + sum(item["amp"] for item in vite_rows)
            ws.Cells(total_row, 2).Value = normalize_output_number(total_kw)
            ws.Cells(total_row, 3).Value = normalize_output_number(total_amp)

        ws.Cells(total_row, 4).Value = ""
        ws.Cells(total_row, 5).Value = ""
        ws.Cells(total_row, 6).Value = ""
        ws.Range(ws.Cells(total_row, 1), ws.Cells(total_row, 6)).Borders.LineStyle = XL_CONTINUOUS
        ws.Range(ws.Cells(total_row, 1), ws.Cells(total_row, 6)).Borders.Weight = XL_MEDIUM

        # kW e A: sempre 2 decimali visibili
        ws.Range(ws.Cells(7, 2), ws.Cells(total_row, 3)).NumberFormat = number_format

        # Bordo destro colonna F dalla riga 1 fino a due righe sotto il totale.
        end_row = total_row + 2
        fcol = ws.Range(ws.Cells(1, 6), ws.Cells(end_row, 6))
        fcol.Borders(XL_EDGE_RIGHT).LineStyle = XL_CONTINUOUS
        fcol.Borders(XL_EDGE_RIGHT).Weight = XL_MEDIUM

        # Riga due righe sotto il totale: bordo basso da A a F.
        bottom_line = ws.Range(ws.Cells(end_row, 1), ws.Cells(end_row, 6))
        bottom_line.Borders(XL_EDGE_BOTTOM).LineStyle = XL_CONTINUOUS
        bottom_line.Borders(XL_EDGE_BOTTOM).Weight = XL_MEDIUM

        # Colonne A, B, C, E adattate al contenuto.
        for _col in ("A", "B", "C", "E"):
            ws.Columns(_col).AutoFit()

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
        decimal_sep = choose_decimal_separator()
        excel_nf = excel_two_decimal_format(decimal_sep)
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
        log(f"Decimali Excel: {'virgola (,)' if decimal_sep == ',' else 'punto (.)'}")

        pythoncom.CoInitialize()
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        summary_rows: list[dict] = []
        vite_summary_rows: list[dict] = []
        try:
            wb_input = excel.Workbooks.Open(str(output_input_path))
            try:
                if line_type_is_cous_cous(selected_line_type.key):
                    tr_sheet = get_sheet(wb_input, "TR")
                    if tr_sheet is None:
                        raise RuntimeError(
                            "Tipologia Cous Cous: e' obbligatorio il foglio TR con "
                            "COMANDO TRITURATORE PRODOTTO FINE e COMANDO TRITURATORE PRODOTTO GROSSO "
                            "(testo in colonna A o B). Foglio TR assente nel file."
                        )
                    found_tr = collect_comando_trituratore_rows(tr_sheet)
                    validate_trituratore_sheet(found_tr)
                    vite_summary_rows = trituratore_dict_to_summary_rows(found_tr)
                    log("[OK] Foglio TR: COMANDO TRITURATORE FINE e GROSSO verificati.")
                else:
                    expected_vite_n = expected_vite_count_for_line_type(selected_line_type.key)
                    if expected_vite_n is not None:
                        v_sheet = get_sheet(wb_input, "V")
                        if v_sheet is None:
                            raise RuntimeError(
                                f"Tipologia «{selected_line_type.label}»: è obbligatorio il foglio V con "
                                f"{expected_vite_n} righe COMANDO VITE (numeri da 1 a {expected_vite_n}). "
                                "Foglio V assente nel file."
                            )
                        found_vite = collect_comando_vite_rows(v_sheet)
                        validate_vite_sheet(expected_vite_n, found_vite)
                        vite_summary_rows = vite_dict_to_summary_rows(found_vite)
                        log(f"[OK] Foglio V: {expected_vite_n} COMANDO VITE verificati.")

                for cfg in config_rows:
                    log(f"[INFO] Elaboro foglio '{cfg.sheet_name}'...")
                    if cfg.sheet_name.strip().lower() == "v":
                        continue

                    sheet = get_sheet(wb_input, cfg.sheet_name)
                    if sheet is None:
                        log(f"[INFO] Foglio '{cfg.sheet_name}' non presente: salto.")
                        continue

                    use_groups = sheet_uses_range_grouping(cfg.sheet_name)
                    unmatched_rows: list[UnmatchedRow] = []

                    if use_groups:
                        totals_by_group, unmatched_rows, mr_row_shift = insert_totals_rows(
                            sheet,
                            split_cfg,
                            sheet_name=cfg.sheet_name,
                            number_format=excel_nf,
                            cous_cous_line=line_type_is_cous_cous(selected_line_type.key),
                        )

                        map_only_groups = map_group_keys_for_sheet(display_cfg, cfg.sheet_name)
                        for entry in group_definitions(split_cfg):
                            key = entry["key"]
                            if map_only_groups is not None and key not in map_only_groups:
                                continue
                            default_for_title = group_display_label_for_sheet(
                                cfg.sheet_name, key, entry["label"]
                            )
                            title = title_for_sheet_group(
                                cfg.sheet_name, key, default_for_title, display_cfg
                            )
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
                        insert_grand_total_row(sheet, number_format=excel_nf)
                        annotate_mr_group_column_p(
                            sheet,
                            split_cfg,
                            sheet_name=cfg.sheet_name,
                            cous_cous_line=line_type_is_cous_cous(selected_line_type.key),
                        )
                        kw_g, amp_g = compute_sheet_grand_total(sheet)
                        title_u = first_map_title_for_sheet(display_cfg, cfg.sheet_name)
                        if not title_u:
                            title_u = cfg.sheet_name
                            log(
                                f"[WARN] Foglio '{cfg.sheet_name}': nessun titolo in MAP; "
                                "smistamento usa il nome foglio."
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
                            adj_row = item.row + mr_row_shift
                            if row_skips_mr_data_row(
                                sheet.Cells(adj_row, "A").Value,
                                sheet.Cells(adj_row, "B").Value,
                            ):
                                continue
                            log(
                                f"       - riga {adj_row}: codice='{item.raw_code}', parsed={item.parsed_code}, "
                                f"kW={item.kw:.2f}, A={item.amp:.2f}"
                            )

                remove_sheets_not_in_config(wb_input, config_rows, selected_line_type.key)

                wb_input.Save()
                log(f"[OK] File output input aggiornato: {output_input_path.name}")
            finally:
                wb_input.Close(SaveChanges=False)

            create_summary_file(
                excel,
                output_summary_path,
                summary_rows,
                vite_rows=vite_summary_rows,
                number_format=excel_nf,
            )
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
