from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO
from typing import Iterable

from odf.namespaces import OFFICENS, TABLENS, TEXTNS
from odf.opendocument import load
from odf.style import Style, TableCellProperties
from odf.table import Table

from auth import get_dropbox_client


@dataclass(frozen=True)
class SheetCell:
    value: str
    bg_color: str | None


@dataclass(frozen=True)
class SheetGrid:
    name: str
    rows: list[list[SheetCell]]


def download_ods_bytes(dropbox_path: str) -> bytes:
    dbx = get_dropbox_client()
    _, response = dbx.files_download(dropbox_path)
    return response.content


def list_sheet_names(ods_bytes: bytes) -> list[str]:
    document = load(BytesIO(ods_bytes))
    return [table.getAttribute("name") for table in document.spreadsheet.getElementsByType(Table)]


def list_year_sheet_names(ods_bytes: bytes) -> list[str]:
    names = []
    for name in list_sheet_names(ods_bytes):
        if name == "GenCal":
            continue
        if name.isdigit() and len(name) == 4:
            names.append(name)
    return sorted(names)


def load_sheet_grid(ods_bytes: bytes, sheet_name: str, max_columns: int = 48, max_rows: int = 200) -> SheetGrid:
    document = load(BytesIO(ods_bytes))
    style_bg = _extract_cell_backgrounds(document)
    table = _find_table(document, sheet_name)

    rows: list[list[SheetCell]] = []
    used_row_idx = -1
    used_col_idx = -1

    row_index = 0
    for row_node in table.childNodes:
        if row_index >= max_rows:
            break
        if row_node.qname != (TABLENS, "table-row"):
            continue

        row_repeat_raw = int(row_node.getAttribute("numberrowsrepeated") or 1)
        parsed_row = _parse_row(row_node, style_bg, max_columns)

        row_has_values = any(cell.value or cell.bg_color for cell in parsed_row)
        if row_has_values:
            used_col_idx = max(used_col_idx, _last_used_col(parsed_row))
            # Le righe con contenuto raramente si ripetono molto.
            row_repeat = min(row_repeat_raw, 10)
        else:
            # Le righe vuote si ripetono spesso in grandi blocchi (anche 65536).
            # Le rispettiamo fino a max_rows per mantenere gli indici allineati
            # al foglio reale (es. dati a G52 devono trovarsi all'indice 51).
            row_repeat = min(row_repeat_raw, max_rows)

        for _ in range(row_repeat):
            if row_index >= max_rows:
                break
            rows.append(parsed_row)
            if row_has_values:
                used_row_idx = row_index
            row_index += 1

    if used_row_idx >= 0 and used_col_idx >= 0:
        rows = [row[: used_col_idx + 1] for row in rows[: used_row_idx + 1]]
    else:
        rows = []

    return SheetGrid(name=sheet_name, rows=rows)


def _extract_cell_backgrounds(document) -> dict[str, str | None]:
    style_bg: dict[str, str | None] = {}
    style_nodes: Iterable[Style] = (
        document.automaticstyles.getElementsByType(Style) + document.styles.getElementsByType(Style)
    )

    for style in style_nodes:
        if style.getAttribute("family") != "table-cell":
            continue

        style_name = style.getAttribute("name")
        properties = style.getElementsByType(TableCellProperties)
        bg_color = properties[0].getAttribute("backgroundcolor") if properties else None
        style_bg[style_name] = _normalize_background_color(bg_color)

    return style_bg


def _parse_row(row_node, style_bg: dict[str, str | None], max_columns: int) -> list[SheetCell]:
    parsed_row: list[SheetCell] = []

    for cell_node in row_node.childNodes:
        if cell_node.qname not in ((TABLENS, "table-cell"), (TABLENS, "covered-table-cell")):
            continue

        col_repeat = int(cell_node.getAttribute("numbercolumnsrepeated") or 1)
        col_repeat = min(col_repeat, max_columns)

        style_name = cell_node.getAttribute("stylename")
        bg_color = style_bg.get(style_name)
        value = _extract_cell_text(cell_node).strip()

        for _ in range(col_repeat):
            if len(parsed_row) >= max_columns:
                break
            parsed_row.append(SheetCell(value=value, bg_color=bg_color))

        if len(parsed_row) >= max_columns:
            break

    if len(parsed_row) < max_columns:
        parsed_row.extend(SheetCell(value="", bg_color=None) for _ in range(max_columns - len(parsed_row)))

    return parsed_row


def _extract_cell_text(cell_node) -> str:
    """Estrae il testo della cella ignorando le annotazioni (commenti ODS)."""
    chunks: list[str] = []
    for child in cell_node.childNodes:
        # Salta office:annotation (commenti)
        if child.qname == (OFFICENS, "annotation"):
            continue
        # Nodo testo puro
        if hasattr(child, "data"):
            chunks.append(child.data)
            continue
        # Paragrafo testo (text:p) → entra ricorsivamente escludendo annotazioni
        if child.qname == (TEXTNS, "p"):
            chunks.append(_extract_text_paragraph(child))
    return "".join(chunks)


def _extract_text_paragraph(node) -> str:
    parts: list[str] = []
    for child in node.childNodes:
        if not hasattr(child, "qname"):
            if hasattr(child, "data"):
                parts.append(child.data)
            continue
        if child.qname == (OFFICENS, "annotation"):
            continue
        if hasattr(child, "data"):
            parts.append(child.data)
        else:
            parts.append(_extract_text_paragraph(child))
    return "".join(parts)


def _last_used_col(row: list[SheetCell]) -> int:
    for idx in range(len(row) - 1, -1, -1):
        if row[idx].value or row[idx].bg_color:
            return idx
    return -1


def _find_table(document, sheet_name: str) -> Table:
    for table in document.spreadsheet.getElementsByType(Table):
        if table.getAttribute("name") == sheet_name:
            return table
    raise ValueError(f"Foglio '{sheet_name}' non trovato nel file ODS")


def _normalize_background_color(color: str | None) -> str | None:
    if not color:
        return None

    clean = color.strip().lower()
    if clean in {"transparent", "white", "#ffffff"}:
        return None

    if clean.startswith("#"):
        return clean.upper()

    return color

