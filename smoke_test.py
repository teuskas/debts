from __future__ import annotations

from ods_service import download_ods_bytes, list_year_sheet_names, load_sheet_grid

DROPBOX_ODS_PATH = "/Me/DEBITI/RM+RF+RC.ods"


def main() -> None:
    ods_bytes = download_ods_bytes(DROPBOX_ODS_PATH)
    sheets = list_year_sheet_names(ods_bytes)
    if not sheets:
        raise SystemExit("Nessun foglio annuale trovato")

    latest_sheet = sheets[-1]
    grid = load_sheet_grid(ods_bytes, latest_sheet)
    print(f"Foglio selezionato: {latest_sheet}")
    print(f"Righe: {len(grid.rows)}")
    print(f"Colonne: {len(grid.rows[0]) if grid.rows else 0}")


if __name__ == "__main__":
    main()

