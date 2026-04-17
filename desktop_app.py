from __future__ import annotations

import re
import threading
import tkinter as tk
from datetime import datetime
from io import BytesIO
from tkinter import messagebox, ttk

from ods_service import SheetGrid, download_ods_bytes, list_year_sheet_names, load_sheet_grid

DEFAULT_DROPBOX_PATH = "/Me/DEBITI/RM+RF+RC.ods"
ANNUAL_YEARS_SEPARATOR = "--------------------"

# Palette allineata al look Financiary desktop.
BG = "#1e1e2e"
BG_FRAME = "#2a2a3e"
BG_TABLE = "#2a2a3e"
FG = "#cdd6f4"
FG_HEADER = "#89b4fa"
FG_ACCENT = "#f38ba8"
FG_POSITIVE = "#22c55e"
FG_NEGATIVE = "#ef4444"

TOTAL_RATES_BY_TARGET = {
    "Mutuo": 320,
    "Finanziamento Casa": 120,
    "Macchina": 84,
}

TARGET_COLUMN_INDEX = {
    "Mutuo": 2,  # Colonna C
    "Finanziamento Casa": 5,  # Colonna F
    "Macchina": 8,  # Colonna I
}

TOTAL_AMOUNT_BY_TARGET = {
    "Mutuo": 193592.39,
    "Finanziamento Casa": 16555.0,
    "Macchina": 13498.75,
}

TOTAL_CAPITAL_BY_TARGET = {
    "Mutuo": 120000.0,
    "Finanziamento Casa": 10000.0,
    "Macchina": 10058.0,
}

TOTAL_COLUMN_BY_TARGET = {
    "Mutuo": "RM",
    "Finanziamento Casa": "RF",
    "Macchina": "RC",
}

DEBT_COMBINATIONS = {
    "Mutuo + Macchina": ("Mutuo", "Macchina"),
    "Mutuo + Finanziamento Casa": ("Mutuo", "Finanziamento Casa"),
    "Finanziamento Casa + Macchina": ("Finanziamento Casa", "Macchina"),
}

RIMANENZE_COMBINATIONS = {
    "Mutuo + Fhome": ("Mutuo", "Finanziamento Casa"),
    "Mutuo + Fcar": ("Mutuo", "Macchina"),
    "Fhome + Fcar": ("Finanziamento Casa", "Macchina"),
    "All Debts": ("Mutuo", "Finanziamento Casa", "Macchina"),
}

DEBT_SHORT_NAME = {
    "Mutuo": "Mutuo",
    "Finanziamento Casa": "Fhome",
    "Macchina": "Fcar",
}

# Label header GenCal per ciascun debito (sezione)
GENCAL_SECTION_LABELS = {
    "Mutuo": {"RM", "MUTUO"},
    "Finanziamento Casa": {"RF", "FINANZIAMENTO CASA", "FCASA", "FHOME"},
    "Macchina": {"RC", "MACCHINA", "FCAR"},
}

# Alias trovati nella colonna DR del GenCal (formato attuale con una riga per debito/mese).
GENCAL_DR_TO_DEBT = {
    "CASA": "Mutuo",
    "MUTUO": "Mutuo",
    "RM": "Mutuo",
    "FHOME": "Finanziamento Casa",
    "FINANZIAMENTO CASA": "Finanziamento Casa",
    "RF": "Finanziamento Casa",
    "CAR": "Macchina",
    "MACCHINA": "Macchina",
    "FCAR": "Macchina",
    "RC": "Macchina",
}

MONTH_NAMES_IT = ["Gen", "Feb", "Mar", "Apr", "Mag", "Giu", "Lug", "Ago", "Set", "Ott", "Nov", "Dic"]

CAPITAL_FILE_BY_TARGET = {
    "Mutuo": "/Me/DEBITI/RMQ.ods",
    "Finanziamento Casa": "/Me/DEBITI/FCQ.ods",
    "Macchina": "/Me/DEBITI/FCAR.ods",
}

PLATFORM_NAMES = ("Bondora", "Mintos", "ReLender")
NEW_TOTAL_INV_PATH_CANDIDATES = (
    "/Me/NEW TOTAL INV.xlsx",
    "/Me/DEBITI/NEW TOTAL INV.ods",
    "/Me/DEBITI/NEW TOTAL INV",
)
PLATFORM_CELL_COORDS = {
    "ReLender": (51, 6),  # G52
    "Bondora": (52, 6),  # G53
    "Mintos": (53, 6),  # G54
}


class DebtsDesktopApp(tk.Tk):
    def __init__(self, dropbox_path: str = DEFAULT_DROPBOX_PATH) -> None:
        super().__init__(className="Debts")
        self.title("Own Finance - Debts")
        self.geometry("1180x720")
        self.configure(bg=BG)

        self.dropbox_path = dropbox_path
        self.ods_bytes: bytes | None = None
        self.sheet_names: list[str] = []
        self.gencal_sections: dict[str, SheetGrid] = {}
        self.current_annual_grid: SheetGrid | None = None
        self.capital_paid_by_target: dict[str, float] = {}
        self.platform_amounts_by_name: dict[str, float] = {name: 0.0 for name in PLATFORM_NAMES}
        self._capital_file_cache: dict[str, bytes] = {}
        self._platform_file_cache: dict[str, bytes] = {}
        self._hover_tooltip: tk.Toplevel | None = None

        self.annual_year_var = tk.StringVar()
        self.future_year_var = tk.StringVar()
        self.completion_mode_var = tk.StringVar(value="Per Capitale")
        self.completion_target_var = tk.StringVar(value="Mutuo")
        self.diff_debt_var = tk.StringVar(value="Tutti")
        self.diff_platform_var = tk.StringVar(value="Tutti")
        self.diff_scope_var = tk.StringVar(value="Totale")
        self.rimanenze_year_var = tk.StringVar()
        self.rimanenze_combo_var = tk.StringVar(value="Fhome + Fcar")
        self.rimanenze_scope_var = tk.StringVar(value="Capitale")
        self.status_var = tk.StringVar(value="Inizializzazione...")

        self._build_style()
        self._build_layout()
        self._load_all_async()

    def _build_style(self) -> None:
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TNotebook", background=BG, borderwidth=0)
        style.configure("TNotebook.Tab", background=BG_FRAME, foreground=FG, font=("Segoe UI", 10, "bold"), padding=[14, 6])
        style.map("TNotebook.Tab", background=[("selected", BG_TABLE)], foreground=[("selected", FG_HEADER)])
        style.configure("TFrame", background=BG)
        style.configure("Header.TLabel", background=BG, foreground=FG_HEADER, font=("Segoe UI", 14, "bold"))
        style.configure("SubHeader.TLabel", background=BG, foreground=FG, font=("Segoe UI", 10))
        style.configure("Status.TLabel", background=BG, foreground=FG, font=("Segoe UI", 9))
        style.configure("TLabel", background=BG, foreground=FG, font=("Segoe UI", 10))
        style.configure("TButton", background=BG_FRAME, foreground=FG, font=("Segoe UI", 10, "bold"), padding=6)
        style.map("TButton", background=[("active", FG_ACCENT)], foreground=[("active", "#1f1f2e")])
        style.configure("TCombobox", fieldbackground="#f3f4f6", background="#f3f4f6", foreground="#000000")
        style.map(
            "TCombobox",
            foreground=[("readonly", "#000000"), ("!disabled", "#000000")],
            fieldbackground=[("readonly", "#f3f4f6"), ("!disabled", "#f3f4f6")],
            selectforeground=[("readonly", "#000000")],
        )
        self.option_add("*TCombobox*Listbox.Foreground", "#000000")
        self.option_add("*TCombobox*Listbox.Background", "#f3f4f6")
        self.option_add("*TCombobox*Listbox.selectForeground", "#000000")
        self.option_add("*TCombobox*Listbox.selectBackground", "#d1d5db")

    def _build_layout(self) -> None:
        top = ttk.Frame(self)
        top.pack(fill="x", padx=14, pady=(12, 8))

        ttk.Label(top, text="Own Finance - Debts", style="Header.TLabel").pack(anchor="w")
        ttk.Label(top, text="Rate annuali RM + RF + RC (GenCal escluso)", style="SubHeader.TLabel").pack(anchor="w")

        ttk.Label(self, text=f"Percorso: {self.dropbox_path}", style="SubHeader.TLabel").pack(fill="x", padx=14, pady=(0, 8), anchor="w")

        notebook = ttk.Notebook(self)
        notebook.pack(fill="both", expand=True, padx=10, pady=(0, 8))

        annual_tab = tk.Frame(notebook, bg=BG)
        future_tab = tk.Frame(notebook, bg=BG)
        completion_tab = tk.Frame(notebook, bg=BG)
        differences_tab = tk.Frame(notebook, bg=BG)
        rimanenze_tab = tk.Frame(notebook, bg=BG)
        notebook.add(annual_tab, text="Rate annuali")
        notebook.add(future_tab, text="Rate future")
        notebook.add(completion_tab, text="Completamento")
        notebook.add(differences_tab, text="Differenze")
        notebook.add(rimanenze_tab, text="Rimanenze")

        annual_controls = ttk.Frame(annual_tab)
        annual_controls.pack(fill="x", padx=8, pady=(8, 6))
        ttk.Label(annual_controls, text="Anno:").pack(side="left")
        self.annual_year_cb = ttk.Combobox(annual_controls, width=12, textvariable=self.annual_year_var, state="disabled")
        self.annual_year_cb.pack(side="left", padx=(6, 12))
        self.annual_year_cb.bind("<<ComboboxSelected>>", self._on_annual_year_change)

        future_controls = ttk.Frame(future_tab)
        future_controls.pack(fill="x", padx=8, pady=(8, 6))
        ttk.Label(future_controls, text="Anno (GenCal):").pack(side="left")
        self.future_year_cb = ttk.Combobox(future_controls, width=12, textvariable=self.future_year_var, state="disabled")
        self.future_year_cb.pack(side="left", padx=(6, 12))
        self.future_year_cb.bind("<<ComboboxSelected>>", self._on_future_year_change)

        completion_controls = ttk.Frame(completion_tab)
        completion_controls.pack(fill="x", padx=8, pady=(8, 6))
        ttk.Label(completion_controls, text="Vista:").pack(side="left")
        self.completion_mode_cb = ttk.Combobox(
            completion_controls,
            width=16,
            textvariable=self.completion_mode_var,
            state="disabled",
            values=["Per Rate", "Per Totale", "Per Capitale"],
        )
        self.completion_mode_cb.pack(side="left", padx=(6, 12))
        self.completion_mode_cb.bind("<<ComboboxSelected>>", self._on_completion_mode_change)

        ttk.Label(completion_controls, text="Voce:").pack(side="left")
        self.completion_target_cb = ttk.Combobox(
            completion_controls,
            width=22,
            textvariable=self.completion_target_var,
            state="disabled",
            values=["Mutuo", "Finanziamento Casa", "Macchina"],
        )
        self.completion_target_cb.pack(side="left", padx=(6, 12))
        self.completion_target_cb.bind("<<ComboboxSelected>>", self._on_completion_target_change)

        completion_hint = ttk.Label(
            completion_tab,
            text="Completamento: Per Rate, Per Totale, Per Capitale.",
            style="SubHeader.TLabel",
        )
        completion_hint.pack(anchor="w", padx=8, pady=(2, 8))

        self.completion_chart_canvas = tk.Canvas(completion_tab, bg=BG_TABLE, highlightthickness=0, height=300)
        self.completion_chart_canvas.pack(fill="both", expand=True, padx=8, pady=(0, 8))

        self.completion_chart_legend_var = tk.StringVar(value="In attesa dei dati...")
        self.completion_chart_legend = ttk.Label(completion_tab, textvariable=self.completion_chart_legend_var, style="SubHeader.TLabel")
        self.completion_chart_legend.pack(anchor="w", padx=8, pady=(0, 8))

        differences_controls = ttk.Frame(differences_tab)
        differences_controls.pack(fill="x", padx=8, pady=(8, 6))

        ttk.Label(differences_controls, text="Debito:").pack(side="left")
        self.diff_debt_cb = ttk.Combobox(
            differences_controls,
            width=22,
            textvariable=self.diff_debt_var,
            state="disabled",
            values=[
                "Mutuo",
                "Finanziamento Casa",
                "Macchina",
                "Mutuo + Macchina",
                "Mutuo + Finanziamento Casa",
                "Finanziamento Casa + Macchina",
                "Tutti",
            ],
        )
        self.diff_debt_cb.pack(side="left", padx=(6, 12))
        self.diff_debt_cb.bind("<<ComboboxSelected>>", self._on_differences_filter_change)

        ttk.Label(differences_controls, text="Piattaforma:").pack(side="left")
        self.diff_platform_cb = ttk.Combobox(
            differences_controls,
            width=16,
            textvariable=self.diff_platform_var,
            state="disabled",
            values=["Bondora", "Mintos", "ReLender", "Tutti"],
        )
        self.diff_platform_cb.pack(side="left", padx=(6, 12))
        self.diff_platform_cb.bind("<<ComboboxSelected>>", self._on_differences_filter_change)

        ttk.Label(differences_controls, text="Ambito:").pack(side="left")
        self.diff_scope_cb = ttk.Combobox(
            differences_controls,
            width=14,
            textvariable=self.diff_scope_var,
            state="disabled",
            values=["Totale", "Solo Capitale"],
        )
        self.diff_scope_cb.pack(side="left", padx=(6, 12))
        self.diff_scope_cb.bind("<<ComboboxSelected>>", self._on_differences_filter_change)

        ttk.Label(
            differences_tab,
            text="Filtri differenze pronti: in attesa della logica di confronto.",
            style="SubHeader.TLabel",
        ).pack(anchor="w", padx=8, pady=(2, 8))

        self.differences_table_body = self._build_table_area(differences_tab)

        # --- Tab Rimanenze ---
        rimanenze_controls = ttk.Frame(rimanenze_tab)
        rimanenze_controls.pack(fill="x", padx=8, pady=(8, 6))

        ttk.Label(rimanenze_controls, text="Anno:").pack(side="left")
        self.rimanenze_year_cb = ttk.Combobox(rimanenze_controls, width=10, textvariable=self.rimanenze_year_var, state="disabled")
        self.rimanenze_year_cb.pack(side="left", padx=(6, 12))
        self.rimanenze_year_cb.bind("<<ComboboxSelected>>", self._on_rimanenze_filter_change)

        ttk.Label(rimanenze_controls, text="Combinazione:").pack(side="left")
        self.rimanenze_combo_cb = ttk.Combobox(
            rimanenze_controls,
            width=18,
            textvariable=self.rimanenze_combo_var,
            state="disabled",
            values=list(RIMANENZE_COMBINATIONS.keys()),
        )
        self.rimanenze_combo_cb.pack(side="left", padx=(6, 12))
        self.rimanenze_combo_cb.bind("<<ComboboxSelected>>", self._on_rimanenze_filter_change)

        ttk.Label(rimanenze_controls, text="Ambito:").pack(side="left")
        self.rimanenze_scope_cb = ttk.Combobox(
            rimanenze_controls,
            width=12,
            textvariable=self.rimanenze_scope_var,
            state="disabled",
            values=["Capitale", "Totale"],
        )
        self.rimanenze_scope_cb.pack(side="left", padx=(6, 12))
        self.rimanenze_scope_cb.bind("<<ComboboxSelected>>", self._on_rimanenze_filter_change)

        self.rimanenze_table_body = self._build_table_area(rimanenze_tab)

        self.annual_table_body = self._build_table_area(annual_tab)
        self.future_table_body = self._build_table_area(future_tab)

        status = ttk.Label(self, textvariable=self.status_var, style="Status.TLabel")
        status.pack(fill="x", padx=14, pady=(0, 12), anchor="w")
        self._refresh_completion_view()
        self._refresh_differences_view()
        self._refresh_rimanenze_view()

    def _build_table_area(self, parent: tk.Frame) -> tk.Frame:
        table_shell = tk.Frame(parent, bg=BG_FRAME, bd=1, relief="solid", highlightthickness=0)
        table_shell.pack(fill="both", expand=True, padx=8, pady=(0, 8))

        canvas = tk.Canvas(table_shell, bg=BG_TABLE, highlightthickness=0)
        vsb = ttk.Scrollbar(table_shell, orient="vertical", command=canvas.yview)
        hsb = ttk.Scrollbar(table_shell, orient="horizontal", command=canvas.xview)
        canvas.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")
        canvas.pack(side="left", fill="both", expand=True)

        body = tk.Frame(canvas, bg=BG_TABLE)
        table_window = canvas.create_window((0, 0), window=body, anchor="nw")

        body.bind("<Configure>", lambda _event: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda event: canvas.itemconfigure(table_window, width=max(event.width, body.winfo_reqwidth())))
        return body

    def _on_annual_year_change(self, _event=None) -> None:
        year = self.annual_year_var.get().strip()
        if year == ANNUAL_YEARS_SEPARATOR:
            # Il separatore e' solo visuale: ripristina l'anno caricato.
            if self.current_annual_grid is not None:
                self.annual_year_var.set(self.current_annual_grid.name)
            return
        if year:
            self._load_annual_sheet_async(year)

    def _on_future_year_change(self, _event=None) -> None:
        year = self.future_year_var.get().strip()
        if not year:
            return
        grid = self.gencal_sections.get(year)
        if grid is None:
            return
        self._render_grid(self.future_table_body, grid)
        self.status_var.set(f"GenCal {year} caricato - {len(grid.rows)} righe")

    def _on_completion_mode_change(self, _event=None) -> None:
        mode = self.completion_mode_var.get().strip()
        if mode:
            self.status_var.set(f"Completamento: {mode} | {self.completion_target_var.get().strip()}")
        self._refresh_completion_view()

    def _on_completion_target_change(self, _event=None) -> None:
        target = self.completion_target_var.get().strip()
        if target:
            self.status_var.set(f"Completamento: {self.completion_mode_var.get().strip()} | {target}")
        self._refresh_completion_view()

    def _on_differences_filter_change(self, _event=None) -> None:
        self.status_var.set(
            "Differenze: "
            f"{self.diff_debt_var.get().strip()} | "
            f"{self.diff_platform_var.get().strip()} | "
            f"{self.diff_scope_var.get().strip()}"
        )
        self._refresh_differences_view()

    def _on_rimanenze_filter_change(self, _event=None) -> None:
        self._refresh_rimanenze_view()

    def _show_hover_tooltip(self, event, text: str) -> None:
        self._hide_hover_tooltip()
        tooltip = tk.Toplevel(self)
        tooltip.wm_overrideredirect(True)
        tooltip.configure(bg="#111827")
        tk.Label(
            tooltip,
            text=text,
            bg="#111827",
            fg="#ffffff",
            font=("Segoe UI", 9),
            padx=8,
            pady=4,
        ).pack()

        x = event.x_root + 12
        y = event.y_root + 12
        tooltip.wm_geometry(f"+{x}+{y}")
        self._hover_tooltip = tooltip

    def _hide_hover_tooltip(self, _event=None) -> None:
        if self._hover_tooltip is not None:
            try:
                self._hover_tooltip.destroy()
            except Exception:
                pass
            self._hover_tooltip = None

    def _set_controls_busy(self, busy: bool) -> None:
        self.annual_year_cb.configure(state="disabled" if busy or not self.sheet_names else "readonly")
        self.future_year_cb.configure(state="disabled" if busy or not self.gencal_sections else "readonly")
        self.completion_mode_cb.configure(state="disabled" if busy else "readonly")
        self.completion_target_cb.configure(state="disabled" if busy else "readonly")
        self.diff_debt_cb.configure(state="disabled" if busy else "readonly")
        self.diff_platform_cb.configure(state="disabled" if busy else "readonly")
        self.diff_scope_cb.configure(state="disabled" if busy else "readonly")
        self.rimanenze_year_cb.configure(state="disabled" if busy or not self.gencal_sections else "readonly")
        self.rimanenze_combo_cb.configure(state="disabled" if busy else "readonly")
        self.rimanenze_scope_cb.configure(state="disabled" if busy else "readonly")

    def _load_all_async(self) -> None:
        self.status_var.set("Carico dati da Dropbox...")
        self._set_controls_busy(True)

        def job() -> None:
            try:
                if self.ods_bytes is None:
                    self.ods_bytes = download_ods_bytes(self.dropbox_path)

                years = list_year_sheet_names(self.ods_bytes)
                if not years:
                    raise ValueError("Nessun foglio anno trovato nel file (GenCal escluso)")

                gencal_grid = load_sheet_grid(self.ods_bytes, "GenCal", max_columns=64, max_rows=2000)
                gencal_sections = _split_gencal_by_year(gencal_grid)
                if not gencal_sections:
                    raise ValueError("GenCal presente ma non suddivisibile per anno")

                current_year = str(datetime.now().year)
                if current_year in years:
                    annual_preferred = current_year
                elif self.annual_year_var.get() in years:
                    annual_preferred = self.annual_year_var.get()
                else:
                    annual_preferred = years[-1]

                future_years = sorted(gencal_sections.keys())
                if current_year in future_years:
                    future_preferred = current_year
                elif self.future_year_var.get() in future_years:
                    future_preferred = self.future_year_var.get()
                else:
                    future_preferred = future_years[-1]

                annual_grid = load_sheet_grid(self.ods_bytes, annual_preferred)
                future_grid = gencal_sections[future_preferred]
                capital_paid_by_target = _extract_paid_capital_by_target_from_dropbox(
                    self._capital_file_cache,
                    str(datetime.now().year),
                )
                platform_amounts_by_name = _extract_platform_amounts_from_dropbox(
                    self._platform_file_cache,
                    str(datetime.now().year),
                )
            except Exception as exc:  # pragma: no cover - dipende da I/O Dropbox
                self.after(0, self._on_load_error, exc)
                return

            self.after(
                0,
                self._on_load_success,
                years,
                annual_preferred,
                annual_grid,
                gencal_sections,
                future_preferred,
                future_grid,
                capital_paid_by_target,
                platform_amounts_by_name,
            )

        threading.Thread(target=job, daemon=True).start()

    def _load_annual_sheet_async(self, year: str) -> None:
        if self.ods_bytes is None:
            self._load_all_async()
            return

        self.status_var.set(f"Carico foglio {year}...")
        self._set_controls_busy(True)

        def job() -> None:
            try:
                grid = load_sheet_grid(self.ods_bytes, year)
                capital_paid_by_target = _extract_paid_capital_by_target_from_dropbox(
                    self._capital_file_cache,
                    str(datetime.now().year),
                )
                platform_amounts_by_name = _extract_platform_amounts_from_dropbox(
                    self._platform_file_cache,
                    str(datetime.now().year),
                )
            except Exception as exc:  # pragma: no cover - dipende da contenuto ODS
                self.after(0, self._on_load_error, exc)
                return
            self.after(0, self._on_annual_sheet_success, year, grid, capital_paid_by_target, platform_amounts_by_name)

        threading.Thread(target=job, daemon=True).start()

    def _on_load_error(self, exc: Exception) -> None:
        self._set_controls_busy(False)
        self.status_var.set("Errore durante il caricamento")
        messagebox.showerror("Debts", f"Errore: {exc}")

    def _on_load_success(
        self,
        years: list[str],
        selected_year: str,
        annual_grid: SheetGrid,
        gencal_sections: dict[str, SheetGrid],
        selected_future_year: str,
        future_grid: SheetGrid,
        capital_paid_by_target: dict[str, float],
        platform_amounts_by_name: dict[str, float],
    ) -> None:
        self.sheet_names = years
        self.gencal_sections = gencal_sections
        annual_year_values = _build_annual_year_dropdown_values(years, datetime.now().year)
        self.annual_year_cb.configure(values=annual_year_values)
        self.future_year_cb.configure(values=sorted(gencal_sections.keys()))
        self.annual_year_var.set(selected_year)
        self.future_year_var.set(selected_future_year)
        self.current_annual_grid = annual_grid
        self.capital_paid_by_target = capital_paid_by_target
        self.platform_amounts_by_name = platform_amounts_by_name
        self._render_grid(self.annual_table_body, annual_grid)
        self._render_grid(self.future_table_body, future_grid)
        self._refresh_completion_view()
        self._refresh_differences_view()
        # Rimanenze year dropdown: usa gli anni del GenCal (anni futuri da pagare)
        rim_years = sorted(gencal_sections.keys())
        self.rimanenze_year_cb.configure(values=rim_years)
        current_year_str = str(datetime.now().year)
        if not self.rimanenze_year_var.get() or self.rimanenze_year_var.get() not in rim_years:
            self.rimanenze_year_var.set(current_year_str if current_year_str in rim_years else (rim_years[0] if rim_years else ""))
        self._refresh_rimanenze_view()
        self._set_controls_busy(False)
        self.status_var.set(
            f"Annuale {selected_year} ({len(annual_grid.rows)} righe) | "
            f"GenCal {selected_future_year} ({len(future_grid.rows)} righe)"
        )

    def _on_annual_sheet_success(
        self,
        year: str,
        grid: SheetGrid,
        capital_paid_by_target: dict[str, float],
        platform_amounts_by_name: dict[str, float],
    ) -> None:
        self.annual_year_var.set(year)
        self.current_annual_grid = grid
        self.capital_paid_by_target = capital_paid_by_target
        self.platform_amounts_by_name = platform_amounts_by_name
        self._render_grid(self.annual_table_body, grid)
        self._refresh_completion_view()
        self._refresh_differences_view()
        self._refresh_rimanenze_view()
        self._set_controls_busy(False)
        self.status_var.set(f"Foglio {year} caricato - {len(grid.rows)} righe")

    def _refresh_differences_view(self) -> None:
        for child in self.differences_table_body.winfo_children():
            child.destroy()

        if self.current_annual_grid is None:
            tk.Label(
                self.differences_table_body,
                text="Dati annuali non ancora caricati",
                bg=BG_TABLE,
                fg=FG,
                padx=10,
                pady=10,
                anchor="w",
            ).grid(row=0, column=0, sticky="w")
            return

        debt_scope = self.diff_scope_var.get().strip()
        selected_debt = self.diff_debt_var.get().strip()
        selected_platform = self.diff_platform_var.get().strip()

        # Debiti mancanti: in rosso (sono soldi che devo).
        outstanding_by_debt = _compute_outstanding_by_debt(
            self.current_annual_grid,
            self.capital_paid_by_target,
            debt_scope,
        )

        # Disponibilita piattaforme: in verde (sono soldi che ho), da NEW TOTAL INV.
        available_by_platform = self.platform_amounts_by_name

        if selected_debt == "Tutti":
            debt_items = [(name, outstanding_by_debt.get(name, 0.0)) for name in TOTAL_AMOUNT_BY_TARGET]
        elif selected_debt in DEBT_COMBINATIONS:
            members = DEBT_COMBINATIONS[selected_debt]
            debt_items = [(name, outstanding_by_debt.get(name, 0.0)) for name in members]
        else:
            debt_items = [(selected_debt, outstanding_by_debt.get(selected_debt, 0.0))]

        if selected_platform == "Tutti":
            platform_items = [(name, available_by_platform.get(name, 0.0)) for name in PLATFORM_NAMES]
        else:
            platform_items = [(selected_platform, available_by_platform.get(selected_platform, 0.0))]

        total_debt = sum(value for _, value in debt_items)
        total_platform = sum(value for _, value in platform_items)
        difference = total_debt - total_platform

        rows: list[tuple[str, str, float, str]] = []
        for name, value in debt_items:
            rows.append(("Debito", name, value, FG_NEGATIVE))
        if selected_debt == "Tutti" or selected_debt in DEBT_COMBINATIONS:
            rows.append(("Debito", "Somma Debiti", total_debt, FG_NEGATIVE))

        for name, value in platform_items:
            rows.append(("Piattaforma", name, value, FG_POSITIVE))
        if selected_platform == "Tutti":
            rows.append(("Piattaforma", "Somma Piattaforme", total_platform, FG_POSITIVE))

        difference_color = FG_NEGATIVE if difference > 0 else FG_POSITIVE
        rows.append(("Differenza", "Debiti - Piattaforme", difference, difference_color))

        headers = ("Categoria", "Voce", "Importo")
        for cidx, text in enumerate(headers):
            tk.Label(
                self.differences_table_body,
                text=text,
                bg=BG_FRAME,
                fg=FG_HEADER,
                font=("Segoe UI", 10, "bold"),
                bd=1,
                relief="solid",
                padx=8,
                pady=5,
                anchor="center",
            ).grid(row=0, column=cidx, sticky="nsew")

        for ridx, (category, label, amount, color) in enumerate(rows, start=1):
            tk.Label(
                self.differences_table_body,
                text=category,
                bg=BG_TABLE,
                fg=FG,
                font=("Segoe UI", 10),
                bd=1,
                relief="solid",
                padx=8,
                pady=5,
                anchor="w",
            ).grid(row=ridx, column=0, sticky="nsew")

            tk.Label(
                self.differences_table_body,
                text=label,
                bg=BG_TABLE,
                fg=FG,
                font=("Segoe UI", 10, "bold" if category == "Differenza" else "normal"),
                bd=1,
                relief="solid",
                padx=8,
                pady=5,
                anchor="w",
            ).grid(row=ridx, column=1, sticky="nsew")

            tk.Label(
                self.differences_table_body,
                text=_format_signed_amount(amount) if category == "Differenza" else _format_amount(amount),
                bg=BG_TABLE,
                fg=color,
                font=("Segoe UI", 10, "bold"),
                bd=1,
                relief="solid",
                padx=8,
                pady=5,
                anchor="e",
            ).grid(row=ridx, column=2, sticky="nsew")

        self.differences_table_body.grid_columnconfigure(0, weight=0, minsize=120)
        self.differences_table_body.grid_columnconfigure(1, weight=1, minsize=260)
        self.differences_table_body.grid_columnconfigure(2, weight=0, minsize=170)

    def _refresh_rimanenze_view(self) -> None:
        self._hide_hover_tooltip()
        for child in self.rimanenze_table_body.winfo_children():
            child.destroy()

        selected_year = self.rimanenze_year_var.get().strip()
        combo = self.rimanenze_combo_var.get().strip()
        scope = self.rimanenze_scope_var.get().strip()

        if not selected_year or not self.gencal_sections:
            tk.Label(self.rimanenze_table_body, text="Dati non ancora disponibili", bg=BG_TABLE, fg=FG, padx=10, pady=10).grid(row=0, column=0)
            return

        debts = RIMANENZE_COMBINATIONS.get(combo, ("Finanziamento Casa", "Macchina"))
        grid = self.gencal_sections.get(selected_year)
        if grid is None:
            tk.Label(self.rimanenze_table_body, text=f"Nessun dato GenCal per l'anno {selected_year}", bg=BG_TABLE, fg=FG, padx=10, pady=10).grid(row=0, column=0)
            return

        # Calcola il saldo iniziale rimanente per ciascun debito (da logica esistente)
        if scope == "Capitale":
            initial_remaining = {
                debt: max(0.0, TOTAL_CAPITAL_BY_TARGET[debt] - self.capital_paid_by_target.get(debt, 0.0))
                for debt in debts
            }
        else:
            paid_totals = {}
            if self.current_annual_grid is not None:
                paid_totals = _extract_paid_total_amounts_by_target(self.current_annual_grid)
            initial_remaining = {
                debt: max(0.0, TOTAL_AMOUNT_BY_TARGET[debt] - paid_totals.get(debt, 0.0))
                for debt in debts
            }

        # Per il Capitale usa i file dedicati RMQ/FCQ/FCAR (richiesta utente).
        if scope == "Capitale":
            monthly_rows = _extract_monthly_capital_rimanenze_from_dedicated_files(
                debts,
                selected_year,
                initial_remaining,
                self._capital_file_cache,
            )
        else:
            # Totale: continua a usare il GenCal.
            show_full_year = selected_year == str(datetime.now().year)
            monthly_rows = _extract_monthly_rimanenze(
                grid,
                debts,
                scope,
                initial_remaining,
                selected_year,
                fill_full_year=show_full_year,
            )

        if not monthly_rows:
            tk.Label(self.rimanenze_table_body, text="Impossibile estrarre dati mensili dal GenCal", bg=BG_TABLE, fg=FG, padx=10, pady=10).grid(row=0, column=0)
            return

        # Intestazioni: Mese + un header per ogni debito + Totale
        debt_short = [DEBT_SHORT_NAME.get(d, d) for d in debts]
        headers = ["Mese"] + debt_short + ["Totale"]
        col_count = len(headers)

        for cidx, h in enumerate(headers):
            tk.Label(
                self.rimanenze_table_body, text=h, bg=BG_FRAME, fg=FG_HEADER,
                font=("Segoe UI", 10, "bold"), bd=1, relief="solid", padx=8, pady=5, anchor="center",
            ).grid(row=0, column=cidx, sticky="nsew")

        previous_total: float | None = None

        for ridx, (month_label, values_by_debt, paid_flags_by_debt, paid_amounts_by_debt) in enumerate(monthly_rows, start=1):
            row_vals = [values_by_debt.get(d, 0.0) for d in debts]
            total = sum(row_vals)
            mixed_paid_status = any(bool(paid_flags_by_debt.get(debt, False)) for debt in debts) and not all(
                bool(paid_flags_by_debt.get(debt, False)) for debt in debts
            )

            tk.Label(
                self.rimanenze_table_body, text=month_label, bg=BG_TABLE, fg=FG,
                font=("Segoe UI", 10), bd=1, relief="solid", padx=8, pady=5, anchor="w",
            ).grid(row=ridx, column=0, sticky="nsew")

            for cidx, debt in enumerate(debts, start=1):
                val = values_by_debt.get(debt, 0.0)
                paid_in_month = bool(paid_flags_by_debt.get(debt, False))
                cell_label = tk.Label(
                    self.rimanenze_table_body,
                    text=_format_amount(val),
                    bg=FG_POSITIVE if paid_in_month else BG_TABLE,
                    fg="#ffffff" if paid_in_month else FG_ACCENT,
                    font=("Segoe UI", 10, "bold" if paid_in_month else "normal"),
                    bd=1,
                    relief="solid",
                    padx=8,
                    pady=5,
                    anchor="e",
                )
                cell_label.grid(row=ridx, column=cidx, sticky="nsew")

                paid_amount_for_cell = max(0.0, paid_amounts_by_debt.get(debt, 0.0))
                if paid_in_month and mixed_paid_status and previous_total is not None and paid_amount_for_cell > 0:
                    tooltip_value = max(0.0, previous_total - paid_amount_for_cell)
                    paid_label = "capitale" if scope == "Capitale" else "totale"
                    tooltip_text = (
                        f"Totale mese precedente - {paid_label} pagato: "
                        f"{_format_amount(previous_total)} - {_format_amount(paid_amount_for_cell)} = {_format_amount(tooltip_value)}"
                    )
                    cell_label.bind("<Enter>", lambda event, t=tooltip_text: self._show_hover_tooltip(event, t))
                    cell_label.bind("<Leave>", self._hide_hover_tooltip)

            all_paid = all(bool(paid_flags_by_debt.get(debt, False)) for debt in debts)
            tk.Label(
                self.rimanenze_table_body,
                text=_format_amount(total),
                bg=FG_POSITIVE if all_paid else BG_TABLE,
                fg="#ffffff" if all_paid else FG_HEADER,
                font=("Segoe UI", 10, "bold"),
                bd=1,
                relief="solid",
                padx=8,
                pady=5,
                anchor="e",
            ).grid(row=ridx, column=col_count - 1, sticky="nsew")

            previous_total = total

        self.rimanenze_table_body.grid_columnconfigure(0, weight=0, minsize=100)
        for cidx in range(1, col_count):
            self.rimanenze_table_body.grid_columnconfigure(cidx, weight=1, minsize=140)

    def _refresh_completion_view(self) -> None:
        self.completion_chart_canvas.delete("all")

        mode = self.completion_mode_var.get().strip()
        target = self.completion_target_var.get().strip()

        if target not in TOTAL_RATES_BY_TARGET:
            self.completion_chart_legend_var.set("Voce non supportata")
            return

        if mode == "Per Rate":
            if self.current_annual_grid is None:
                self.completion_chart_legend_var.set("Dati annuali non ancora caricati")
                return
            paid_by_target = _extract_paid_installments_by_target(self.current_annual_grid)
            total = float(TOTAL_RATES_BY_TARGET[target])
            paid = float(max(0, min(int(total), paid_by_target.get(target, 0))))
            self._render_completion_pie(mode, target, paid, total)
            return

        if mode == "Per Totale":
            if self.current_annual_grid is None:
                self.completion_chart_legend_var.set("Dati annuali non ancora caricati")
                return
            paid_by_target = _extract_paid_total_amounts_by_target(self.current_annual_grid)
            total = TOTAL_AMOUNT_BY_TARGET[target]
            paid = max(0.0, min(total, paid_by_target.get(target, 0.0)))
            self._render_completion_pie(mode, target, paid, total)
            return

        if mode == "Per Capitale":
            total = TOTAL_CAPITAL_BY_TARGET[target]
            paid = max(0.0, min(total, self.capital_paid_by_target.get(target, 0.0)))
            self._render_completion_pie(mode, target, paid, total)
            return

        self.completion_chart_canvas.create_text(
            20,
            20,
            anchor="nw",
            fill=FG,
            font=("Segoe UI", 11, "bold"),
            text="Vista non supportata",
        )
        self.completion_chart_legend_var.set("Nessun grafico per la vista selezionata")

    def _render_completion_pie(self, mode: str, target: str, paid: float, total: float) -> None:
        canvas = self.completion_chart_canvas
        canvas.update_idletasks()
        width = max(canvas.winfo_width(), 420)
        height = max(canvas.winfo_height(), 280)

        diameter = min(width - 80, height - 80)
        diameter = max(180, diameter)
        left = (width - diameter) / 2
        top = 24
        right = left + diameter
        bottom = top + diameter

        remaining = max(0.0, total - paid)
        safe_total = max(1.0, paid + remaining)
        paid_extent = (paid / safe_total) * 360

        # Slice pagato (verde) + residuo (grigio)
        canvas.create_arc(left, top, right, bottom, start=90, extent=-paid_extent, fill="#22c55e", outline=BG_TABLE, width=2)
        canvas.create_arc(
            left,
            top,
            right,
            bottom,
            start=90 - paid_extent,
            extent=-(360 - paid_extent),
            fill="#6b7280",
            outline=BG_TABLE,
            width=2,
        )

        canvas.create_oval(left + 70, top + 70, right - 70, bottom - 70, fill=BG_TABLE, outline=BG_TABLE)
        canvas.create_text(
            width / 2,
            top + (diameter / 2) - 6,
            fill=FG_HEADER,
            font=("Segoe UI", 14, "bold"),
            text=f"{target}",
        )
        canvas.create_text(
            width / 2,
            top + (diameter / 2) + 18,
            fill=FG,
            font=("Segoe UI", 11),
            text=f"Pagato {_format_amount(paid)} / {_format_amount(paid + remaining)}",
        )

        percent_paid = (paid / (paid + remaining)) * 100 if (paid + remaining) else 0
        self.completion_chart_legend_var.set(
            f"{mode} | Pagato: {_format_amount(paid)} ({percent_paid:.1f}%) | Residuo: {_format_amount(remaining)} ({100 - percent_paid:.1f}%)"
        )

    def _render_grid(self, table_body: tk.Frame, grid: SheetGrid) -> None:
        for child in table_body.winfo_children():
            child.destroy()

        if not grid.rows:
            empty = tk.Label(table_body, text="Nessun dato disponibile", bg=BG_TABLE, fg=FG, padx=10, pady=10)
            empty.grid(row=0, column=0, sticky="w")
            return

        column_widths = _compute_column_widths(grid)

        for ridx, row in enumerate(grid.rows):
            is_header = ridx == 0
            for cidx, cell in enumerate(row):
                bg = cell.bg_color or (BG_FRAME if is_header else BG_TABLE)
                fg = _best_fg_for_background(bg, FG if not is_header else FG_HEADER)
                text = cell.value if cell.value else ""
                is_empty = text == ""

                lbl = tk.Label(
                    table_body,
                    text=text,
                    bg=bg,
                    fg=fg,
                    font=("Segoe UI", 10, "bold" if is_header else "normal"),
                    bd=1,
                    relief="solid",
                    padx=2 if is_empty else 8,
                    pady=5,
                    anchor="center",
                )
                lbl.grid(row=ridx, column=cidx, sticky="nsew")

            table_body.grid_rowconfigure(ridx, weight=0)

        for cidx, min_width in enumerate(column_widths):
            table_body.grid_columnconfigure(cidx, weight=0, minsize=min_width)


def _best_fg_for_background(bg: str, default_fg: str) -> str:
    if not bg.startswith("#") or len(bg) != 7:
        return default_fg

    try:
        r = int(bg[1:3], 16)
        g = int(bg[3:5], 16)
        b = int(bg[5:7], 16)
    except ValueError:
        return default_fg

    luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255
    return "#111827" if luminance > 0.62 else "#F8FAFC"


def _compute_column_widths(grid: SheetGrid) -> list[int]:
    if not grid.rows:
        return []

    total_cols = len(grid.rows[0])
    widths: list[int] = []

    for cidx in range(total_cols):
        values = [row[cidx].value.strip() for row in grid.rows if cidx < len(row)]
        has_text = any(values)

        if not has_text:
            widths.append(30)
            continue

        max_len = max(len(v) for v in values)
        approx = 18 + (max_len * 7)
        widths.append(min(max(70, approx), 180))

    return widths


def _build_annual_year_dropdown_values(years: list[str], current_year: int) -> list[str]:
    numeric_years = sorted({year.strip() for year in years if re.fullmatch(r"\d{4}", year.strip())}, key=int)
    if not numeric_years:
        return years

    current = str(current_year)
    non_past = [year for year in numeric_years if int(year) >= current_year]
    past = [year for year in numeric_years if int(year) < current_year]

    ordered: list[str] = []
    if current in non_past:
        ordered.append(current)
        non_past = [year for year in non_past if year != current]

    ordered.extend(non_past)
    if ordered and past:
        ordered.append(ANNUAL_YEARS_SEPARATOR)
    ordered.extend(sorted(past, reverse=True))

    # Mantiene eventuali etichette non-standard senza rompere la UI.
    non_numeric = [year for year in years if not re.fullmatch(r"\d{4}", year.strip())]
    ordered.extend(non_numeric)
    return ordered


def _split_gencal_by_year(grid: SheetGrid) -> dict[str, SheetGrid]:
    if not grid.rows:
        return {}

    header_row = _detect_header_row(grid.rows)
    sections: dict[str, list[list]] = {}
    block_rows: list[list] = []

    def flush_block() -> None:
        nonlocal block_rows
        if not block_rows:
            return

        year = _year_from_block(block_rows)
        if not year:
            block_rows = []
            return

        target = sections.setdefault(year, [])
        if not target and header_row is not None:
            target.append(header_row)
        target.extend(_trim_trailing_empty_rows(block_rows))
        block_rows = []

    for row in grid.rows:
        if _row_is_gray_separator(row):
            flush_block()
            continue

        if header_row is not None and row is header_row:
            continue

        if _row_has_content(row):
            block_rows.append(row)

    flush_block()
    return {year: SheetGrid(name=year, rows=rows) for year, rows in sections.items() if rows}


def _row_is_gray_separator(row) -> bool:
    gray_cells = 0
    for cell in row:
        if _is_gray(cell.bg_color):
            gray_cells += 1
    return gray_cells >= max(1, len(row) // 3)


def _detect_header_row(rows) -> list | None:
    for row in rows[:8]:
        labels = [cell.value.strip().upper() for cell in row if cell.value.strip()]
        if "DATA" in labels and "DR" in labels:
            return row
    return None


def _year_from_block(rows) -> str | None:
    for row in rows:
        year = _extract_year_from_date_cell(row)
        if year:
            return year
    return None


def _extract_year_from_date_cell(row) -> str | None:
    if not row:
        return None

    raw = row[0].value.strip()
    if not raw:
        return None

    match = re.search(r"\b(20\d{2})\b", raw)
    if match:
        return match.group(1)

    # Fallback per date senza separatori (es. 01012026)
    compact = re.sub(r"\D", "", raw)
    if len(compact) >= 8:
        maybe_year = compact[-4:]
        if re.fullmatch(r"20\d{2}", maybe_year):
            return maybe_year

    return None


def _row_has_content(row) -> bool:
    return any(cell.value.strip() or cell.bg_color for cell in row)


def _is_gray(color: str | None) -> bool:
    if not color or not color.startswith("#") or len(color) != 7:
        return False
    try:
        r = int(color[1:3], 16)
        g = int(color[3:5], 16)
        b = int(color[5:7], 16)
    except ValueError:
        return False

    channel_delta = max(r, g, b) - min(r, g, b)
    return channel_delta <= 16 and 70 <= (r + g + b) // 3 <= 220


def _trim_trailing_empty_rows(rows):
    end = len(rows)
    while end > 0:
        row = rows[end - 1]
        if any(cell.value.strip() or cell.bg_color for cell in row):
            break
        end -= 1
    return rows[:end]


def _extract_paid_installments_by_target(grid: SheetGrid) -> dict[str, int]:
    paid_by_target: dict[str, int] = {}
    for target, column_index in TARGET_COLUMN_INDEX.items():
        paid_by_target[target] = _last_green_numeric_value_in_column(grid, column_index)
    return paid_by_target


def _last_green_numeric_value_in_column(grid: SheetGrid, column_index: int) -> int:
    last_value = 0
    for row in grid.rows:
        if column_index >= len(row):
            continue

        cell = row[column_index]
        if not _is_green(cell.bg_color):
            continue

        value = _safe_extract_int(cell.value)
        if value is not None:
            last_value = value

    return last_value


def _safe_extract_int(raw: str) -> int | None:
    match = re.search(r"\d+", raw or "")
    if not match:
        return None
    try:
        return int(match.group(0))
    except ValueError:
        return None


def _is_green(color: str | None) -> bool:
    if not color or not color.startswith("#") or len(color) != 7:
        return False

    try:
        r = int(color[1:3], 16)
        g = int(color[3:5], 16)
        b = int(color[5:7], 16)
    except ValueError:
        return False

    return g >= 90 and g > r + 15 and g > b + 15


def _extract_paid_total_amounts_by_target(grid: SheetGrid) -> dict[str, float]:
    paid_by_target = {target: 0.0 for target in TOTAL_COLUMN_BY_TARGET}
    if not grid.rows:
        return paid_by_target

    paid_row_idx = _find_row_index_with_label(grid, "PAGATO")
    if paid_row_idx is None and len(grid.rows) >= 20:
        # Fallback esplicito richiesto: riga 20 (indice 19)
        paid_row_idx = 19

    if paid_row_idx is None or paid_row_idx >= len(grid.rows):
        return paid_by_target

    row = grid.rows[paid_row_idx]
    column_index_by_code = _find_total_column_indexes(grid)

    for target, column_code in TOTAL_COLUMN_BY_TARGET.items():
        column_index = column_index_by_code.get(column_code)
        if column_index is None or column_index >= len(row):
            continue
        amount = _safe_extract_amount(row[column_index].value)
        if amount is not None:
            paid_by_target[target] = max(0.0, amount)

    return paid_by_target


def _find_total_column_indexes(grid: SheetGrid) -> dict[str, int]:
    wanted = {"RM", "RF", "RC"}
    if not grid.rows:
        return {}

    best_row: list | None = None
    best_hits = 0
    for row in grid.rows[:40]:
        labels = {_normalize_label(cell.value) for cell in row if cell.value.strip()}
        hits = len(labels & wanted)
        if hits > best_hits:
            best_hits = hits
            best_row = row
        if hits == len(wanted):
            break

    result: dict[str, int] = {}
    if best_row is not None:
        for cidx, cell in enumerate(best_row):
            label = _normalize_label(cell.value)
            if label in wanted and label not in result:
                result[label] = cidx

    if len(result) == len(wanted):
        return result

    # Fallback: cerca ogni label su tutte le prime righe utili
    for code in wanted - set(result):
        idx = _find_column_index_with_label(grid, code)
        if idx is not None:
            result[code] = idx
    return result


def _extract_paid_capital_by_target_from_dropbox(capital_file_cache: dict[str, bytes], year: str) -> dict[str, float]:
    paid_by_target = {target: 0.0 for target in CAPITAL_FILE_BY_TARGET}

    for target, dropbox_path in CAPITAL_FILE_BY_TARGET.items():
        try:
            ods_bytes = capital_file_cache.get(dropbox_path)
            if ods_bytes is None:
                ods_bytes = download_ods_bytes(dropbox_path)
                capital_file_cache[dropbox_path] = ods_bytes

            try:
                grid = load_sheet_grid(ods_bytes, year, max_columns=64, max_rows=1200)
            except Exception:
                # Fallback su ultimo foglio anno disponibile se l'anno corrente manca.
                year_sheets = list_year_sheet_names(ods_bytes)
                if not year_sheets:
                    continue
                fallback_year = year if year in year_sheets else year_sheets[-1]
                grid = load_sheet_grid(ods_bytes, fallback_year, max_columns=64, max_rows=1200)

            value = _extract_intersection_total_capitale(grid)
            if value is not None:
                paid_by_target[target] = max(0.0, value)
        except Exception:
            # Non blocca il caricamento principale della UI.
            continue

    return paid_by_target


def _compute_outstanding_by_debt(
    annual_grid: SheetGrid,
    capital_paid_by_target: dict[str, float],
    debt_scope: str,
) -> dict[str, float]:
    if debt_scope == "Solo Capitale":
        return {
            target: max(0.0, TOTAL_CAPITAL_BY_TARGET[target] - capital_paid_by_target.get(target, 0.0))
            for target in TOTAL_CAPITAL_BY_TARGET
        }

    paid_totals = _extract_paid_total_amounts_by_target(annual_grid)
    return {
        target: max(0.0, TOTAL_AMOUNT_BY_TARGET[target] - paid_totals.get(target, 0.0))
        for target in TOTAL_AMOUNT_BY_TARGET
    }


def _extract_platform_amounts_from_dropbox(platform_file_cache: dict[str, bytes], year: str) -> dict[str, float]:
    for path in NEW_TOTAL_INV_PATH_CANDIDATES:
        try:
            file_bytes = platform_file_cache.get(path)
            if file_bytes is None:
                file_bytes = download_ods_bytes(path)
                platform_file_cache[path] = file_bytes

            if path.lower().endswith(".xlsx"):
                return _extract_platform_amounts_from_xlsx_bytes(file_bytes, year)

            try:
                grid = load_sheet_grid(file_bytes, year, max_columns=16, max_rows=200)
            except Exception:
                year_sheets = list_year_sheet_names(file_bytes)
                if not year_sheets:
                    continue
                fallback_year = year if year in year_sheets else year_sheets[-1]
                grid = load_sheet_grid(file_bytes, fallback_year, max_columns=16, max_rows=200)

            return _extract_platform_cells_amounts(grid)
        except Exception:
            continue

    return {name: 0.0 for name in PLATFORM_NAMES}


def _extract_platform_amounts_from_xlsx_bytes(xlsx_bytes: bytes, year: str) -> dict[str, float]:
    from openpyxl import load_workbook

    workbook = load_workbook(BytesIO(xlsx_bytes), data_only=True, read_only=True)
    year_sheets = _list_numeric_xlsx_sheet_names(workbook.sheetnames)
    if not year_sheets:
        return {name: 0.0 for name in PLATFORM_NAMES}

    target_sheet = year if year in year_sheets else year_sheets[-1]
    worksheet = workbook[target_sheet]

    amounts = {name: 0.0 for name in PLATFORM_NAMES}
    for platform_name, (row_idx, col_idx) in PLATFORM_CELL_COORDS.items():
        cell_value = worksheet.cell(row=row_idx + 1, column=col_idx + 1).value
        amount = _safe_extract_amount(str(cell_value)) if cell_value is not None else None
        if amount is not None:
            amounts[platform_name] = max(0.0, amount)

    return amounts


def _list_numeric_xlsx_sheet_names(sheet_names: list[str]) -> list[str]:
    return sorted([name for name in sheet_names if name.isdigit() and len(name) == 4])


def _extract_platform_cells_amounts(grid: SheetGrid) -> dict[str, float]:
    amounts = {name: 0.0 for name in PLATFORM_NAMES}
    for platform_name, (row_idx, col_idx) in PLATFORM_CELL_COORDS.items():
        if row_idx >= len(grid.rows):
            continue
        row = grid.rows[row_idx]
        if col_idx >= len(row):
            continue
        value = _safe_extract_amount(row[col_idx].value)
        if value is not None:
            amounts[platform_name] = max(0.0, value)
    return amounts


def _last_numeric_value_in_column(grid: SheetGrid, column_index: int) -> float:
    fallback = 0.0
    for row in grid.rows:
        if column_index >= len(row):
            continue
        value = _safe_extract_amount(row[column_index].value)
        if value is None:
            continue
        if _is_green(row[column_index].bg_color):
            fallback = value
        else:
            fallback = value
    return max(0.0, fallback)


def _extract_intersection_total_capitale(grid: SheetGrid) -> float | None:
    if not grid.rows:
        return None

    total_row_indexes = _find_row_indexes_with_label(grid, "TOTALE")
    capital_col_indexes = _find_column_indexes_with_label(grid, "CAPITALE")

    for ridx in total_row_indexes:
        if ridx >= len(grid.rows):
            continue
        row = grid.rows[ridx]
        for cidx in capital_col_indexes:
            if cidx >= len(row):
                continue
            value = _safe_extract_amount(row[cidx].value)
            if value is not None:
                return value

    return None


def _find_row_index_with_label(grid: SheetGrid, label: str) -> int | None:
    normalized = _normalize_label(label)
    for ridx, row in enumerate(grid.rows):
        for cell in row:
            if _normalize_label(cell.value) == normalized:
                return ridx
    return None


def _find_row_indexes_with_label(grid: SheetGrid, label: str) -> list[int]:
    normalized = _normalize_label(label)
    indexes: list[int] = []
    for ridx, row in enumerate(grid.rows):
        if any(_normalize_label(cell.value) == normalized for cell in row):
            indexes.append(ridx)
    return indexes


def _find_column_index_with_label(grid: SheetGrid, label: str) -> int | None:
    matches = _find_column_indexes_with_label(grid, label)
    return matches[0] if matches else None


def _find_column_indexes_with_label(grid: SheetGrid, label: str) -> list[int]:
    normalized = _normalize_label(label)
    indexes: list[int] = []
    max_cols = max((len(row) for row in grid.rows), default=0)
    for cidx in range(max_cols):
        for row in grid.rows:
            if cidx >= len(row):
                continue
            if _normalize_label(row[cidx].value) == normalized:
                indexes.append(cidx)
                break
    return indexes


def _normalize_label(value: str) -> str:
    return re.sub(r"\s+", " ", (value or "").strip().upper())


def _safe_extract_amount(raw: str) -> float | None:
    if not raw:
        return None

    cleaned = raw.strip().replace(" ", "").replace("'", "")
    cleaned = cleaned.replace("EUR", "").replace("euro", "")

    if "," in cleaned and "." in cleaned:
        cleaned = cleaned.replace(".", "").replace(",", ".")
    elif "," in cleaned:
        cleaned = cleaned.replace(",", ".")

    match = re.search(r"-?\d+(?:\.\d+)?", cleaned)
    if not match:
        return None

    try:
        return float(match.group(0))
    except ValueError:
        return None


def _format_amount(value: float) -> str:
    if abs(value - round(value)) < 1e-9:
        return f"{int(round(value)):,}".replace(",", ".")
    formatted = f"{value:,.2f}"
    return formatted.replace(",", "#").replace(".", ",").replace("#", ".")


def _format_signed_amount(value: float) -> str:
    sign = "+" if value > 0 else ""
    return f"{sign}{_format_amount(value)}"


def _detect_gencal_column_map(grid: SheetGrid) -> dict[tuple[str, str], int]:
    """Tenta di rilevare le colonne del GenCal per (debt, field) dove field è 'capitale' o 'rata'.
    Ritorna un dict {(debt_key, field): column_index}."""
    result: dict[tuple[str, str], int] = {}
    if not grid.rows:
        return result

    # Cerca le prime righe header (fino a riga 5)
    header_rows = grid.rows[:5]

    # Per ogni riga header, cerca etichette sezione debito e sub-etichette
    # Strategia: trova colonne di sezione (RM/RF/RC) poi scansiona sotto-intestazioni
    section_start: dict[str, int] = {}  # debt -> colonna iniziale sezione

    for row in header_rows:
        for cidx, cell in enumerate(row):
            label = _normalize_label(cell.value)
            if not label:
                continue
            for debt, labels in GENCAL_SECTION_LABELS.items():
                if label in labels and debt not in section_start:
                    section_start[debt] = cidx

    if not section_start:
        return result

    # Determina i range di colonne per ciascuna sezione
    sorted_sections = sorted(section_start.items(), key=lambda x: x[1])
    debt_col_ranges: dict[str, tuple[int, int]] = {}
    for i, (debt, start_col) in enumerate(sorted_sections):
        end_col = sorted_sections[i + 1][1] if i + 1 < len(sorted_sections) else len(grid.rows[0])
        debt_col_ranges[debt] = (start_col, end_col)

    # Cerca sub-etichette CAPITALE / RATA / TOTALE / RESIDUO all'interno di ogni range
    field_patterns = {
        "capitale": {"CAPITALE", "CAP"},
        "rata": {"RATA", "TOTALE", "TOT", "IMPORTO", "PAGAMENTO"},
        "manc": {"MANC", "MANCANTE", "RESIDUO"},
    }

    for row in header_rows:
        for debt, (start_col, end_col) in debt_col_ranges.items():
            for cidx in range(start_col, min(end_col, len(row))):
                label = _normalize_label(row[cidx].value)
                if not label:
                    continue
                for field, patterns in field_patterns.items():
                    if any(p in label for p in patterns):
                        key = (debt, field)
                        if key not in result:
                            result[key] = cidx

    return result


def _extract_monthly_rimanenze(
    grid: SheetGrid,
    debts: tuple[str, ...],
    scope: str,
    initial_remaining: dict[str, float],
    year: str,
    fill_full_year: bool = False,
) -> list[tuple[str, dict[str, float], dict[str, bool], dict[str, float]]]:
    """Estrae rimanenze mensili e flag pagamento per i debiti selezionati.
    Ritorna lista di (label_mese, {debt: remaining_balance}, {debt: paid_in_month}).
    """
    if not grid.rows:
        return []

    field = "capitale" if scope == "Capitale" else "rata"
    col_map = _detect_gencal_column_map(grid)

    # Modalita' GenCal attuale: colonne globali + discriminante DR per debito.
    dr_col_idx = _find_column_index_in_first_rows(grid, "DR")
    global_field_idx = _find_gencal_global_field_column(grid, field)
    global_manc_idx = _find_gencal_global_field_column(grid, "manc")
    use_dr_mode = dr_col_idx is not None and global_field_idx is not None

    # Verifica se abbiamo colonne per i debiti selezionati (fallback per formato a sezioni).
    has_col = {debt: (debt, field) in col_map for debt in debts}

    # Calcola la rata mensile fissa stimata come fallback se non troviamo le colonne
    monthly_payment_fallback: dict[str, float] = {}
    if scope == "Capitale":
        total_months_by_debt = {d: TOTAL_RATES_BY_TARGET[d] for d in debts}
        for d in debts:
            monthly_payment_fallback[d] = TOTAL_CAPITAL_BY_TARGET[d] / max(1, total_months_by_debt[d])
    else:
        total_months_by_debt = {d: TOTAL_RATES_BY_TARGET[d] for d in debts}
        for d in debts:
            monthly_payment_fallback[d] = TOTAL_AMOUNT_BY_TARGET[d] / max(1, total_months_by_debt[d])

    monthly_paid: dict[str, dict[str, float]] = {}
    monthly_manc: dict[str, dict[str, float]] = {}
    monthly_paid_flags: dict[str, dict[str, bool]] = {}
    monthly_labels: dict[str, str] = {}

    for row in grid.rows:
        if not row:
            continue

        date_val = row[0].value.strip() if len(row) > 0 else ""
        month_info = _parse_month_key_label(date_val)
        if month_info is None:
            continue

        month_key, month_label = month_info
        monthly_labels[month_key] = month_label
        month_bucket = monthly_paid.setdefault(month_key, {debt: 0.0 for debt in debts})
        manc_bucket = monthly_manc.setdefault(month_key, {})
        paid_flags_bucket = monthly_paid_flags.setdefault(month_key, {debt: False for debt in debts})

        if use_dr_mode:
            # In questo formato ogni riga appartiene a un solo debito, indicato in DR.
            if dr_col_idx >= len(row):
                continue
            debt = _debt_from_gencal_dr(row[dr_col_idx].value)
            if debt is None or debt not in debts:
                continue

            paid_signal = False
            if global_manc_idx is not None and global_manc_idx < len(row):
                manc_amount = _safe_extract_amount(row[global_manc_idx].value)
                if manc_amount is not None and manc_amount >= 0:
                    manc_bucket[debt] = manc_amount
                if _is_green(row[global_manc_idx].bg_color):
                    paid_signal = True

            if global_field_idx < len(row):
                amount = _safe_extract_amount(row[global_field_idx].value)
                if amount is not None and amount > 0:
                    month_bucket[debt] += amount
                    paid_signal = True
                if _is_green(row[global_field_idx].bg_color):
                    paid_signal = True

            if paid_signal:
                paid_flags_bucket[debt] = True
            continue

        # Fallback: formato GenCal a sezioni (colonne dedicate per debito)
        for debt in debts:
            paid_signal = False
            manc_key = (debt, "manc")
            if manc_key in col_map:
                m_idx = col_map[manc_key]
                if m_idx < len(row):
                    manc_amount = _safe_extract_amount(row[m_idx].value)
                    if manc_amount is not None and manc_amount >= 0:
                        manc_bucket[debt] = manc_amount
                    if _is_green(row[m_idx].bg_color):
                        paid_signal = True

            col_key = (debt, field)
            if col_key in col_map:
                cidx = col_map[col_key]
                if cidx < len(row):
                    amount = _safe_extract_amount(row[cidx].value)
                    if amount is not None and amount > 0:
                        month_bucket[debt] += amount
                        paid_signal = True
                    if _is_green(row[cidx].bg_color):
                        paid_signal = True

            if paid_signal:
                paid_flags_bucket[debt] = True

    if not monthly_paid and not fill_full_year:
        return []

    remaining: dict[str, float] = {d: initial_remaining.get(d, 0.0) for d in debts}
    result: list[tuple[str, dict[str, float], dict[str, bool], dict[str, float]]] = []

    if fill_full_year:
        month_sequence = _build_month_sequence_for_year(year)
    else:
        month_sequence = [(mkey, monthly_labels[mkey]) for mkey in sorted(monthly_paid.keys())]

    # Ordinamento cronologico stabile grazie alla chiave YYYY-MM.
    for month_key, month_label in month_sequence:
        row_remaining: dict[str, float] = {}
        month_bucket = monthly_paid.get(month_key, {})
        manc_bucket = monthly_manc.get(month_key, {})
        month_paid_flags = monthly_paid_flags.get(month_key, {})

        for debt in debts:
            if debt in manc_bucket:
                remaining[debt] = max(0.0, manc_bucket[debt])
            elif has_col.get(debt, False):
                paid_amount = month_bucket.get(debt, 0.0)
                if paid_amount > 0:
                    remaining[debt] = max(0.0, remaining[debt] - paid_amount)
            elif month_key in monthly_paid:
                remaining[debt] = max(0.0, remaining[debt] - monthly_payment_fallback.get(debt, 0.0))

            row_remaining[debt] = remaining[debt]

        # Flag mese: pagato se esiste movimento positivo o cella marcata verde.
        row_flags = {debt: bool(month_paid_flags.get(debt, False)) for debt in debts}
        row_paid_amounts = {
            debt: month_bucket.get(debt, 0.0) if row_flags.get(debt, False) else 0.0
            for debt in debts
        }
        label = month_label or monthly_labels.get(month_key, month_key)
        result.append((label, row_remaining, row_flags, row_paid_amounts))

    return result


def _build_month_sequence_for_year(year: str) -> list[tuple[str, str]]:
    if not re.fullmatch(r"\d{4}", year or ""):
        return []
    y = int(year)
    return [(f"{y:04d}-{m:02d}", f"{MONTH_NAMES_IT[m - 1]} {y}") for m in range(1, 13)]


def _extract_monthly_capital_rimanenze_from_dedicated_files(
    debts: tuple[str, ...],
    year: str,
    initial_remaining: dict[str, float],
    capital_file_cache: dict[str, bytes],
) -> list[tuple[str, dict[str, float], dict[str, bool], dict[str, float]]]:
    month_sequence = _build_month_sequence_for_year(year)
    if not month_sequence:
        return []

    current_year = str(datetime.now().year)
    selected_year_int = int(year)
    current_year_int = int(current_year)

    paid_amounts_by_debt: dict[str, dict[str, float]] = {debt: {} for debt in debts}
    paid_flags_by_debt: dict[str, dict[str, bool]] = {debt: {} for debt in debts}

    for debt in debts:
        dropbox_path = CAPITAL_FILE_BY_TARGET.get(debt)
        if not dropbox_path:
            continue
        try:
            ods_bytes = _load_dedicated_file_bytes(capital_file_cache, dropbox_path)
            grid = _load_dedicated_year_grid(ods_bytes, year)
            amounts, flags = _extract_dedicated_capital_amounts_and_paid_flags(grid, year)
            paid_amounts_by_debt[debt] = amounts
            paid_flags_by_debt[debt] = flags
        except Exception:
            continue

    # Ricostruisce la rimanenza mese per mese partendo dalla rimanenza attuale.
    # Le cifre decrescono sempre in base al piano capitale mensile; il colore verde
    # resta legato solo ai mesi realmente pagati (celle verdi nei file dedicati).
    result: list[tuple[str, dict[str, float], dict[str, bool], dict[str, float]]] = []
    running_remaining_by_debt: dict[str, float] = {}

    for debt in debts:
        debt_remaining = max(0.0, initial_remaining.get(debt, 0.0))
        dropbox_path = CAPITAL_FILE_BY_TARGET.get(debt)

        # Per gli anni futuri, porta il residuo da "oggi" all'inizio dell'anno selezionato.
        if dropbox_path and selected_year_int > current_year_int:
            try:
                ods_bytes = _load_dedicated_file_bytes(capital_file_cache, dropbox_path)

                for bridge_year in range(current_year_int, selected_year_int):
                    bridge_year_str = str(bridge_year)
                    bridge_grid = _load_dedicated_year_grid(ods_bytes, bridge_year_str)
                    bridge_amounts, bridge_flags = _extract_dedicated_capital_amounts_and_paid_flags(bridge_grid, bridge_year_str)
                    if bridge_year == current_year_int:
                        # Nell'anno corrente il residuo e' gia' al netto del pagato: sottraggo solo il non ancora pagato.
                        to_subtract = sum(amount for mkey, amount in bridge_amounts.items() if not bridge_flags.get(mkey, False))
                    else:
                        # Anni ponte pieni: sottraggo l'intero piano dell'anno.
                        to_subtract = sum(bridge_amounts.values())
                    debt_remaining = max(0.0, debt_remaining - to_subtract)
            except Exception:
                pass

        # Nell'anno corrente, mostra il valore mensile "dopo la rata del mese":
        # riallineo al saldo prima di gennaio aggiungendo il capitale gia' pagato quest'anno.
        if selected_year_int == current_year_int:
            amounts = paid_amounts_by_debt.get(debt, {})
            flags = paid_flags_by_debt.get(debt, {})
            paid_so_far = sum(amount for mkey, amount in amounts.items() if flags.get(mkey, False))
            debt_remaining = max(0.0, debt_remaining + paid_so_far)

        running_remaining_by_debt[debt] = debt_remaining

    for mkey, mlabel in month_sequence:
        row_values: dict[str, float] = {}
        row_flags: dict[str, bool] = {}
        row_paid_amounts: dict[str, float] = {}
        for debt in debts:
            amounts = paid_amounts_by_debt.get(debt, {})
            flags = paid_flags_by_debt.get(debt, {})

            month_capital = amounts.get(mkey, 0.0)
            if month_capital > 0:
                running_remaining_by_debt[debt] = max(0.0, running_remaining_by_debt[debt] - month_capital)

            row_values[debt] = running_remaining_by_debt[debt]
            row_flags[debt] = bool(flags.get(mkey, False))
            row_paid_amounts[debt] = month_capital if row_flags[debt] else 0.0

        result.append((mlabel, row_values, row_flags, row_paid_amounts))

    return result


def _load_dedicated_file_bytes(capital_file_cache: dict[str, bytes], dropbox_path: str) -> bytes:
    ods_bytes = capital_file_cache.get(dropbox_path)
    if ods_bytes is None:
        ods_bytes = download_ods_bytes(dropbox_path)
        capital_file_cache[dropbox_path] = ods_bytes
    return ods_bytes


def _load_dedicated_year_grid(ods_bytes: bytes, year: str) -> SheetGrid:
    try:
        return load_sheet_grid(ods_bytes, year, max_columns=32, max_rows=400)
    except Exception:
        years = list_year_sheet_names(ods_bytes)
        if not years:
            raise
        fallback_year = year if year in years else years[-1]
        return load_sheet_grid(ods_bytes, fallback_year, max_columns=32, max_rows=400)


def _extract_dedicated_capital_amounts_and_paid_flags(grid: SheetGrid, year: str) -> tuple[dict[str, float], dict[str, bool]]:
    month_idx, capital_idx, _year_idx = _find_dedicated_schedule_columns(grid)
    if month_idx is None or capital_idx is None:
        return {}, {}

    month_amounts: dict[str, float] = {}
    month_paid_flags: dict[str, bool] = {}

    for row in grid.rows:
        if month_idx >= len(row) or capital_idx >= len(row):
            continue

        month_raw = row[month_idx].value
        # Usa sempre l'anno del foglio selezionato: in alcuni file la colonna ANNO
        # puo' contenere valori incoerenti e causare slittamento al foglio precedente.
        month_info = _parse_month_key_label(month_raw, fallback_year=year)
        if month_info is None:
            continue

        month_key, _ = month_info
        amount = _safe_extract_amount(row[capital_idx].value)
        if amount is None or amount <= 0:
            continue

        month_amounts[month_key] = amount
        # Pagato solo se la riga/cella è marcata in verde nel file dedicato.
        month_paid_flags[month_key] = _is_green(row[capital_idx].bg_color) or _is_green(row[month_idx].bg_color)

    return month_amounts, month_paid_flags


def _find_dedicated_schedule_columns(grid: SheetGrid) -> tuple[int | None, int | None, int | None]:
    for ridx, row in enumerate(grid.rows[:8]):
        month_idx = None
        capital_idx = None
        year_idx = None
        for cidx, cell in enumerate(row):
            label = _normalize_label(cell.value)
            if not label:
                continue
            if month_idx is None and "MESE" in label:
                month_idx = cidx
            if capital_idx is None and "CAPITALE" in label:
                capital_idx = cidx
            if year_idx is None and label == "ANNO":
                year_idx = cidx
        if month_idx is not None and capital_idx is not None:
            return month_idx, capital_idx, year_idx
    return None, None, None


def _find_column_index_in_first_rows(grid: SheetGrid, label: str, header_scan_rows: int = 8) -> int | None:
    normalized = _normalize_label(label)
    for row in grid.rows[:header_scan_rows]:
        for cidx, cell in enumerate(row):
            if _normalize_label(cell.value) == normalized:
                return cidx
    return None


def _find_gencal_global_field_column(grid: SheetGrid, field: str) -> int | None:
    patterns_by_field = {
        "capitale": ("CAPITALE", "CAP"),
        "rata": ("RATA", "TOTALE", "TOT", "IMPORTO", "PAGAMENTO"),
        "manc": ("MANC", "MANCANTE", "RESIDUO"),
    }
    patterns = patterns_by_field.get(field, ())
    if not patterns:
        return None

    for row in grid.rows[:8]:
        for cidx, cell in enumerate(row):
            label = _normalize_label(cell.value)
            if label and any(pat in label for pat in patterns):
                return cidx
    return None


def _parse_month_key_label(raw_date: str, fallback_year: str | None = None) -> tuple[str, str] | None:
    text = (raw_date or "").strip()
    if not text:
        return None

    # Formati comuni: dd/mm/yyyy, dd-mm-yyyy, yyyy-mm-dd o varianti con testo.
    match = re.search(r"(20\d{2})[^\d]?([01]?\d)", text)
    if match:
        year = int(match.group(1))
        month = int(match.group(2))
        if 1 <= month <= 12:
            return f"{year:04d}-{month:02d}", f"{MONTH_NAMES_IT[month - 1]} {year}"

    compact = re.sub(r"\D", "", text)
    if len(compact) >= 8:
        # ddmmyyyy
        dmy_year = int(compact[-4:])
        dmy_month = int(compact[-6:-4])
        if 2000 <= dmy_year <= 2099 and 1 <= dmy_month <= 12:
            return f"{dmy_year:04d}-{dmy_month:02d}", f"{MONTH_NAMES_IT[dmy_month - 1]} {dmy_year}"
        # yyyymmdd
        ymd_year = int(compact[:4])
        ymd_month = int(compact[4:6])
        if 2000 <= ymd_year <= 2099 and 1 <= ymd_month <= 12:
            return f"{ymd_year:04d}-{ymd_month:02d}", f"{MONTH_NAMES_IT[ymd_month - 1]} {ymd_year}"

    # Formati tipici dei file dedicati: "01.Gennaio" o "01. Gennaio" (senza anno)
    if fallback_year and re.fullmatch(r"\d{4}", fallback_year):
        month_number = _parse_month_number_it(text)
        if month_number is not None:
            year = int(fallback_year)
            return f"{year:04d}-{month_number:02d}", f"{MONTH_NAMES_IT[month_number - 1]} {year}"

    return None


def _parse_month_number_it(raw: str) -> int | None:
    text = _normalize_label(raw)
    if not text:
        return None

    # Preferisce un prefisso numerico esplicito (es. 01.Gennaio)
    num_match = re.match(r"\s*(\d{1,2})", raw or "")
    if num_match:
        month_num = int(num_match.group(1))
        if 1 <= month_num <= 12:
            return month_num

    month_names = {
        "GEN": 1,
        "GENNAIO": 1,
        "FEB": 2,
        "FEBBRAIO": 2,
        "MAR": 3,
        "MARZO": 3,
        "APR": 4,
        "APRILE": 4,
        "MAG": 5,
        "MAGGIO": 5,
        "GIU": 6,
        "GIUGNO": 6,
        "LUG": 7,
        "LUGLIO": 7,
        "AGO": 8,
        "AGOSTO": 8,
        "SET": 9,
        "SETTEMBRE": 9,
        "OTT": 10,
        "OTTOBRE": 10,
        "NOV": 11,
        "NOVEMBRE": 11,
        "DIC": 12,
        "DICEMBRE": 12,
    }
    for token, month_num in month_names.items():
        if re.search(rf"\b{token}\b", text):
            return month_num
    return None


def _parse_year_value(raw: str) -> str | None:
    text = (raw or "").strip()
    match = re.search(r"\b(20\d{2})\b", text)
    if match:
        return match.group(1)
    return None


def _debt_from_gencal_dr(raw: str) -> str | None:
    key = _normalize_label(raw)
    if not key:
        return None
    for alias, debt in GENCAL_DR_TO_DEBT.items():
        if alias in key:
            return debt
    return None


def run_desktop_app(dropbox_path: str = DEFAULT_DROPBOX_PATH) -> None:
    app = DebtsDesktopApp(dropbox_path=dropbox_path)
    app.mainloop()


if __name__ == "__main__":
    run_desktop_app()

