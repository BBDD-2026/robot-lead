import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
import re
import subprocess
from datetime import datetime
import matplotlib
matplotlib.use("TkAgg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.gridspec import GridSpec

# ── Colores dark mode ──────────────────────────────────────────
BG      = "#1e1e2e"
BG2     = "#2a2a3e"
BG3     = "#313244"
FG      = "#cdd6f4"
ACCENT  = "#89b4fa"
GREEN   = "#a6e3a1"
RED     = "#f38ba8"
YELLOW  = "#f9e2af"
GRAY    = "#6c7086"
WHITE   = "#ffffff"

MESES = {
    "01": "Enero", "02": "Febrero", "03": "Marzo", "04": "Abril",
    "05": "Mayo",  "06": "Junio",   "07": "Julio", "08": "Agosto",
    "09": "Septiembre", "10": "Octubre", "11": "Noviembre", "12": "Diciembre"
}

# Columnas a mostrar en el acumulado
COLS_VISTA = ["DB_ID", "record_id", "customer_firstname", "customer_lastname",
              "PhoneNumber", "City", "Province", "_tipo", "_periodo"]


def decode_db_id(value) -> tuple[str, str]:
    """Decodifica DB_ID tipo '0326P' → ('Porta', 'Marzo 2026')"""
    if not value or not isinstance(value, str):
        return ("?", "?")
    v = value.strip()
    if len(v) < 5:
        return ("?", v)
    mm   = v[:2]
    yy   = v[2:4]
    tipo = v[4].upper()
    mes  = MESES.get(mm, mm)
    year = f"20{yy}"
    nombre = "Porta" if tipo == "P" else "Baf" if tipo == "B" else tipo
    return (nombre, f"{mes} {year}")


def is_valid_name(value) -> bool:
    if not value or not isinstance(value, str):
        return False
    v = value.strip()
    if not v or len(v) > 30 or "@" in v:
        return False
    if not re.fullmatch(r'[a-zA-Z\u00e0-\u00ff ]+', v):
        return False
    if any(len(p) < 2 for p in v.split()):
        return False
    return True


class RobotLeadApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Robot Lead — Procesador de Archivos")
        self.root.configure(bg=BG)
        self.root.geometry("960x620")
        self.root.resizable(True, False)

        self.df = None
        self.filepath = None
        self.df_processed = None
        self.last_csv_path = None

        # Acumuladores separados por tipo
        self.acum: dict[str, list[pd.DataFrame]] = {"Porta": [], "Baf": []}

        # Registro diario de lotes procesados
        # cada item: {tipo, archivo, hora, si, dupl, invalid, sin_datos}
        self.lotes: list[dict] = []

        self._build_ui()

    # ── UI ────────────────────────────────────────────────────────
    def _build_ui(self):
        hdr = tk.Frame(self.root, bg=BG, pady=12)
        hdr.pack(fill="x")
        tk.Label(hdr, text="ROBOT LEAD", font=("Segoe UI", 22, "bold"),
                 bg=BG, fg=ACCENT).pack()
        tk.Label(hdr, text="Procesador de Archivos Excel · Porta / Baf",
                 font=("Segoe UI", 9), bg=BG, fg=GRAY).pack()

        load = tk.Frame(self.root, bg=BG2, pady=12, padx=20)
        load.pack(fill="x", padx=20, pady=(0, 8))
        tk.Button(load, text="  Cargar Archivo Excel  ",
                  command=self._load_file,
                  bg=ACCENT, fg=WHITE, font=("Segoe UI", 10, "bold"),
                  relief="flat", padx=14, pady=6, cursor="hand2"
                  ).pack(side="left")
        self.lbl_file = tk.Label(load, text="Ningun archivo cargado",
                                  font=("Segoe UI", 10), bg=BG2, fg=GRAY)
        self.lbl_file.pack(side="left", padx=14)

        info_frame = tk.Frame(self.root, bg=BG, padx=20)
        info_frame.pack(fill="x")
        self._cards = {}
        for key, title in [("filas", "Total Filas"), ("si", "Subir: si"),
                            ("dupl", "Subir: dupl"), ("invalid", "Subir: Invalid"),
                            ("sin_datos", "Sin Datos")]:
            card = tk.Frame(info_frame, bg=BG3, padx=14, pady=10)
            card.pack(side="left", padx=(0, 6), pady=8)
            tk.Label(card, text=title, font=("Segoe UI", 8), bg=BG3, fg=GRAY).pack()
            lbl = tk.Label(card, text="—", font=("Segoe UI", 15, "bold"),
                           bg=BG3, fg=FG)
            lbl.pack()
            self._cards[key] = lbl

        # Botones (side=bottom antes del log)
        act = tk.Frame(self.root, bg=BG, padx=20, pady=8)
        act.pack(fill="x", side="bottom")

        self.btn_proc = tk.Button(act, text="  Procesar  ",
                                  command=self._process,
                                  bg=GRAY, fg=WHITE, font=("Segoe UI", 10, "bold"),
                                  relief="flat", padx=14, pady=6,
                                  state="disabled", cursor="hand2")
        self.btn_proc.pack(side="left", padx=(0, 8))

        self.btn_save = tk.Button(act, text="  Guardar  ",
                                  command=self._save,
                                  bg=GRAY, fg=WHITE, font=("Segoe UI", 10, "bold"),
                                  relief="flat", padx=14, pady=6,
                                  state="disabled", cursor="hand2")
        self.btn_save.pack(side="left", padx=(0, 8))

        self.btn_download = tk.Button(act, text="  Abrir Carpeta  ",
                                      command=self._open_folder,
                                      bg=GRAY, fg=WHITE, font=("Segoe UI", 10, "bold"),
                                      relief="flat", padx=14, pady=6,
                                      state="disabled", cursor="hand2")
        self.btn_download.pack(side="left", padx=(0, 8))

        self.btn_acum = tk.Button(act, text="  Ver Acumulado  ",
                                  command=self._ver_acumulado,
                                  bg=GRAY, fg=WHITE, font=("Segoe UI", 10, "bold"),
                                  relief="flat", padx=14, pady=6,
                                  state="disabled", cursor="hand2")
        self.btn_acum.pack(side="left", padx=(0, 8))

        self.btn_muestreo = tk.Button(act, text="  Muestreo  ",
                                      command=self._ver_muestreo,
                                      bg=GRAY, fg=WHITE, font=("Segoe UI", 10, "bold"),
                                      relief="flat", padx=14, pady=6,
                                      state="disabled", cursor="hand2")
        self.btn_muestreo.pack(side="left", padx=(0, 20))

        self.lbl_status = tk.Label(act, text="", font=("Segoe UI", 9),
                                   bg=BG, fg=GREEN)
        self.lbl_status.pack(side="left")

        log_outer = tk.Frame(self.root, bg=BG, padx=20)
        log_outer.pack(fill="both", expand=True)
        tk.Label(log_outer, text="Registro de procesamiento",
                 font=("Segoe UI", 9), bg=BG, fg=GRAY).pack(anchor="w")
        txt_frame = tk.Frame(log_outer, bg=BG2)
        txt_frame.pack(fill="both", expand=True)
        self.log_box = tk.Text(txt_frame, bg=BG2, fg=FG, font=("Consolas", 9),
                               relief="flat", state="disabled", wrap="word",
                               insertbackground=FG)
        scroll = tk.Scrollbar(txt_frame, command=self.log_box.yview, bg=BG2)
        self.log_box.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        self.log_box.pack(fill="both", expand=True, padx=6, pady=4)

    # ── Helpers ───────────────────────────────────────────────────
    def _log(self, msg: str):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", f"{msg}\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def _set_card(self, key: str, value, color: str = FG):
        self._cards[key].configure(text=str(value), fg=color)

    def _detect_tipo(self) -> str:
        """Detecta Porta o Baf por nombre de archivo."""
        name = os.path.basename(self.filepath or "").lower()
        if "porta" in name:
            return "Porta"
        if "baf" in name:
            return "Baf"
        return "Otro"

    # ── Cargar ────────────────────────────────────────────────────
    def _load_file(self):
        path = filedialog.askopenfilename(
            title="Seleccionar archivo Excel",
            filetypes=[("Excel", "*.xlsx *.xls")],
            initialdir=os.path.expanduser("~/Desktop/Robot Leads")
        )
        if not path:
            return
        try:
            self.df = pd.read_excel(path)
            self.filepath = path
            self.df_processed = None
            name = os.path.basename(path)
            self.lbl_file.configure(text=f"✓  {name}", fg=GREEN)
            for k in self._cards:
                self._set_card(k, "—", FG)
            self._log(f"{'─'*52}")
            self._log(f"Archivo cargado: {name}")
            self._log(f"Filas: {len(self.df)}  |  Columnas: {len(self.df.columns)}")
            self.btn_proc.configure(state="normal", bg=ACCENT, fg=WHITE)
            self.btn_save.configure(state="disabled", bg=GRAY, fg=WHITE)
            self.lbl_status.configure(text="")
        except Exception as exc:
            messagebox.showerror("Error al cargar", str(exc))

    # ── Procesar ──────────────────────────────────────────────────
    def _process(self):
        if self.df is None:
            return

        df = self.df.copy()
        self._log(f"{'─'*52}")
        self._log("Iniciando procesamiento...")

        if "Ruta" in df.columns and "Province" in df.columns:
            df["Province"] = df["Ruta"]
            self._log("✓  Ruta  →  Province")

        if "Localidad" in df.columns and "City" in df.columns:
            df["City"] = df["Localidad"]
            self._log("✓  Localidad  →  City")

        sin_datos_count = 0
        if "customer_firstname" in df.columns:
            df["customer_firstname"] = df["customer_firstname"].apply(
                lambda x: x.strip() if is_valid_name(x) else "Sin Datos"
            )
            sin_datos_count = (df["customer_firstname"] == "Sin Datos").sum()
            self._log(f"✓  customer_firstname limpiado  →  {sin_datos_count} 'Sin Datos'")

        if "customer_firstname" in df.columns:
            df = df.sort_values("customer_firstname", ascending=True,
                                ignore_index=True)
            self._log("✓  Ordenado A→Z por customer_firstname")

        df["Subir"] = ""

        dupl_count = 0
        if "Gen_Insert" in df.columns:
            mask_dupl = df["Gen_Insert"].notna()
            df.loc[mask_dupl, "Subir"] = "dupl"
            dupl_count = int(mask_dupl.sum())
            self._log(f"✓  Gen_Insert con fecha  →  {dupl_count} marcados 'dupl'")

        invalid_count = 0
        if "Ruta" in df.columns:
            mask_inv = (
                df["Ruta"].isna() |
                (df["Ruta"].astype(str).str.strip().str.upper() == "NULL")
            ) & (df["Subir"] == "")
            df.loc[mask_inv, "Subir"] = "Invalid"
            invalid_count = int(mask_inv.sum())
            self._log(f"✓  Ruta NULL  →  {invalid_count} marcados 'Invalid'")

        mask_si = df["Subir"] == ""
        df.loc[mask_si, "Subir"] = "si"
        si_count = int(mask_si.sum())
        self._log(f"✓  Resto  →  {si_count} marcados 'si'")

        self._log(f"{'─'*52}")
        self._log(f"  Total: {len(df)}  |  si={si_count}  dupl={dupl_count}  Invalid={invalid_count}")

        self.df_processed = df
        self._set_card("filas", len(df))
        self._set_card("si", si_count, GREEN)
        self._set_card("dupl", dupl_count, YELLOW)
        self._set_card("invalid", invalid_count, RED)
        self._set_card("sin_datos", sin_datos_count, RED if sin_datos_count else GREEN)

        self._log(f"{'─'*52}")
        self.btn_save.configure(state="normal", bg=GREEN, fg=BG)
        self.lbl_status.configure(text="Listo para guardar", fg=YELLOW)

    # ── Guardar ───────────────────────────────────────────────────
    def _save(self):
        if self.df_processed is None:
            return
        try:
            df = self.df_processed
            base, _ = os.path.splitext(self.filepath)
            ts = datetime.now().strftime("%d%m_%H%M")
            folder = os.path.dirname(self.filepath)
            orig_name = os.path.basename(self.filepath)

            df.to_excel(self.filepath, index=False)
            self._log(f"✓  Sobreescrito: {orig_name}")

            excel_path = f"{base}_procesado_{ts}.xlsx"
            df.to_excel(excel_path, index=False)
            self._log(f"✓  Excel procesado: {os.path.basename(excel_path)}")

            df_si = df[df["Subir"] == "si"].copy().reset_index(drop=True)
            df_si["record_id"] = range(1, len(df_si) + 1)
            df_si["chain_id"]  = df_si["record_id"]

            if "record_id" in df_si.columns:
                start_col = df_si.columns.get_loc("record_id")
                df_si = df_si.iloc[:, start_col:]
            if "DB_ID" in df_si.columns:
                end_col = df_si.columns.get_loc("DB_ID") + 1
                df_si = df_si.iloc[:, :end_col]

            fecha_match = re.search(r'_(\d{4})', orig_name)
            fecha_code  = fecha_match.group(1) if fecha_match else ts[:4]
            tipo_str    = self._detect_tipo()
            csv_path    = os.path.join(folder, f"Lote_Leads_{tipo_str}_{fecha_code}.csv")
            df_si.to_csv(csv_path, index=False, sep=",", encoding="latin-1", errors="replace")
            self._log(f"✓  CSV generado ({len(df_si)} filas): {os.path.basename(csv_path)}")

            # ── Acumular filas "si" con metadata ─────────────────
            tipo = self._detect_tipo()
            df_acum = df[df["Subir"] == "si"].copy()

            # Decodificar DB_ID para columnas extra
            if "DB_ID" in df_acum.columns:
                decoded = df_acum["DB_ID"].apply(
                    lambda x: decode_db_id(str(x)) if pd.notna(x) else ("?", "?")
                )
                df_acum["_tipo"]    = decoded.apply(lambda x: x[0])
                df_acum["_periodo"] = decoded.apply(lambda x: x[1])
            else:
                df_acum["_tipo"]    = tipo
                df_acum["_periodo"] = "?"

            if tipo in self.acum:
                self.acum[tipo].append(df_acum)
            else:
                self.acum["Porta"].append(df_acum)

            total_acum = sum(len(d) for lst in self.acum.values() for d in lst)
            self._log(f"✓  Acumulado actualizado: {total_acum} filas totales")

            # Registrar lote diario
            self.lotes.append({
                "tipo":      tipo,
                "archivo":   orig_name,
                "hora":      datetime.now().strftime("%H:%M"),
                "si":        int((df["Subir"] == "si").sum()),
                "dupl":      int((df["Subir"] == "dupl").sum()),
                "invalid":   int((df["Subir"] == "Invalid").sum()),
                "sin_datos": int((df.get("customer_firstname", pd.Series()) == "Sin Datos").sum()),
            })

            self.last_csv_path = csv_path
            self.lbl_status.configure(text="✓  Guardado", fg=GREEN)
            self.btn_download.configure(state="normal", bg=ACCENT, fg=WHITE)
            self.btn_acum.configure(state="normal", bg=YELLOW, fg=BG)
            self.btn_muestreo.configure(state="normal", bg="#cba6f7", fg=BG)

            messagebox.showinfo(
                "Guardado correctamente",
                f"Archivos generados en:\n{folder}\n\n"
                f"• {orig_name}  (sobreescrito)\n"
                f"• {os.path.basename(excel_path)}\n"
                f"• {os.path.basename(csv_path)}  ({len(df_si)} filas 'si')"
            )
        except Exception as exc:
            messagebox.showerror("Error al guardar", str(exc))

    # ── Abrir carpeta ─────────────────────────────────────────────
    def _open_folder(self):
        folder = os.path.dirname(self.filepath) if self.filepath else None
        if not folder:
            return
        if self.last_csv_path and os.path.exists(self.last_csv_path):
            subprocess.Popen(f'explorer /select,"{os.path.normpath(self.last_csv_path)}"')
        else:
            subprocess.Popen(f'explorer "{os.path.normpath(folder)}"')

    # ── Ver Acumulado ─────────────────────────────────────────────
    def _ver_acumulado(self):
        win = tk.Toplevel(self.root)
        win.title("Acumulado")
        win.configure(bg=BG)
        win.geometry("1100x560")

        notebook = ttk.Notebook(win)
        notebook.pack(fill="both", expand=True, padx=10, pady=10)

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TNotebook",        background=BG,  borderwidth=0)
        style.configure("TNotebook.Tab",    background=BG3, foreground=FG,
                        padding=[12, 5])
        style.map("TNotebook.Tab",          background=[("selected", ACCENT)],
                  foreground=[("selected", BG)])
        style.configure("Treeview",         background=BG2, foreground=FG,
                        fieldbackground=BG2, rowheight=22, font=("Segoe UI", 9))
        style.configure("Treeview.Heading", background=BG3, foreground=ACCENT,
                        font=("Segoe UI", 9, "bold"))
        style.map("Treeview",               background=[("selected", ACCENT)],
                  foreground=[("selected", BG)])

        for tipo in ("Porta", "Baf"):
            frames = self.acum.get(tipo, [])
            if frames:
                df_all = pd.concat(frames, ignore_index=True)
            else:
                df_all = pd.DataFrame()

            tab = tk.Frame(notebook, bg=BG)
            notebook.add(tab, text=f"  {tipo}  ({len(df_all)} filas)  ")

            # Encabezado con totales por periodo
            if not df_all.empty and "_periodo" in df_all.columns:
                resumen = df_all.groupby("_periodo").size().reset_index(name="filas")
                hdr_txt = "  |  ".join(
                    f"{row['_periodo']}: {row['filas']}" for _, row in resumen.iterrows()
                )
            else:
                hdr_txt = "Sin datos"

            tk.Label(tab, text=hdr_txt, font=("Segoe UI", 9), bg=BG, fg=YELLOW,
                     anchor="w").pack(fill="x", padx=8, pady=(6, 2))

            # Tabla
            cols_disp = [c for c in COLS_VISTA if c in (df_all.columns if not df_all.empty else [])]
            if not cols_disp and not df_all.empty:
                cols_disp = list(df_all.columns[:10])

            tree_frame = tk.Frame(tab, bg=BG)
            tree_frame.pack(fill="both", expand=True, padx=8, pady=4)

            tree = ttk.Treeview(tree_frame, columns=cols_disp, show="headings",
                                selectmode="browse")
            vsb = ttk.Scrollbar(tree_frame, orient="vertical",   command=tree.yview)
            hsb = ttk.Scrollbar(tree_frame, orient="horizontal",  command=tree.xview)
            tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
            vsb.pack(side="right",  fill="y")
            hsb.pack(side="bottom", fill="x")
            tree.pack(fill="both", expand=True)

            for col in cols_disp:
                tree.heading(col, text=col)
                tree.column(col, width=130, anchor="w", stretch=True)

            if not df_all.empty:
                for _, row in df_all[cols_disp].iterrows():
                    tree.insert("", "end", values=[str(v) if pd.notna(v) else "" for v in row])

            # Pie: total + botón descargar
            pie = tk.Frame(tab, bg=BG)
            pie.pack(fill="x", padx=8, pady=(2, 6))

            tk.Label(pie, text=f"Total acumulado {tipo}: {len(df_all)} filas",
                     font=("Segoe UI", 9, "bold"), bg=BG, fg=GREEN).pack(side="left")

            if not df_all.empty:
                def _descargar(t=tipo, d=df_all):
                    dest = filedialog.asksaveasfilename(
                        title=f"Guardar acumulado {t}",
                        defaultextension=".csv",
                        filetypes=[("CSV", "*.csv")],
                        initialfile=f"acumulado_{t}_{datetime.now().strftime('%d%m_%H%M')}.csv",
                        initialdir=os.path.expanduser("~/Desktop/Robot Leads")
                    )
                    if dest:
                        d.to_csv(dest, index=False, sep=",", encoding="utf-8-sig")
                        messagebox.showinfo("Descargado", f"Acumulado {t} guardado en:\n{dest}")

                tk.Button(pie, text="  Descargar  ",
                          command=_descargar,
                          bg=ACCENT, fg=WHITE, font=("Segoe UI", 9, "bold"),
                          relief="flat", padx=10, pady=4, cursor="hand2"
                          ).pack(side="right")


    # ── Muestreo diario estilo Power BI ──────────────────────────
    def _ver_muestreo(self):
        if not self.lotes:
            messagebox.showinfo("Muestreo", "No hay datos procesados aun.")
            return

        hoy      = datetime.now().strftime("%d/%m/%Y")
        lotes    = self.lotes
        porta_si = sum(l["si"] for l in lotes if l["tipo"] == "Porta")
        baf_si   = sum(l["si"] for l in lotes if l["tipo"] == "Baf")
        total_si = porta_si + baf_si

        C_PORTA = "#89b4fa"
        C_BAF   = "#a6e3a1"
        C_TEXT  = "#cdd6f4"
        C_GRID  = "#313244"

        # ── Ventana ───────────────────────────────────────────────
        win = tk.Toplevel(self.root)
        win.title(f"Muestreo del dia — {hoy}")
        win.configure(bg=BG)
        win.geometry("960x580")

        tk.Label(win, text=f"Procesamiento del dia  {hoy}",
                 font=("Segoe UI", 13, "bold"), bg=BG, fg=FG).pack(pady=(12, 4))

        # ── Tarjetas ──────────────────────────────────────────────
        cards_row = tk.Frame(win, bg=BG, padx=20)
        cards_row.pack(fill="x", pady=(0, 8))

        for label, valor, color in [
            ("TOTAL PROCESADO", total_si, FG),
            ("PORTA",           porta_si, C_PORTA),
            ("BAF",             baf_si,   C_BAF),
            ("LOTES",           len(lotes), YELLOW),
        ]:
            c = tk.Frame(cards_row, bg=BG3, padx=22, pady=10)
            c.pack(side="left", padx=(0, 10))
            tk.Label(c, text=label, font=("Segoe UI", 8), bg=BG3, fg=GRAY).pack()
            tk.Label(c, text=str(valor), font=("Segoe UI", 20, "bold"),
                     bg=BG3, fg=color).pack()

        # ── Gráfico de barras por lote ────────────────────────────
        fig = plt.Figure(figsize=(9.4, 2.8), facecolor=BG)
        ax  = fig.add_subplot(111)
        ax.set_facecolor(BG2)
        ax.tick_params(colors=C_TEXT, labelsize=8)
        for spine in ax.spines.values():
            spine.set_edgecolor(C_GRID)
        ax.yaxis.grid(True, color=C_GRID, linewidth=0.6)
        ax.set_axisbelow(True)

        etiquetas = [f"{l['tipo']}\n{l['hora']}" for l in lotes]
        valores   = [l["si"] for l in lotes]
        colores   = [C_PORTA if l["tipo"] == "Porta" else C_BAF for l in lotes]

        bars = ax.bar(etiquetas, valores, color=colores, width=0.5, zorder=3)
        for bar, v in zip(bars, valores):
            ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 0.3,
                    str(v), ha="center", va="bottom", fontsize=9, color=C_TEXT)

        ax.set_title("Registros 'si' por lote procesado", color=C_TEXT, fontsize=10, pad=6)
        ax.yaxis.set_major_locator(mticker.MaxNLocator(integer=True))
        fig.tight_layout(pad=1.2)

        canvas = FigureCanvasTkAgg(fig, master=win)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="x", padx=14, pady=(0, 6))

        # ── Tabla de lotes ────────────────────────────────────────
        tabla_frame = tk.Frame(win, bg=BG, padx=14)
        tabla_frame.pack(fill="x", pady=(0, 10))

        style = ttk.Style()
        style.configure("Dia.Treeview",
                        background=BG2, foreground=FG,
                        fieldbackground=BG2, rowheight=22,
                        font=("Segoe UI", 9))
        style.configure("Dia.Treeview.Heading",
                        background=BG3, foreground=ACCENT,
                        font=("Segoe UI", 9, "bold"))
        style.map("Dia.Treeview",
                  background=[("selected", ACCENT)],
                  foreground=[("selected", BG)])

        cols = ["Hora", "Archivo", "Tipo", "si", "dupl", "Invalid", "Sin Datos"]
        tree = ttk.Treeview(tabla_frame, columns=cols, show="headings",
                            height=min(len(lotes) + 1, 5),
                            style="Dia.Treeview")

        anchos = {"Hora": 60, "Archivo": 220, "Tipo": 70,
                  "si": 70, "dupl": 70, "Invalid": 70, "Sin Datos": 80}
        for col in cols:
            tree.heading(col, text=col)
            tree.column(col, width=anchos[col], anchor="center")

        for l in lotes:
            color_tag = "porta" if l["tipo"] == "Porta" else "baf"
            tree.insert("", "end",
                        values=(l["hora"], l["archivo"], l["tipo"],
                                l["si"], l["dupl"], l["invalid"], l["sin_datos"]),
                        tags=(color_tag,))

        # Fila de totales
        tree.insert("", "end",
                    values=("", "TOTAL", "",
                            sum(l["si"]        for l in lotes),
                            sum(l["dupl"]      for l in lotes),
                            sum(l["invalid"]   for l in lotes),
                            sum(l["sin_datos"] for l in lotes)),
                    tags=("total",))

        tree.tag_configure("porta", foreground=C_PORTA)
        tree.tag_configure("baf",   foreground=C_BAF)
        tree.tag_configure("total", foreground=YELLOW,
                           font=("Segoe UI", 9, "bold"))
        tree.pack(fill="x")


if __name__ == "__main__":
    root = tk.Tk()
    RobotLeadApp(root)
    root.mainloop()
