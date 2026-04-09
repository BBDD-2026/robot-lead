import streamlit as st
import pandas as pd
import re
import io
import json
import os
from datetime import datetime

# ── Rutas de persistencia ─────────────────────────────────────
DATA_DIR   = os.path.join(os.path.dirname(__file__), "data")
LOTES_FILE = os.path.join(DATA_DIR, "lotes.json")
ACUM_PORTA = os.path.join(DATA_DIR, "acum_Porta.csv")
ACUM_BAF   = os.path.join(DATA_DIR, "acum_Baf.csv")
os.makedirs(DATA_DIR, exist_ok=True)


def _load_lotes() -> list:
    if os.path.exists(LOTES_FILE):
        with open(LOTES_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return []


def _save_lotes(lotes: list):
    with open(LOTES_FILE, "w", encoding="utf-8") as f:
        json.dump(lotes, f, ensure_ascii=False, indent=2)


def _load_acum(path: str) -> pd.DataFrame | None:
    if os.path.exists(path):
        try:
            return pd.read_csv(path, encoding="utf-8-sig", low_memory=False)
        except Exception:
            return None
    return None


def _save_acum(df_new: pd.DataFrame, path: str):
    if os.path.exists(path):
        df_old = pd.read_csv(path, encoding="utf-8-sig", low_memory=False)
        df_all = pd.concat([df_old, df_new], ignore_index=True)
    else:
        df_all = df_new
    df_all.to_csv(path, index=False, encoding="utf-8-sig")


# ── Config ────────────────────────────────────────────────────
st.set_page_config(
    page_title="Robot Lead",
    page_icon="🤖",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ── Estilos dark mode ─────────────────────────────────────────
st.markdown("""
<style>
    .stApp { background-color: #1e1e2e; color: #cdd6f4; }
    .block-container { padding-top: 1.5rem; }
    .metric-card {
        background: #313244; border-radius: 10px;
        padding: 16px 24px; text-align: center;
    }
    .metric-card .label { font-size: 11px; color: #6c7086; margin-bottom: 4px; }
    .metric-card .value { font-size: 28px; font-weight: 700; }
    .section-title {
        font-size: 13px; color: #6c7086;
        margin-bottom: 6px; margin-top: 12px;
    }
    div[data-testid="stDataFrame"] { border-radius: 8px; }
    .stDownloadButton > button {
        background-color: #89b4fa !important;
        color: #1e1e2e !important; font-weight: 700 !important;
        border: none !important; border-radius: 6px !important;
    }
    .stButton > button {
        background-color: #313244 !important;
        color: #cdd6f4 !important; border: none !important;
        border-radius: 6px !important;
    }
    hr { border-color: #313244; }
</style>
""", unsafe_allow_html=True)

MESES = {
    "01": "Enero",   "02": "Febrero", "03": "Marzo",    "04": "Abril",
    "05": "Mayo",    "06": "Junio",   "07": "Julio",    "08": "Agosto",
    "09": "Septiembre", "10": "Octubre", "11": "Noviembre", "12": "Diciembre"
}

# ── Session state ─────────────────────────────────────────────
if "acum" not in st.session_state:
    acum_porta = _load_acum(ACUM_PORTA)
    acum_baf   = _load_acum(ACUM_BAF)
    st.session_state.acum = {
        "Porta": [acum_porta] if acum_porta is not None else [],
        "Baf":   [acum_baf]   if acum_baf   is not None else [],
    }
if "lotes" not in st.session_state:
    st.session_state.lotes = _load_lotes()


# ── Helpers ───────────────────────────────────────────────────
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


def decode_db_id(value) -> tuple:
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


def procesar(df: pd.DataFrame) -> dict:
    df = df.copy()

    # Ruta → Province
    if "Ruta" in df.columns and "Province" in df.columns:
        df["Province"] = df["Ruta"]

    # Localidad → City
    if "Localidad" in df.columns and "City" in df.columns:
        df["City"] = df["Localidad"]

    # Limpiar customer_firstname
    sin_datos_count = 0
    if "customer_firstname" in df.columns:
        df["customer_firstname"] = df["customer_firstname"].apply(
            lambda x: x.strip() if is_valid_name(x) else "Sin Datos"
        )
        sin_datos_count = int((df["customer_firstname"] == "Sin Datos").sum())

    # Ordenar A→Z
    if "customer_firstname" in df.columns:
        df = df.sort_values("customer_firstname", ascending=True, ignore_index=True)

    # Columna Subir
    df["Subir"] = ""

    dupl_count = 0
    if "Gen_Insert" in df.columns:
        mask_dupl = df["Gen_Insert"].notna()
        df.loc[mask_dupl, "Subir"] = "dupl"
        dupl_count = int(mask_dupl.sum())

    invalid_count = 0
    if "Ruta" in df.columns:
        mask_inv = (
            df["Ruta"].isna() |
            (df["Ruta"].astype(str).str.strip().str.upper() == "NULL")
        ) & (df["Subir"] == "")
        df.loc[mask_inv, "Subir"] = "Invalid"
        invalid_count = int(mask_inv.sum())

    mask_si = df["Subir"] == ""
    df.loc[mask_si, "Subir"] = "si"
    si_count = int(mask_si.sum())

    return {
        "df": df,
        "si": si_count,
        "dupl": dupl_count,
        "invalid": invalid_count,
        "sin_datos": sin_datos_count,
    }


def df_to_excel_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    df.to_excel(buf, index=False)
    return buf.getvalue()


def df_to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False, sep=",").encode("latin-1", errors="replace")


def build_csv_si(df: pd.DataFrame) -> pd.DataFrame:
    df_si = df[df["Subir"] == "si"].copy().reset_index(drop=True)
    df_si["record_id"] = range(1, len(df_si) + 1)
    df_si["chain_id"]  = df_si["record_id"]
    if "record_id" in df_si.columns:
        start = df_si.columns.get_loc("record_id")
        df_si = df_si.iloc[:, start:]
    if "DB_ID" in df_si.columns:
        end = df_si.columns.get_loc("DB_ID") + 1
        df_si = df_si.iloc[:, :end]
    return df_si


def metric_card(label, value, color="#cdd6f4"):
    return f"""
    <div class="metric-card">
        <div class="label">{label}</div>
        <div class="value" style="color:{color}">{value}</div>
    </div>"""


def build_period_table(frames_porta: list, frames_baf: list) -> pd.DataFrame | None:
    """Construye tabla pivoteada: Período | Porta | Baf | Total."""
    rows = []
    for df_chunk in frames_porta:
        if "_periodo" in df_chunk.columns:
            tmp = df_chunk[["_periodo"]].copy()
            tmp["_tipo"] = "Porta"
            rows.append(tmp)
    for df_chunk in frames_baf:
        if "_periodo" in df_chunk.columns:
            tmp = df_chunk[["_periodo"]].copy()
            tmp["_tipo"] = "Baf"
            rows.append(tmp)

    if not rows:
        return None

    df_all = pd.concat(rows, ignore_index=True)
    grp = df_all.groupby(["_periodo", "_tipo"]).size().reset_index(name="n")
    pivot = grp.pivot_table(index="_periodo", columns="_tipo", values="n", aggfunc="sum", fill_value=0)
    pivot.columns.name = None
    pivot = pivot.reset_index().rename(columns={"_periodo": "Período"})
    for col in ("Porta", "Baf"):
        if col not in pivot.columns:
            pivot[col] = 0
    pivot["Total"] = pivot["Porta"] + pivot["Baf"]
    return pivot[["Período", "Porta", "Baf", "Total"]]


# ── Header ────────────────────────────────────────────────────
st.markdown("<h1 style='text-align:center;color:#89b4fa;margin-bottom:0'>ROBOT LEAD</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align:center;color:#6c7086;margin-top:0'>Procesador de Archivos Excel · Porta / Baf</p>", unsafe_allow_html=True)
st.markdown("---")

# ── Tabs principales ──────────────────────────────────────────
tab_proc, tab_muestreo, tab_acum = st.tabs(["⚙  Procesar", "📊  Muestreo del Dia", "📋  Acumulado"])


# ══════════════════════════════════════════════════════════════
# TAB 1 — PROCESAR
# ══════════════════════════════════════════════════════════════
with tab_proc:
    col_porta, col_sep, col_baf = st.columns([10, 1, 10])

    def _render_panel(col, tipo: str, up_key: str, color: str):
        with col:
            st.markdown(f"<p style='color:{color};font-weight:700;font-size:16px;margin-bottom:4px'>{tipo}</p>",
                        unsafe_allow_html=True)
            uploaded = st.file_uploader(f"Archivo {tipo}", type=["xlsx", "xls"],
                                        key=up_key, label_visibility="collapsed")
            if not uploaded:
                return

            df_raw = pd.read_excel(uploaded)
            st.markdown(
                metric_card("FILAS CARGADAS", len(df_raw), "#cdd6f4"),
                unsafe_allow_html=True
            )
            st.markdown("<br>", unsafe_allow_html=True)

            res_key = f"res_{tipo}"
            nom_key = f"nom_{tipo}"

            if st.button(f"⚙  Procesar {tipo}", key=f"btn_{tipo}", use_container_width=True):
                with st.spinner("Procesando..."):
                    st.session_state[res_key] = procesar(df_raw)
                    st.session_state[nom_key] = uploaded.name

            if res_key in st.session_state and st.session_state.get(nom_key) == uploaded.name:
                res  = st.session_state[res_key]
                df_p = res["df"]

                st.markdown("<div class='section-title'>Resultado</div>", unsafe_allow_html=True)
                r1, r2 = st.columns(2)
                r1.markdown(metric_card("si",        res["si"],        "#a6e3a1"), unsafe_allow_html=True)
                r2.markdown(metric_card("dupl",      res["dupl"],      "#f9e2af"), unsafe_allow_html=True)
                r3, r4 = st.columns(2)
                r3.markdown(metric_card("Invalid",   res["invalid"],   "#f38ba8"), unsafe_allow_html=True)
                r4.markdown(metric_card("Sin Datos", res["sin_datos"], "#f38ba8" if res["sin_datos"] else "#a6e3a1"), unsafe_allow_html=True)

                st.markdown("<br>", unsafe_allow_html=True)

                ts   = datetime.now().strftime("%d%m_%H%M")
                base = uploaded.name.rsplit(".", 1)[0]
                df_csv = build_csv_si(df_p)

                st.download_button(
                    label=f"📥 Excel procesado ({len(df_p)} filas)",
                    data=df_to_excel_bytes(df_p),
                    file_name=f"{base}_procesado_{ts}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key=f"dl_excel_{tipo}"
                )

                import re as _re
                fecha_match = _re.search(r'_(\d{4})', uploaded.name)
                fecha_code  = fecha_match.group(1) if fecha_match else datetime.now().strftime("%d%m")
                csv_name    = f"Lote_Leads_{tipo}_{fecha_code}.csv"

                st.download_button(
                    label=f"📥 CSV subir ({len(df_csv)} filas 'si')",
                    data=df_to_csv_bytes(df_csv),
                    file_name=csv_name,
                    mime="text/csv",
                    use_container_width=True,
                    key=f"dl_csv_{tipo}"
                )

                # Acumular una sola vez por archivo
                key_guard = f"guardado_{uploaded.name}"
                if key_guard not in st.session_state:
                    st.session_state[key_guard] = True
                    df_acum = df_p[df_p["Subir"] == "si"].copy()
                    df_acum["_fecha_proceso"] = datetime.now().strftime("%d/%m/%Y")
                    if "DB_ID" in df_acum.columns:
                        decoded = df_acum["DB_ID"].apply(
                            lambda x: decode_db_id(str(x)) if pd.notna(x) else ("?", "?")
                        )
                        df_acum["_tipo"]    = decoded.apply(lambda x: x[0])
                        df_acum["_periodo"] = decoded.apply(lambda x: x[1])
                    else:
                        df_acum["_tipo"]    = tipo
                        df_acum["_periodo"] = "?"

                    st.session_state.acum[tipo].append(df_acum)

                    lote = {
                        "tipo":      tipo,
                        "archivo":   uploaded.name,
                        "fecha":     datetime.now().strftime("%d/%m/%Y"),
                        "hora":      datetime.now().strftime("%H:%M"),
                        "si":        res["si"],
                        "dupl":      res["dupl"],
                        "invalid":   res["invalid"],
                        "sin_datos": res["sin_datos"],
                    }
                    st.session_state.lotes.append(lote)

                    _save_lotes(st.session_state.lotes)
                    acum_path = ACUM_PORTA if tipo == "Porta" else ACUM_BAF
                    _save_acum(df_acum, acum_path)

    _render_panel(col_porta, "Porta", "up_porta", "#89b4fa")
    with col_sep:
        st.markdown("<div style='border-left:1px solid #313244;height:100%;margin:0 auto'></div>",
                    unsafe_allow_html=True)
    _render_panel(col_baf, "Baf", "up_baf", "#a6e3a1")


# ══════════════════════════════════════════════════════════════
# TAB 2 — MUESTREO DEL DIA
# ══════════════════════════════════════════════════════════════
with tab_muestreo:
    hoy   = datetime.now().strftime("%d/%m/%Y")
    lotes = st.session_state.lotes
    lotes_hoy = [l for l in lotes if l.get("fecha") == hoy]

    st.markdown(f"<h4 style='color:#cdd6f4'>Procesamiento del dia — {hoy}</h4>", unsafe_allow_html=True)

    if not lotes_hoy:
        st.info("No se procesaron archivos hoy.")
    else:
        porta_si = sum(l["si"] for l in lotes_hoy if l["tipo"] == "Porta")
        baf_si   = sum(l["si"] for l in lotes_hoy if l["tipo"] == "Baf")
        total_si = porta_si + baf_si

        mc1, mc2, mc3 = st.columns(3)
        mc1.markdown(metric_card("TOTAL", total_si, "#cdd6f4"), unsafe_allow_html=True)
        mc2.markdown(metric_card("PORTA", porta_si, "#89b4fa"), unsafe_allow_html=True)
        mc3.markdown(metric_card("BAF",   baf_si,   "#a6e3a1"), unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # Período por DB_ID — solo registros de hoy
        frames_porta_hoy, frames_baf_hoy = [], []
        for tipo, frames_list in (("Porta", frames_porta_hoy), ("Baf", frames_baf_hoy)):
            for df_chunk in st.session_state.acum.get(tipo, []):
                if "_periodo" in df_chunk.columns and "_fecha_proceso" in df_chunk.columns:
                    df_hoy = df_chunk[df_chunk["_fecha_proceso"] == hoy]
                    if not df_hoy.empty:
                        frames_list.append(df_hoy)

        tabla_hoy = build_period_table(frames_porta_hoy, frames_baf_hoy)

        if tabla_hoy is not None:
            st.markdown("<div class='section-title'>Registros 'si' por período (DB_ID)</div>", unsafe_allow_html=True)
            st.dataframe(tabla_hoy, use_container_width=False, hide_index=True)
        else:
            st.caption("Sin desglose por período disponible para hoy.")

        st.markdown("---")
        st.markdown("<div class='section-title'>Archivos procesados hoy</div>", unsafe_allow_html=True)
        df_lotes_hoy = pd.DataFrame(lotes_hoy)[["tipo", "archivo", "hora", "si", "dupl", "invalid", "sin_datos"]]
        df_lotes_hoy.columns = ["Tipo", "Archivo", "Hora", "si", "dupl", "Invalid", "Sin Datos"]
        st.dataframe(df_lotes_hoy, use_container_width=True, hide_index=True)


# ══════════════════════════════════════════════════════════════
# TAB 3 — ACUMULADO
# ══════════════════════════════════════════════════════════════
with tab_acum:
    col_rst, _ = st.columns([1, 4])
    with col_rst:
        if st.button("🗑  Dejar en cero", use_container_width=True):
            for path in [LOTES_FILE, ACUM_PORTA, ACUM_BAF]:
                if os.path.exists(path):
                    os.remove(path)
            st.session_state.acum  = {"Porta": [], "Baf": []}
            st.session_state.lotes = []
            for k in list(st.session_state.keys()):
                if k.startswith("guardado_"):
                    del st.session_state[k]
            st.success("Acumulado reiniciado correctamente.")
            st.rerun()

    st.markdown("---")

    frames_porta = st.session_state.acum.get("Porta", [])
    frames_baf   = st.session_state.acum.get("Baf",   [])
    tabla_acum   = build_period_table(frames_porta, frames_baf)

    if tabla_acum is not None:
        st.markdown("<div class='section-title'>Total acumulado por período (DB_ID)</div>", unsafe_allow_html=True)
        st.dataframe(tabla_acum, use_container_width=False, hide_index=True)
        st.markdown("<br>", unsafe_allow_html=True)
    else:
        st.info("Sin datos acumulados.")

    # Descargas por tipo
    dl1, dl2 = st.columns(2)
    for col_dl, tipo in ((dl1, "Porta"), (dl2, "Baf")):
        lst = st.session_state.acum.get(tipo, [])
        if lst:
            df_t = pd.concat(lst, ignore_index=True)
            ts_dl = datetime.now().strftime("%d%m_%H%M")
            col_dl.download_button(
                label=f"📥 Descargar acumulado {tipo}",
                data=df_to_csv_bytes(df_t),
                file_name=f"acumulado_{tipo}_{ts_dl}.csv",
                mime="text/csv",
                use_container_width=True,
                key=f"dl_acum_{tipo}"
            )
