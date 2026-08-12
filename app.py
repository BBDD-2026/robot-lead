import streamlit as st
import pandas as pd
import re
import io
import os
import random
import hashlib
import sqlite3
from datetime import datetime

from reference_data import REGION_MAP, AGENT_IDS

# ── Rutas de persistencia ─────────────────────────────────────
DATA_DIR = os.path.join(os.path.dirname(__file__), "data")
DB_FILE  = os.path.join(DATA_DIR, "robot_lead.db")
os.makedirs(DATA_DIR, exist_ok=True)


# ── SQLite ────────────────────────────────────────────────────
def _db_con():
    return sqlite3.connect(DB_FILE)

def _db_init():
    con = _db_con()
    con.execute("""
        CREATE TABLE IF NOT EXISTS lotes (
            id           INTEGER PRIMARY KEY AUTOINCREMENT,
            archivo_hash TEXT    UNIQUE,
            tipo         TEXT,
            archivo      TEXT,
            fecha        TEXT,
            hora         TEXT,
            mes          INTEGER,
            anio         INTEGER,
            si           INTEGER,
            dupl         INTEGER,
            invalid      INTEGER,
            sin_datos    INTEGER
        )
    """)
    con.commit()
    con.close()

_db_init()

def _hash_bytes(b: bytes) -> str:
    return hashlib.md5(b).hexdigest()

def _db_hash_exists(h: str) -> bool:
    con = _db_con()
    row = con.execute("SELECT 1 FROM lotes WHERE archivo_hash=?", (h,)).fetchone()
    con.close()
    return row is not None

def _db_insert_lote(d: dict):
    con = _db_con()
    con.execute("""
        INSERT OR IGNORE INTO lotes
        (archivo_hash,tipo,archivo,fecha,hora,mes,anio,si,dupl,invalid,sin_datos)
        VALUES (:archivo_hash,:tipo,:archivo,:fecha,:hora,:mes,:anio,:si,:dupl,:invalid,:sin_datos)
    """, d)
    con.commit()
    con.close()

def _db_insert_acum(df: pd.DataFrame):
    con = _db_con()
    df.to_sql("acumulado", con, if_exists="append", index=False)
    con.close()

def _db_load_lotes(fecha: str = None) -> list:
    con = _db_con()
    try:
        if fecha:
            df = pd.read_sql("SELECT * FROM lotes WHERE fecha=?", con, params=(fecha,))
        else:
            df = pd.read_sql("SELECT * FROM lotes", con)
        return df.to_dict("records")
    except Exception:
        return []
    finally:
        con.close()

def _db_load_acum(tipo: str = None, mes: int = None, anio: int = None, fecha: str = None) -> pd.DataFrame | None:
    con = _db_con()
    try:
        conditions, params = ["1=1"], []
        if tipo:  conditions.append("_tipo=?");           params.append(tipo)
        if mes:   conditions.append("_mes=?");            params.append(mes)
        if anio:  conditions.append("_anio=?");           params.append(anio)
        if fecha: conditions.append("_fecha_proceso=?");  params.append(fecha)
        df = pd.read_sql(
            f"SELECT * FROM acumulado WHERE {' AND '.join(conditions)}",
            con, params=params if params else None
        )
        return df if not df.empty else None
    except Exception:
        return None
    finally:
        con.close()

def _db_available_periods() -> list[tuple]:
    """Devuelve lista de (mes, anio) disponibles en lotes, ordenados desc."""
    con = _db_con()
    try:
        df = pd.read_sql("SELECT DISTINCT mes, anio FROM lotes ORDER BY anio DESC, mes DESC", con)
        return [(int(r.mes), int(r.anio)) for _, r in df.iterrows()]
    except Exception:
        return []
    finally:
        con.close()

def _db_reset():
    con = _db_con()
    con.execute("DROP TABLE IF EXISTS lotes")
    con.execute("DROP TABLE IF EXISTS acumulado")
    con.commit()
    con.close()
    _db_init()


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
    1: "Enero",    2: "Febrero",   3: "Marzo",    4: "Abril",
    5: "Mayo",     6: "Junio",     7: "Julio",    8: "Agosto",
    9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
}
MESES_STR = {f"{k:02d}": v for k, v in MESES.items()}


# ── Helpers ───────────────────────────────────────────────────
def is_valid_name(value) -> bool:
    if not value or not isinstance(value, str):
        return False
    v = value.strip()
    if not v or len(v) > 30 or "@" in v:
        return False
    if not re.fullmatch(r'[a-zA-Zà-ÿ ]+', v):
        return False
    if any(len(p) < 2 for p in v.split()):
        return False
    return True


def decode_db_id(value) -> tuple:
    if not value or not isinstance(value, str):
        return ("?", "?")
    v = value.strip()
    if len(v) < 6 or v[0].upper() != "L":
        return ("?", v)
    mm   = v[1:3]
    yy   = v[3:5]
    tipo = v[5].upper()
    mes  = MESES_STR.get(mm, mm)
    year = f"20{yy}"
    nombre = "Porta" if tipo == "P" else "Baf" if tipo == "B" else tipo
    return (nombre, f"{mes} {year}")


# ── Estructura de salida CRM (52 columnas) ──────────────────────
TARGET_COLS = [
    "RECORD_ID", "CONTACT_INFO", "CONTACT_INFO_TYPE", "RECORD_TYPE", "RECORD_STATUS",
    "CALL_RESULT", "ATTEMPT", "DIAL_SCHED_TIME", "CALL_TIME", "DAILY_FROM", "DAILY_TILL",
    "TZ_DBID", "CAMPAIGN_ID", "AGENT_ID", "CHAIN_ID", "CHAIN_N", "GROUP_ID", "APP_ID",
    "TREATMENTS", "MEDIA_REF", "EMAIL_SUBJECT", "EMAIL_TEMPLATE_ID", "SWITCH_ID",
    "CORTEBASE", "DISPOSITIONCODE", "CLIENTE_NOMBRE", "CLIENTE_APELLIDO", "REGION",
    "UBICACION", "SITIOS", "EMAIL", "CAMPANIA", "FECHA_INSERCION",
    "CUST_DATA_1", "CUST_DATA_2", "CUST_DATA_3", "CUST_DATA_4", "CUST_DATA_5",
    "CUST_DATA_6", "CUST_DATA_7", "CUST_DATA_8", "CUST_DATA_9", "CUST_DATA_10",
    "CUST_DATA_11", "CUST_DATA_12", "CUST_DATA_13", "CUST_DATA_14", "CUST_DATA_15",
    "CUST_DATA_ED_1", "CUST_DATA_ED_2", "CUST_DATA_ED_3", "CUST_DATA_ED_4",
]

# Columnas que se copian tal cual desde el archivo fuente (case-insensitive)
SIMPLE_MAP = {
    "CONTACT_INFO": "contact_info", "CONTACT_INFO_TYPE": "contact_info_type",
    "RECORD_TYPE": "record_type", "RECORD_STATUS": "record_status",
    "CALL_RESULT": "call_result", "ATTEMPT": "attempt",
    "DIAL_SCHED_TIME": "dial_sched_time", "CALL_TIME": "call_time",
    "DAILY_FROM": "daily_from", "DAILY_TILL": "daily_till", "TZ_DBID": "tz_dbid",
    "CAMPAIGN_ID": "campaign_id", "CHAIN_N": "chain_n", "GROUP_ID": "group_id",
    "APP_ID": "app_id", "TREATMENTS": "treatments", "MEDIA_REF": "media_ref",
    "EMAIL_SUBJECT": "email_subject", "EMAIL_TEMPLATE_ID": "email_template_id",
    "SWITCH_ID": "switch_id", "CORTEBASE": "DB_ID", "DISPOSITIONCODE": "DispositionCode",
    "CLIENTE_NOMBRE": "customer_firstname", "CLIENTE_APELLIDO": "customer_lastname",
    "FECHA_INSERCION": "Gen_Insert", "CUST_DATA_13": "idinterno_gestor",
    "CUST_DATA_14": "Localidad", "CUST_DATA_15": "Ruta",
}

# Columnas con valor literal fijo
CONSTANT_MAP = {"CAMPANIA": "PortOut", "CUST_DATA_12": "MOVISTAR"}

# Columnas sin dato de origen: siempre vacías
EMPTY_COLS = (
    ["UBICACION", "SITIOS", "EMAIL"]
    + [f"CUST_DATA_{i}" for i in range(1, 12)]
    + [f"CUST_DATA_ED_{i}" for i in range(1, 5)]
)


def build_target_df(df_si: pd.DataFrame) -> pd.DataFrame:
    """Arma la matriz de 52 columnas (formato CRM) a partir de las filas 'si'."""
    n = len(df_si)
    col_lower = {c.lower(): c for c in df_si.columns}

    def get(source_col: str) -> pd.Series:
        real = col_lower.get(source_col.lower())
        return df_si[real].reset_index(drop=True) if real else pd.Series([None] * n)

    out = pd.DataFrame(index=range(n))
    for target, source in SIMPLE_MAP.items():
        out[target] = get(source)
    for target, const in CONSTANT_MAP.items():
        out[target] = const
    for target in EMPTY_COLS:
        out[target] = ""

    out["RECORD_ID"] = range(1, n + 1)
    out["CHAIN_ID"] = out["RECORD_ID"]

    ruta_key = get("Ruta").astype(str).str.strip().str.upper()
    out["REGION"] = ruta_key.map(REGION_MAP).fillna("")

    out["AGENT_ID"] = (
        [random.choice(AGENT_IDS) for _ in range(n)] if AGENT_IDS else ""
    )

    return out[TARGET_COLS]


def procesar(df: pd.DataFrame) -> dict:
    df = df.copy()
    if "Ruta" in df.columns and "Province" in df.columns:
        df["Province"] = df["Ruta"]
    if "Localidad" in df.columns and "City" in df.columns:
        df["City"] = df["Localidad"]

    sin_datos_count = 0
    if "customer_firstname" in df.columns:
        df["customer_firstname"] = df["customer_firstname"].apply(
            lambda x: x.strip() if is_valid_name(x) else "Sin Datos"
        )
        sin_datos_count = int((df["customer_firstname"] == "Sin Datos").sum())
    if "customer_firstname" in df.columns:
        df = df.sort_values("customer_firstname", ascending=True, ignore_index=True)

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

    return {"df": df, "si": si_count, "dupl": dupl_count, "invalid": invalid_count, "sin_datos": sin_datos_count}


def df_to_excel_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    df.to_excel(buf, index=False)
    return buf.getvalue()


def df_to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False, sep=",").encode("latin-1", errors="replace")


def build_csv_si(df: pd.DataFrame) -> pd.DataFrame:
    df_si = df[df["Subir"] == "si"].copy().reset_index(drop=True)
    return build_target_df(df_si)


def build_download_acum(df: pd.DataFrame) -> pd.DataFrame:
    """Filtra filas 'si' y arma la matriz de 52 columnas (formato CRM)."""
    if "Subir" in df.columns:
        df = df[df["Subir"] == "si"].copy().reset_index(drop=True)
    else:
        df = df.copy().reset_index(drop=True)
    return build_target_df(df)


def metric_card(label, value, color="#cdd6f4"):
    return f"""
    <div class="metric-card">
        <div class="label">{label}</div>
        <div class="value" style="color:{color}">{value}</div>
    </div>"""


def build_period_table(df: pd.DataFrame) -> pd.DataFrame | None:
    """Tabla pivote Período | Porta | Baf | Total (solo filas 'si')."""
    if df is None or "_periodo" not in df.columns or "_tipo" not in df.columns:
        return None
    if "Subir" in df.columns:
        df = df[df["Subir"] == "si"]
    grp   = df.groupby(["_periodo", "_tipo"]).size().reset_index(name="n")
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

tab_proc, tab_muestreo, tab_acum, tab_extractor = st.tabs(["⚙  Procesar", "📊  Muestreo del Dia", "📋  Acumulado", "🔍  Extractor"])


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
            st.markdown(metric_card("FILAS CARGADAS", len(df_raw), "#cdd6f4"), unsafe_allow_html=True)
            st.markdown("<br>", unsafe_allow_html=True)

            res_key  = f"res_{tipo}"
            nom_key  = f"nom_{tipo}"
            hash_key = f"hash_{tipo}"

            if st.button(f"⚙  Procesar {tipo}", key=f"btn_{tipo}", use_container_width=True):
                with st.spinner("Procesando..."):
                    st.session_state[res_key]  = procesar(df_raw)
                    st.session_state[nom_key]  = uploaded.name
                    st.session_state[hash_key] = _hash_bytes(uploaded.getvalue())

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
                    use_container_width=True, key=f"dl_excel_{tipo}"
                )

                fecha_match = re.search(r'_(\d{4})', uploaded.name)
                fecha_code  = fecha_match.group(1) if fecha_match else datetime.now().strftime("%d%m")
                csv_name    = f"Lote_Fija_{fecha_code}.csv" if tipo == "Baf" else f"Lote_Leads_{tipo}_{fecha_code}.csv"

                st.download_button(
                    label=f"📥 CSV subir ({len(df_csv)} filas 'si')",
                    data=df_to_csv_bytes(df_csv),
                    file_name=csv_name,
                    mime="text/csv",
                    use_container_width=True, key=f"dl_csv_{tipo}"
                )

                # ── Acumular: una vez por hash de archivo (persiste entre sesiones) ──
                file_hash = st.session_state.get(hash_key, _hash_bytes(uploaded.getvalue()))
                guard_key = f"acum_{file_hash}"
                if guard_key not in st.session_state:
                    st.session_state[guard_key] = True
                    if _db_hash_exists(file_hash):
                        st.warning("⚠ Este archivo ya estaba acumulado. No se duplicó.")
                    else:
                        now = datetime.now()
                        df_acum = df_p.copy()  # todas las filas (si, dupl, Invalid)
                        df_acum["_fecha_proceso"] = now.strftime("%d/%m/%Y")
                        df_acum["_mes"]  = now.month
                        df_acum["_anio"] = now.year
                        if "DB_ID" in df_p.columns:
                            decoded = df_p["DB_ID"].apply(
                                lambda x: decode_db_id(str(x)) if pd.notna(x) else ("?", "?")
                            )
                            df_acum["_tipo"]    = decoded.apply(lambda t: t[0] if t[0] != "?" else tipo)
                            df_acum["_periodo"] = decoded.apply(lambda t: t[1])
                        else:
                            df_acum["_tipo"]    = tipo
                            df_acum["_periodo"] = "?"
                        df_acum["_archivo_hash"] = file_hash
                        _db_insert_acum(df_acum)

                        _db_insert_lote({
                            "archivo_hash": file_hash,
                            "tipo":      tipo,
                            "archivo":   uploaded.name,
                            "fecha":     now.strftime("%d/%m/%Y"),
                            "hora":      now.strftime("%H:%M"),
                            "mes":       now.month,
                            "anio":      now.year,
                            "si":        res["si"],
                            "dupl":      res["dupl"],
                            "invalid":   res["invalid"],
                            "sin_datos": res["sin_datos"],
                        })
                        st.success("✅ Datos acumulados en la base.")

    _render_panel(col_porta, "Porta", "up_porta", "#89b4fa")
    with col_sep:
        st.markdown("<div style='border-left:1px solid #313244;height:100%;margin:0 auto'></div>",
                    unsafe_allow_html=True)
    _render_panel(col_baf, "Baf", "up_baf", "#a6e3a1")


# ══════════════════════════════════════════════════════════════
# TAB 2 — MUESTREO DEL DIA
# ══════════════════════════════════════════════════════════════
with tab_muestreo:
    hoy       = datetime.now().strftime("%d/%m/%Y")
    lotes_hoy = _db_load_lotes(fecha=hoy)

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

        df_hoy    = _db_load_acum(fecha=hoy)
        tabla_hoy = build_period_table(df_hoy)
        if tabla_hoy is not None:
            st.markdown("<div class='section-title'>Registros 'si' por período (DB_ID)</div>", unsafe_allow_html=True)
            st.dataframe(tabla_hoy, use_container_width=False, hide_index=True)

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
            _db_reset()
            for k in list(st.session_state.keys()):
                if k.startswith("acum_"):
                    del st.session_state[k]
            st.success("Acumulado reiniciado correctamente.")
            st.rerun()

    st.markdown("---")

    # ── Selector de período ──
    periods        = _db_available_periods()
    period_options = ["Todo"] + [f"{MESES[m]} {a}" for m, a in periods]
    period_map     = {f"{MESES[m]} {a}": (m, a) for m, a in periods}

    col_sel, _ = st.columns([2, 3])
    with col_sel:
        sel_period = st.selectbox("Período", period_options, key="sel_period")

    if sel_period == "Todo":
        df_acum_all = _db_load_acum()
        suffix = ""
    else:
        mes_fil, anio_fil = period_map[sel_period]
        df_acum_all = _db_load_acum(mes=mes_fil, anio=anio_fil)
        suffix = f"_{sel_period.replace(' ', '_')}"

    tabla_acum = build_period_table(df_acum_all)
    if tabla_acum is not None:
        st.markdown("<div class='section-title'>Total por período (DB_ID)</div>", unsafe_allow_html=True)
        st.dataframe(tabla_acum, use_container_width=False, hide_index=True)
        st.markdown("<br>", unsafe_allow_html=True)
    else:
        st.info("Sin datos acumulados.")

    # ── Descargas por tipo ──
    if df_acum_all is not None and "_tipo" in df_acum_all.columns:
        ts_dl = datetime.now().strftime("%d%m_%H%M")
        dl1, dl2 = st.columns(2)
        for col_dl, tipo_dl in ((dl1, "Porta"), (dl2, "Baf")):
            df_tipo = df_acum_all[df_acum_all["_tipo"] == tipo_dl]
            if not df_tipo.empty:
                df_down = build_download_acum(df_tipo)
                col_dl.download_button(
                    label=f"📥 Descargar {tipo_dl}{suffix} ({len(df_down)} filas)",
                    data=df_to_csv_bytes(df_down),
                    file_name=f"acumulado_{tipo_dl}{suffix}_{ts_dl}.csv",
                    mime="text/csv",
                    use_container_width=True,
                    key=f"dl_acum_{tipo_dl}"
                )


# ══════════════════════════════════════════════════════════════
# TAB 4 — EXTRACTOR
# ══════════════════════════════════════════════════════════════
EXTRACT_COLS = ["idinterno_gestor", "tipo_lead", "contact_info", "Ruta", "Localidad", "Provincia", "Gen_Insert", "customer_firstname"]

with tab_extractor:
    st.markdown("<h4 style='color:#cdd6f4'>Extractor de columnas</h4>", unsafe_allow_html=True)
    st.markdown(
        "<p style='color:#6c7086;font-size:13px'>Sube uno o varios archivos Excel. "
        "Se extraen las columnas <code>idinterno_gestor</code>, <code>tipo_lead</code>, "
        "<code>contact_info</code>, <code>Ruta</code>, <code>Localidad</code>, <code>Provincia</code>, "
        "<code>Gen_Insert</code> y <code>customer_firstname</code> de todos los archivos y se combinan en uno solo.</p>",
        unsafe_allow_html=True
    )

    uploaded_ext = st.file_uploader(
        "Archivos Excel o CSV",
        type=["xlsx", "xls", "csv"],
        accept_multiple_files=True,
        key="up_extractor",
        label_visibility="collapsed"
    )

    if uploaded_ext:
        frames, errores, proc_metrics = [], [], []

        for f in uploaded_ext:
            try:
                if f.name.lower().endswith(".csv"):
                    sheets = {"CSV": pd.read_csv(f, encoding="latin-1", sep=None, engine="python")}
                else:
                    sheets = pd.read_excel(f, sheet_name=None)

                for sheet_name, df_raw in sheets.items():
                    col_lower = {c.lower(): c for c in df_raw.columns}
                    subir_col = col_lower.get("subir")
                    if subir_col is not None:
                        df_si = df_raw[df_raw[subir_col].astype(str).str.strip().str.lower() == "si"]
                    else:
                        # Archivo crudo: procesar primero y filtrar a "si"
                        res = procesar(df_raw)
                        proc_metrics.append({
                            "Archivo": f.name,
                            "si": res["si"],
                            "dupl": res["dupl"],
                            "Invalid": res["invalid"],
                            "Sin Datos": res["sin_datos"],
                        })
                        df_si = res["df"][res["df"]["Subir"] == "si"].copy()

                    if df_si.empty:
                        continue
                    col_lower_si = {c.lower(): c for c in df_si.columns}
                    cols_presentes = [col_lower_si[c.lower()] for c in EXTRACT_COLS if c.lower() in col_lower_si]
                    cols_faltantes = [c for c in EXTRACT_COLS if c.lower() not in col_lower_si]
                    if cols_faltantes:
                        label = f"{f.name} [{sheet_name}]" if sheet_name != "CSV" else f.name
                        errores.append(f"**{label}** — columnas no encontradas: {', '.join(cols_faltantes)}")
                    if cols_presentes:
                        frames.append(df_si[cols_presentes].copy())
            except Exception as e:
                errores.append(f"**{f.name}** — error al leer: {e}")

        # Mostrar resultado del procesamiento si hubo archivos crudos
        if proc_metrics:
            st.markdown("<div class='section-title'>Resultado del procesamiento</div>", unsafe_allow_html=True)
            for pm in proc_metrics:
                st.markdown(f"<p style='color:#6c7086;font-size:12px;margin-bottom:4px'>{pm['Archivo']}</p>", unsafe_allow_html=True)
                pm1, pm2, pm3, pm4 = st.columns(4)
                pm1.markdown(metric_card("si",        pm["si"],        "#a6e3a1"), unsafe_allow_html=True)
                pm2.markdown(metric_card("dupl",      pm["dupl"],      "#f9e2af"), unsafe_allow_html=True)
                pm3.markdown(metric_card("Invalid",   pm["Invalid"],   "#f38ba8"), unsafe_allow_html=True)
                pm4.markdown(metric_card("Sin Datos", pm["Sin Datos"], "#f38ba8" if pm["Sin Datos"] else "#a6e3a1"), unsafe_allow_html=True)
            st.markdown("<br>", unsafe_allow_html=True)

        for err in errores:
            st.warning(err)

        if frames:
            df_result = pd.concat(frames, ignore_index=True)

            st.markdown("<div class='section-title'>Columnas extraidas</div>", unsafe_allow_html=True)
            ex1, ex2, ex3 = st.columns(3)
            ex1.markdown(metric_card("ARCHIVOS",     len(uploaded_ext),    "#cdd6f4"), unsafe_allow_html=True)
            ex2.markdown(metric_card("FILAS TOTALES", len(df_result),      "#89b4fa"), unsafe_allow_html=True)
            ex3.markdown(metric_card("COLUMNAS",     len(df_result.columns), "#a6e3a1"), unsafe_allow_html=True)
            st.markdown("<br>", unsafe_allow_html=True)

            st.dataframe(df_result.head(50), use_container_width=True, hide_index=True)
            if len(df_result) > 50:
                st.caption(f"Mostrando 50 de {len(df_result)} filas.")

            ts_ext = datetime.now().strftime("%d%m_%H%M")
            dl_csv_col, dl_xlsx_col = st.columns(2)
            dl_csv_col.download_button(
                label=f"📥 Descargar CSV ({len(df_result)} filas)",
                data=df_to_csv_bytes(df_result),
                file_name=f"extraccion_{ts_ext}.csv",
                mime="text/csv",
                use_container_width=True,
                key="dl_extractor_csv"
            )
            dl_xlsx_col.download_button(
                label=f"📥 Descargar Excel ({len(df_result)} filas)",
                data=df_to_excel_bytes(df_result),
                file_name=f"extraccion_{ts_ext}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="dl_extractor_xlsx"
            )
        elif not errores:
            st.info("Ningún archivo contenía las columnas requeridas.")
