import streamlit as st
import pandas as pd
from io import BytesIO

from lector import leer_multiples_libros
from comparador import cargar_conceptos_rex, cruzar_conceptos

# ─────────────────────────────────────────────
# CONFIGURACIÓN DE PÁGINA
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Rex+ | Migrador Único",
    page_icon="🔀",
    layout="wide"
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

.rex-header {
    background-color: #1a2744;
    padding: 14px 28px;
    border-radius: 10px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    margin-bottom: 28px;
}
.rex-logo {
    background: white; color: #1a2744; font-weight: 800;
    font-size: 15px; padding: 5px 10px; border-radius: 6px; letter-spacing: 0.5px;
}
.rex-logo span { color: #00b4d8; }
.rex-title { color: white; font-size: 18px; font-weight: 600; margin-left: 16px; }
.rex-badge {
    background: #00b4d8; color: white; font-size: 11px;
    font-weight: 700; padding: 4px 12px; border-radius: 20px; letter-spacing: 1px;
}
.step-card {
    background: white; border: 1px solid #e8edf5;
    border-radius: 10px; padding: 18px 20px; margin-bottom: 12px;
}
.step-label {
    color: #00b4d8; font-size: 11px; font-weight: 700;
    letter-spacing: 1px; text-transform: uppercase; margin-bottom: 4px;
}
.step-title { font-size: 15px; font-weight: 600; color: #1a2744; margin-bottom: 4px; }
.step-desc { font-size: 13px; color: #6b7a9a; }
.section-title { font-size: 20px; font-weight: 700; color: #1a2744; margin-bottom: 4px; }
.section-sub { font-size: 13px; color: #6b7a9a; margin-bottom: 20px; }
.alert-success {
    background: #f0fff4; border-left: 4px solid #38a169;
    border-radius: 6px; padding: 12px 16px; margin: 8px 0;
    font-size: 13px; color: #276749;
}
.alert-warning {
    background: #fffbf0; border-left: 4px solid #d69e2e;
    border-radius: 6px; padding: 12px 16px; margin: 8px 0;
    font-size: 13px; color: #744210;
}
.alert-error {
    background: #fff0f0; border-left: 4px solid #e53e3e;
    border-radius: 6px; padding: 12px 16px; margin: 8px 0;
    font-size: 13px; color: #c53030;
}
.stButton > button {
    background-color: #1a2744 !important; color: white !important;
    border: none !important; border-radius: 8px !important;
    font-weight: 600 !important; padding: 8px 20px !important; font-size: 14px !important;
}
.stButton > button:hover { background-color: #00b4d8 !important; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────
st.markdown("""
<div class="rex-header">
    <div style="display:flex;align-items:center">
        <div class="rex-logo">Rex<span>+</span></div>
        <div class="rex-title">Migrador Único — Cruce de conceptos</div>
    </div>
    <div class="rex-badge">ETAPA 1</div>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────
if "mu_resultados"   not in st.session_state: st.session_state.mu_resultados   = None
if "mu_df_rex"       not in st.session_state: st.session_state.mu_df_rex       = None
if "mu_conceptos"    not in st.session_state: st.session_state.mu_conceptos    = None
if "mu_asignaciones" not in st.session_state: st.session_state.mu_asignaciones = {}

# ─────────────────────────────────────────────
# PASO 1 — CARGA DE ARCHIVOS
# ─────────────────────────────────────────────
st.markdown('<div class="section-title">📂 Paso 1 — Carga de archivos</div>', unsafe_allow_html=True)
st.markdown('<div class="section-sub">Sube los libros del cliente y el listado de conceptos Rex+.</div>', unsafe_allow_html=True)

col_a, col_b = st.columns(2)

with col_a:
    st.markdown("""
    <div class="step-card">
        <div class="step-label">Entrada 1</div>
        <div class="step-title">📋 Libros de remuneraciones</div>
        <div class="step-desc">Uno o más archivos Excel del cliente (máx. 10).</div>
    </div>
    """, unsafe_allow_html=True)
    libros = st.file_uploader(
        "Libros del cliente",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key="mu_libros",
        label_visibility="collapsed"
    )

with col_b:
    st.markdown("""
    <div class="step-card">
        <div class="step-label">Entrada 2</div>
        <div class="step-title">📑 Listado de conceptos Rex+</div>
        <div class="step-desc">Archivo exportado desde Rex con columnas: Concepto, Nombre, Tipo.</div>
    </div>
    """, unsafe_allow_html=True)
    archivo_rex = st.file_uploader(
        "Conceptos Rex+",
        type=["xlsx", "xls"],
        key="mu_rex",
        label_visibility="collapsed"
    )

# ─────────────────────────────────────────────
# PASO 2 — PROCESAR
# ─────────────────────────────────────────────
st.markdown("---")

listo_para_procesar = libros and len(libros) <= 10 and archivo_rex

if libros and len(libros) > 10:
    st.markdown('<div class="alert-error">❌ Máximo 10 libros a la vez.</div>', unsafe_allow_html=True)

if listo_para_procesar:
    if st.button("🔀 Cruzar conceptos"):
        with st.spinner("Leyendo libros y cruzando conceptos..."):
            # Leer libros
            todos_conceptos, errores_lectura = leer_multiples_libros(libros)
            if errores_lectura:
                for err in errores_lectura:
                    st.markdown(f'<div class="alert-error">⚠️ {err}</div>', unsafe_allow_html=True)

            # Leer Rex+
            rex_bytes = archivo_rex.read()
            df_rex = cargar_conceptos_rex(rex_bytes)

            # Cruzar
            lista_conceptos = list(todos_conceptos.keys())
            resultados = cruzar_conceptos(lista_conceptos, df_rex)

            # Guardar en session state
            st.session_state.mu_resultados   = resultados
            st.session_state.mu_df_rex       = df_rex
            st.session_state.mu_conceptos    = todos_conceptos
            st.session_state.mu_asignaciones = {}
            st.rerun()

# ─────────────────────────────────────────────
# PASO 3 — MOSTRAR RESULTADOS
# ─────────────────────────────────────────────
if st.session_state.mu_resultados is not None:
    res       = st.session_state.mu_resultados
    df_rex    = st.session_state.mu_df_rex
    conceptos = st.session_state.mu_conceptos
    asigs     = st.session_state.mu_asignaciones

    n_match    = len(res["match"])
    n_dudoso   = len(res["dudoso"])
    n_sin      = len(res["sin_match"])
    n_total    = n_match + n_dudoso + n_sin

    st.markdown('<div class="section-title">📊 Resultado del cruce</div>', unsafe_allow_html=True)

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total conceptos", n_total)
    m2.metric("✅ Match automático", n_match)
    m3.metric("⚠️ Dudosos", n_dudoso)
    m4.metric("❌ Sin match", n_sin)

    st.markdown("---")

    # ── TAB RESULTS ──────────────────────────────────────────────────────────
    tab_match, tab_dudoso, tab_sin = st.tabs([
        f"✅ Matches ({n_match})",
        f"⚠️ Dudosos ({n_dudoso})",
        f"❌ Sin match ({n_sin})",
    ])

    # ── MATCHES AUTOMÁTICOS ──────────────────────────────────────────────────
    with tab_match:
        if res["match"]:
            df_m = pd.DataFrame(res["match"])
            df_m = df_m.rename(columns={
                "col_cliente":  "Concepto cliente",
                "concepto_rex": "Código Rex+",
                "nombre_rex":   "Nombre Rex+",
                "tipo_rex":     "Tipo",
                "score":        "Score",
                "metodo":       "Método",
            })
            st.dataframe(df_m, use_container_width=True, hide_index=True)
        else:
            st.markdown('<div class="alert-warning">No se encontraron matches automáticos.</div>', unsafe_allow_html=True)

    # ── DUDOSOS ──────────────────────────────────────────────────────────────
    with tab_dudoso:
        if res["dudoso"]:
            st.markdown('<div class="alert-warning">⚠️ Estos conceptos tienen una coincidencia dudosa. Selecciona la opción correcta o descártala.</div>', unsafe_allow_html=True)
            st.markdown("")

            # Opciones Rex+ para el selectbox
            opciones_rex = ["— Sin asignar —"] + [
                f"{row['Concepto']} | {row['Nombre']} ({row['Tipo']})"
                for _, row in df_rex.iterrows()
            ]

            for item in res["dudoso"]:
                col_cli = item["col_cliente"]
                sugs    = item["sugerencias"]

                with st.expander(f"📌 {col_cli}", expanded=True):
                    st.markdown(f"**Nombre limpio analizado:** `{item['col_limpia']}`")
                    st.markdown("**Sugerencias:**")

                    sug_options = ["— Ninguna de estas —"] + [
                        f"{s['concepto']} | {s['nombre']} ({s['tipo']})  —  {s['score']:.0f}%"
                        for s in sugs
                    ]

                    sel = st.radio(
                        "Selecciona el match correcto:",
                        sug_options,
                        key=f"mu_radio_{col_cli}",
                        horizontal=False,
                    )

                    if sel != "— Ninguna de estas —":
                        # Guardar la selección
                        idx_sug = sug_options.index(sel) - 1
                        sug_elegida = sugs[idx_sug]
                        asigs[col_cli] = {
                            "concepto_rex": sug_elegida["concepto"],
                            "nombre_rex":   sug_elegida["nombre"],
                            "tipo_rex":     sug_elegida["tipo"],
                            "metodo":       "Manual (dudoso confirmado)",
                        }
                    else:
                        # Asignación manual libre
                        sel_libre = st.selectbox(
                            "O asigna manualmente desde Rex+:",
                            opciones_rex,
                            key=f"mu_libre_{col_cli}",
                        )
                        if sel_libre != "— Sin asignar —":
                            codigo = sel_libre.split(" | ")[0]
                            fila = df_rex[df_rex["Concepto"] == codigo].iloc[0]
                            asigs[col_cli] = {
                                "concepto_rex": fila["Concepto"],
                                "nombre_rex":   fila["Nombre"],
                                "tipo_rex":     fila["Tipo"],
                                "metodo":       "Manual (desde dudoso)",
                            }
                        elif col_cli in asigs:
                            del asigs[col_cli]
        else:
            st.markdown('<div class="alert-success">✅ No hay conceptos dudosos.</div>', unsafe_allow_html=True)

    # ── SIN MATCH ────────────────────────────────────────────────────────────
    with tab_sin:
        if res["sin_match"]:
            st.markdown('<div class="alert-error">❌ Estos conceptos no tienen equivalencia en Rex+. Puedes asignarlos manualmente.</div>', unsafe_allow_html=True)
            st.markdown("")

            opciones_rex_sin = ["— Sin asignar —"] + [
                f"{row['Concepto']} | {row['Nombre']} ({row['Tipo']})"
                for _, row in df_rex.iterrows()
            ]

            for item in res["sin_match"]:
                col_cli = item["col_cliente"]
                with st.expander(f"❌ {col_cli}"):
                    sel = st.selectbox(
                        "Asignar concepto Rex+:",
                        opciones_rex_sin,
                        key=f"mu_sin_{col_cli}",
                    )
                    if sel != "— Sin asignar —":
                        codigo = sel.split(" | ")[0]
                        fila = df_rex[df_rex["Concepto"] == codigo].iloc[0]
                        asigs[col_cli] = {
                            "concepto_rex": fila["Concepto"],
                            "nombre_rex":   fila["Nombre"],
                            "tipo_rex":     fila["Tipo"],
                            "metodo":       "Manual (sin match)",
                        }
                    elif col_cli in asigs:
                        del asigs[col_cli]
        else:
            st.markdown('<div class="alert-success">✅ Todos los conceptos tienen equivalencia.</div>', unsafe_allow_html=True)

    # ─────────────────────────────────────────────
    # EXPORTAR RESULTADO FINAL
    # ─────────────────────────────────────────────
    st.markdown("---")
    st.markdown('<div class="section-title">💾 Exportar mapeo final</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-sub">Descarga el Excel con todos los conceptos mapeados (automáticos + manuales).</div>', unsafe_allow_html=True)

    # Construir tabla final
    filas = []

    for m in res["match"]:
        filas.append({
            "Concepto cliente":       m["col_cliente"],
            "Archivos":               ", ".join(conceptos.get(m["col_cliente"], [])),
            "Código Rex+":            m["concepto_rex"],
            "Nombre Rex+":            m["nombre_rex"],
            "Tipo Rex+":              m["tipo_rex"],
            "Score":                  m["score"],
            "Método":                 m["metodo"],
        })

    for d in res["dudoso"]:
        col_cli = d["col_cliente"]
        asig = asigs.get(col_cli)
        filas.append({
            "Concepto cliente": col_cli,
            "Archivos":         ", ".join(conceptos.get(col_cli, [])),
            "Código Rex+":      asig["concepto_rex"] if asig else "",
            "Nombre Rex+":      asig["nombre_rex"]   if asig else "",
            "Tipo Rex+":        asig["tipo_rex"]      if asig else "",
            "Score":            "",
            "Método":           asig["metodo"]        if asig else "Pendiente",
        })

    for s in res["sin_match"]:
        col_cli = s["col_cliente"]
        asig = asigs.get(col_cli)
        filas.append({
            "Concepto cliente": col_cli,
            "Archivos":         ", ".join(conceptos.get(col_cli, [])),
            "Código Rex+":      asig["concepto_rex"] if asig else "",
            "Nombre Rex+":      asig["nombre_rex"]   if asig else "",
            "Tipo Rex+":        asig["tipo_rex"]      if asig else "",
            "Score":            "",
            "Método":           asig["metodo"]        if asig else "Sin asignar",
        })

    df_final = pd.DataFrame(filas)

    pendientes = df_final[df_final["Método"].isin(["Pendiente", "Sin asignar"])].shape[0]
    if pendientes > 0:
        st.markdown(f'<div class="alert-warning">⚠️ Hay <b>{pendientes}</b> concepto(s) aún sin asignar. Puedes exportar igualmente y completarlos después.</div>', unsafe_allow_html=True)
    else:
        st.markdown('<div class="alert-success">✅ Todos los conceptos están mapeados. Listo para exportar.</div>', unsafe_allow_html=True)

    st.dataframe(df_final, use_container_width=True, hide_index=True)

    # Generar Excel
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df_final.to_excel(writer, index=False, sheet_name="Mapeo conceptos")
    buf.seek(0)

    st.download_button(
        label="⬇️ Descargar mapeo (.xlsx)",
        data=buf,
        file_name="mapeo_conceptos.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
