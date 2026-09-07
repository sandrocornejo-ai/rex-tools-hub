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
if "mu_eliminados"      not in st.session_state: st.session_state.mu_eliminados      = set()
if "mu_match_overrides" not in st.session_state: st.session_state.mu_match_overrides = {}

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
            match_overrides = st.session_state.mu_match_overrides
            eliminados      = st.session_state.mu_eliminados

            st.markdown(
                '<div class="section-sub">'
                'Puedes cambiar la asignación de cualquier concepto usando el selector de la derecha. '
                'Marca la casilla <b>Eliminar</b> para quitar filas duplicadas y presiona el botón.</div>',
                unsafe_allow_html=True,
            )
            st.markdown("")

            # Encabezados
            h1, h2, h3 = st.columns([2, 3, 1])
            h1.markdown("**Concepto del cliente**")
            h2.markdown("**Concepto Rex+ asignado**")
            h3.markdown("**Eliminar**")
            st.markdown('<hr style="margin:4px 0 12px 0">', unsafe_allow_html=True)

            opciones_rex_m = ["— Sin asignar —"] + [
                f"{row['Concepto']} | {row['Nombre']} ({row['Tipo']})"
                for _, row in df_rex.iterrows()
            ]

            for m in res["match"]:
                col_cli = m["col_cliente"]
                # Usar override si existe, sino el match original
                override = match_overrides.get(col_cli)
                actual_concepto = override["concepto_rex"] if override else m["concepto_rex"]
                actual_nombre   = override["nombre_rex"]   if override else m["nombre_rex"]
                actual_tipo     = override["tipo_rex"]      if override else m["tipo_rex"]

                c_left, c_mid, c_right = st.columns([2, 3, 1])

                with c_left:
                    score_col = "#38a169" if m["score"] >= 90 else "#d69e2e"
                    st.markdown(
                        f'<div style="padding:6px 0 10px 0">'
                        f'<div style="font-weight:600;color:#1a2744;font-size:14px;">{col_cli}</div>'
                        f'<div style="font-size:12px;color:#6b7a9a;margin-top:2px;">'
                        f'Score: <span style="color:{score_col};font-weight:700;">{m["score"]:.0f}%</span>'
                        f' · {m["metodo"]}</div></div>',
                        unsafe_allow_html=True,
                    )

                with c_mid:
                    best_plain = f"{actual_concepto} | {actual_nombre} ({actual_tipo})"
                    default_idx = opciones_rex_m.index(best_plain) if best_plain in opciones_rex_m else 0
                    sel = st.selectbox(
                        f"match_{col_cli}",
                        opciones_rex_m,
                        index=default_idx,
                        key=f"mu_match_sel_{col_cli}",
                        label_visibility="collapsed",
                    )
                    if sel != "— Sin asignar —":
                        codigo = sel.split(" | ")[0]
                        fila   = df_rex[df_rex["Concepto"] == codigo].iloc[0]
                        match_overrides[col_cli] = {
                            "concepto_rex": fila["Concepto"],
                            "nombre_rex":   fila["Nombre"],
                            "tipo_rex":     fila["Tipo"],
                        }
                    elif col_cli in match_overrides:
                        del match_overrides[col_cli]

                with c_right:
                    eliminar = st.checkbox(
                        "del",
                        value=col_cli in eliminados,
                        key=f"mu_del_{col_cli}",
                        label_visibility="collapsed",
                    )
                    if eliminar:
                        eliminados.add(col_cli)
                    elif col_cli in eliminados:
                        eliminados.discard(col_cli)

                st.markdown('<hr style="margin:4px 0 8px 0;border-color:#e8edf5">', unsafe_allow_html=True)

            if eliminados:
                if st.button(f"Eliminar {len(eliminados)} fila(s) marcada(s)"):
                    st.session_state.mu_resultados["match"] = [
                        m for m in res["match"] if m["col_cliente"] not in eliminados
                    ]
                    for k in list(eliminados):
                        match_overrides.pop(k, None)
                    st.session_state.mu_eliminados = set()
                    st.rerun()
        else:
            st.markdown('<div class="alert-warning">No se encontraron matches automáticos.</div>', unsafe_allow_html=True)

    # ── DUDOSOS ──────────────────────────────────────────────────────────────
    with tab_dudoso:
        if res["dudoso"]:
            st.markdown(
                '<div class="alert-warning">⚠️ El sistema no pudo confirmar estos conceptos. '
                'Elige en la columna derecha el concepto Rex+ correcto para cada uno.</div>',
                unsafe_allow_html=True,
            )
            st.markdown("")
            h1, h2 = st.columns([2, 3])
            h1.markdown("**Concepto del cliente**")
            h2.markdown("**Asignar a concepto Rex+**")
            st.markdown('<hr style="margin:4px 0 12px 0">', unsafe_allow_html=True)
            opciones_rex = ["— Sin asignar —"] + [
                f"{row['Concepto']} | {row['Nombre']} ({row['Tipo']})"
                for _, row in df_rex.iterrows()
            ]
            for item in res["dudoso"]:
                col_cli = item["col_cliente"]
                sugs    = item["sugerencias"]
                mejor   = sugs[0] if sugs else None
                c_left, c_right = st.columns([2, 3])
                with c_left:
                    score_txt = f"{mejor['score']:.0f}%" if mejor else "—"
                    score_col = "#d69e2e" if mejor and mejor["score"] >= 70 else "#e53e3e"
                    st.markdown(
                        f'<div style="padding:6px 0 10px 0">'
                        f'<div style="font-weight:600;color:#1a2744;font-size:14px;">{col_cli}</div>'
                        f'<div style="font-size:12px;color:#6b7a9a;margin-top:2px;">Mejor coincidencia: '
                        f'<span style="color:{score_col};font-weight:700;">{score_txt}</span></div></div>',
                        unsafe_allow_html=True,
                    )
                with c_right:
                    default_idx = 0
                    if mejor:
                        best_plain = f"{mejor['concepto']} | {mejor['nombre']} ({mejor['tipo']})"
                        if best_plain in opciones_rex:
                            default_idx = opciones_rex.index(best_plain)
                    sel = st.selectbox(
                        f"concepto_{col_cli}",
                        opciones_rex,
                        index=default_idx,
                        key=f"mu_dud_{col_cli}",
                        label_visibility="collapsed",
                    )
                    if sel != "— Sin asignar —":
                        codigo = sel.split(" | ")[0]
                        fila = df_rex[df_rex["Concepto"] == codigo].iloc[0]
                        asigs[col_cli] = {
                            "concepto_rex": fila["Concepto"],
                            "nombre_rex":   fila["Nombre"],
                            "tipo_rex":     fila["Tipo"],
                            "metodo":       "Manual (dudoso)",
                        }
                    elif col_cli in asigs:
                        del asigs[col_cli]
                st.markdown('<hr style="margin:4px 0 8px 0;border-color:#e8edf5">', unsafe_allow_html=True)
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

    match_ov = st.session_state.get("mu_match_overrides", {})
    for m in res["match"]:
        ov = match_ov.get(m["col_cliente"])
        filas.append({
            "Concepto cliente": m["col_cliente"],
            "Archivos":         ", ".join(conceptos.get(m["col_cliente"], [])),
            "Código Rex+":      ov["concepto_rex"] if ov else m["concepto_rex"],
            "Nombre Rex+":      ov["nombre_rex"]   if ov else m["nombre_rex"],
            "Tipo Rex+":        ov["tipo_rex"]      if ov else m["tipo_rex"],
            "Score":            m["score"],
            "Método":           ("Editado manualmente" if ov else m["metodo"]),
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
