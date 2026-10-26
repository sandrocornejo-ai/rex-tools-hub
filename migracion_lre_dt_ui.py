"""
migracion_lre_dt_ui.py
Elementos visuales de la pantalla "Migración desde LRE DT" (estilos Rex+, indicador de pasos, tarjetas).
Solo presentación: no contiene cálculos. Lo usa main() de migracion_lre_dt.py.
"""
import html

import streamlit as st

AZUL = "#1e5591"
CELESTE = "#00b4ff"
LIMA = "#b7d367"
OSCURO = "#1a2744"
GRIS = "#5b6475"
FONDO = "#f4f7fb"
BORDE = "#dde4ee"

PASOS = ["Archivos", "Revisión", "Maestro", "Archivo de carga", "Descargas"]

_CHECK = ('<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3" '
          'stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M5 12l5 5L20 7"/></svg>')
_ICONOS = {
    "ok": ('<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" '
           'stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M5 12l5 5L20 7"/></svg>'),
    "persona": ('<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" '
                'stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><circle cx="12" cy="8" r="4"/>'
                '<path d="M4 21c1.5-4 4.5-6 8-6s6.5 2 8 6"/></svg>'),
    "balanza": ('<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" '
                'stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M12 3v18"/><path d="M5 7h14"/>'
                '<path d="M5 7l-3 7a3 3 0 0 0 6 0z"/><path d="M19 7l-3 7a3 3 0 0 0 6 0z"/></svg>'),
    "lista": ('<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" '
              'stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M9 6h11"/><path d="M9 12h11"/>'
              '<path d="M9 18h11"/><path d="M4 6h.01"/><path d="M4 12h.01"/><path d="M4 18h.01"/></svg>'),
    "alerta": ('<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" '
               'stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M12 3l10 18H2z"/>'
               '<path d="M12 10v5"/><path d="M12 18h.01"/></svg>'),
}


def esc(texto):
    return html.escape(str(texto))


def md(texto_markdown_simple):
    """Convierte **negrita** a <b> y escapa el resto (para textos dentro de HTML)."""
    partes = str(texto_markdown_simple).split("**")
    return "".join(f"<b>{esc(p)}</b>" if i % 2 else esc(p) for i, p in enumerate(partes))


def estilos():
    """CSS de la página (solo esta página; no cambia el resto de Rex+ Tools)."""
    st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Red+Hat+Display:wght@400;500;600;700;800&display=swap');
[data-testid="stAppViewContainer"], [data-testid="stMain"] {{ background: {FONDO}; }}
[data-testid="stAppViewContainer"] h1, [data-testid="stAppViewContainer"] h2, [data-testid="stAppViewContainer"] h3,
[data-testid="stAppViewContainer"] p, [data-testid="stAppViewContainer"] label, [data-testid="stAppViewContainer"] button,
.rx {{ font-family: 'Red Hat Display', system-ui, sans-serif; }}
.block-container {{ max-width: 1180px; padding-top: 2.2rem; }}
.stButton > button, .stDownloadButton > button {{ border-radius: 999px; font-weight: 600; padding: 0.45rem 1.2rem; }}
.stButton > button[kind="primary"], .stDownloadButton > button[kind="primary"],
button[data-testid="stBaseButton-primary"], button[data-testid="baseButton-primary"] {{
  background: {AZUL}; border-color: {AZUL}; color: #ffffff; }}
.stButton > button[kind="primary"]:hover, button[data-testid="stBaseButton-primary"]:hover,
button[data-testid="baseButton-primary"]:hover {{ background: #174677; border-color: #174677; color: #ffffff; }}
[data-testid="stVerticalBlockBorderWrapper"] {{ background: #ffffff; border-radius: 16px; }}
[data-testid="stExpander"] details {{ background: #ffffff; border-radius: 14px; }}
.rx-eyebrow {{ font-size: 13px; font-weight: 700; color: #0093c9; letter-spacing: .06em; text-transform: uppercase; }}
.rx-title {{ margin: 2px 0 2px; font-size: 30px; font-weight: 800; color: {OSCURO}; line-height: 1.2; }}
.rx-sub {{ font-size: 14px; color: {GRIS}; }}
.rx-steps {{ display: grid; grid-template-columns: repeat(5, minmax(0, 1fr)); gap: 8px; margin: 10px 0 18px; }}
.rx-step {{ display: flex; align-items: center; gap: 10px; background: #fff; border: 1px solid {BORDE};
  border-radius: 12px; padding: 10px 12px; min-width: 0; }}
.rx-step .n {{ flex: 0 0 28px; height: 28px; border-radius: 50%; background: #eef2f7; color: {GRIS};
  display: flex; align-items: center; justify-content: center; font-weight: 800; font-size: 14px; }}
.rx-step .t {{ display: flex; flex-direction: column; min-width: 0; }}
.rx-step .t small {{ font-size: 12px; color: {GRIS}; }}
.rx-step .t span {{ font-size: 14px; font-weight: 700; color: #4a5568; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }}
.rx-step.done .n {{ background: {LIMA}; color: {OSCURO}; }}
.rx-step.done .t span {{ color: {OSCURO}; }}
.rx-step.now {{ background: {AZUL}; border-color: {AZUL}; }}
.rx-step.now .n {{ background: #fff; color: {AZUL}; }}
.rx-step.now .t small {{ color: #cfe3f7; }}
.rx-step.now .t span {{ color: #fff; }}
.rx-kpis {{ display: grid; grid-template-columns: repeat(5, minmax(0, 1fr)); gap: 12px; margin: 4px 0 18px; }}
.rx-kpi {{ background: #fff; border: 1px solid {BORDE}; border-radius: 14px; padding: 14px 16px;
  display: flex; flex-direction: column; gap: 2px; }}
.rx-kpi .l {{ font-size: 13px; color: {GRIS}; }}
.rx-kpi .v {{ font-size: 26px; font-weight: 800; color: {OSCURO}; }}
.rx-kpi .s {{ font-size: 12px; color: {GRIS}; }}
.rx-kpi .s.ok {{ color: #2f7d32; font-weight: 700; }}
.rx-kpi .s.warn {{ color: #8a5300; font-weight: 700; }}
.rx-kpi.warn {{ background: #fff7e6; border-color: #f3d9a4; }}
.rx-kpi.warn .l, .rx-kpi.warn .v, .rx-kpi.warn .s {{ color: #7a4a00; }}
.rx-h2 {{ font-size: 19px; font-weight: 800; color: {OSCURO}; margin: 6px 0 2px; }}
.rx-muted {{ font-size: 14px; color: {GRIS}; }}
.rx-item {{ display: flex; gap: 14px; align-items: flex-start; }}
.rx-ico {{ flex: 0 0 40px; height: 40px; border-radius: 12px; display: flex; align-items: center; justify-content: center; }}
.rx-ico.warn {{ background: #fff1d6; color: #a15c00; }}
.rx-ico.ok {{ background: #e3f0c7; color: #3f6212; }}
.rx-ico.err {{ background: #fde2e2; color: #b42318; }}
.rx-item .tt {{ font-size: 16px; font-weight: 700; color: {OSCURO}; }}
.rx-item .tt.ok {{ color: #3f6212; }}
.rx-item .dd {{ font-size: 14px; color: #4a5568; line-height: 1.45; }}
.rx-pill {{ display: inline-block; font-size: 13px; font-weight: 700; border-radius: 999px; padding: 4px 12px; }}
.rx-pill.ok {{ color: #3f6212; background: #e3f0c7; }}
.rx-pill.warn {{ color: #7a4a00; background: #fff1d6; }}
.rx-pill.info {{ color: {AZUL}; background: #e6f0fa; }}
.rx-up {{ display: flex; align-items: center; gap: 12px; }}
.rx-up .n {{ flex: 0 0 30px; height: 30px; border-radius: 50%; display: flex; align-items: center; justify-content: center;
  font-weight: 800; font-size: 14px; background: #e6f0fa; color: {AZUL}; }}
.rx-up .n.ok {{ background: {LIMA}; color: {OSCURO}; }}
.rx-up .t {{ display: flex; flex-direction: column; }}
.rx-up .t small {{ font-size: 12px; color: {GRIS}; }}
.rx-up .t span {{ font-size: 16px; font-weight: 700; color: {OSCURO}; }}
.rx-up-st {{ font-size: 14px; color: #4a5568; margin: 8px 0 2px; }}
.rx-chip {{ display: inline-block; font-size: 13px; background: #eef4fb; color: {AZUL}; border-radius: 999px;
  padding: 4px 11px; margin: 0 6px 6px 0; }}
.rx-chip.err {{ background: #fde2e2; color: #b42318; }}
.rx-progress {{ background: #fff; border: 1px solid {BORDE}; border-radius: 14px; padding: 14px 18px; margin-bottom: 12px; }}
.rx-progress .bar {{ height: 8px; border-radius: 999px; background: #eef2f7; overflow: hidden; margin-top: 8px; }}
.rx-progress .bar div {{ height: 8px; background: {CELESTE}; border-radius: 999px; }}
.rx-banner {{ display: flex; gap: 16px; align-items: center; border-radius: 16px; padding: 18px 22px; margin: 4px 0 16px; }}
.rx-banner.ok {{ background: #f3f9e8; border: 1px solid #d6e8b0; }}
.rx-banner.warn {{ background: #fff7e6; border: 1px solid #f3d9a4; }}
.rx-banner.err {{ background: #fdecec; border: 1px solid #f5c2c2; }}
.rx-banner .tt {{ font-size: 18px; font-weight: 800; color: {OSCURO}; }}
.rx-banner .dd {{ font-size: 14px; color: #4a5568; }}
.rx-file .tag {{ display: inline-block; font-size: 12px; font-weight: 700; border-radius: 999px; padding: 3px 10px; }}
.rx-file .tag.main {{ background: {AZUL}; color: #fff; }}
.rx-file .tag.info {{ background: #e6f0fa; color: {AZUL}; }}
.rx-file .tag.warn {{ background: #fff1d6; color: #7a4a00; }}
.rx-file .nm {{ font-size: 17px; font-weight: 800; color: {OSCURO}; margin-top: 8px; }}
.rx-file .ds {{ font-size: 14px; color: #4a5568; }}
.rx-file .fn {{ font-size: 12px; color: {GRIS}; font-family: Consolas, monospace; margin: 2px 0 8px; word-break: break-all; }}
[data-testid="stSlider"] [role="slider"] {{ background-color: {AZUL}; border-color: {AZUL}; }}
[data-testid="stSlider"] [data-testid="stSliderThumbValue"], [data-testid="stSliderTickBarMin"],
[data-testid="stSliderTickBarMax"] {{ color: {AZUL}; }}
@media (max-width: 900px) {{
  .rx-steps, .rx-kpis {{ grid-template-columns: repeat(2, minmax(0, 1fr)); }}
}}
</style>
""", unsafe_allow_html=True)


def encabezado(titulo, subtitulo=""):
    st.markdown(
        f'<div class="rx"><div class="rx-eyebrow">Migración desde LRE DT</div>'
        f'<div class="rx-title">{esc(titulo)}</div>'
        f'<div class="rx-sub">{md(subtitulo)}</div></div>', unsafe_allow_html=True)


def indicador_pasos(paso_actual):
    celdas = []
    for i, nombre in enumerate(PASOS, start=1):
        if i < paso_actual:
            clase, num, etq = "done", _CHECK, f"Paso {i}"
        elif i == paso_actual:
            clase, num, etq = "now", str(i), f"Paso {i} · ahora"
        else:
            clase, num, etq = "", str(i), f"Paso {i}"
        celdas.append(f'<div class="rx-step {clase}"><div class="n">{num}</div>'
                      f'<div class="t"><small>{etq}</small><span>{esc(nombre)}</span></div></div>')
    st.markdown(f'<nav class="rx rx-steps" aria-label="Pasos">{"".join(celdas)}</nav>', unsafe_allow_html=True)


def tarjetas_resumen(tarjetas):
    """tarjetas: [(etiqueta, valor, sub, estado)] con estado '', 'ok' o 'warn'."""
    html_t = []
    for etq, valor, sub, estado in tarjetas:
        cls_card = " warn" if estado == "warn-card" else ""
        cls_sub = {"ok": " ok", "warn": " warn"}.get(estado, "")
        html_t.append(f'<div class="rx-kpi{cls_card}"><span class="l">{esc(etq)}</span>'
                      f'<span class="v">{esc(valor)}</span><span class="s{cls_sub}">{esc(sub)}</span></div>')
    st.markdown(f'<section class="rx rx-kpis" aria-label="Resumen">{"".join(html_t)}</section>',
                unsafe_allow_html=True)


def titulo_seccion(texto, sub=""):
    st.markdown(f'<div class="rx"><div class="rx-h2">{esc(texto)}</div>'
                + (f'<div class="rx-muted">{md(sub)}</div>' if sub else "") + "</div>", unsafe_allow_html=True)


def item(titulo, detalle, estado="warn", icono="alerta"):
    """Fila con ícono, título y detalle (estado: warn, ok, err)."""
    ico = _ICONOS["ok"] if estado == "ok" else _ICONOS.get(icono, _ICONOS["alerta"])
    st.markdown(
        f'<div class="rx rx-item"><div class="rx-ico {estado}">{ico}</div><div>'
        f'<div class="tt {"ok" if estado == "ok" else ""}">{md(titulo)}</div>'
        f'<div class="dd">{md(detalle)}</div></div></div>', unsafe_allow_html=True)


def pill(texto, estado="info"):
    return f'<span class="rx-pill {estado}">{esc(texto)}</span>'


def tarjeta_archivo_titulo(numero, tipo, titulo, cargado):
    num = _CHECK if cargado else str(numero)
    return (f'<div class="rx rx-up"><div class="n {"ok" if cargado else ""}">{num}</div>'
            f'<div class="t"><small>Archivo {numero} · {esc(tipo)}</small><span>{esc(titulo)}</span></div></div>')


def progreso(cargados, total, faltan):
    pct = int(round(cargados / total * 100))
    texto_falta = ("Falta: <b>" + ", ".join(esc(f) for f in faltan) + "</b>") if faltan else "Todos los archivos cargados."
    st.markdown(
        f'<div class="rx rx-progress"><div style="display:flex;flex-wrap:wrap;justify-content:space-between;gap:8px">'
        f'<b style="font-size:15px">{cargados} de {total} archivos cargados</b>'
        f'<span class="rx-muted">{texto_falta}</span></div>'
        f'<div class="bar"><div style="width:{pct}%"></div></div></div>', unsafe_allow_html=True)


def banner(titulo, detalle, estado="ok"):
    ico = _ICONOS["ok"] if estado == "ok" else _ICONOS["alerta"]
    fondo = {"ok": LIMA, "warn": "#ffd98a", "err": "#f5b5b5"}[estado]
    st.markdown(
        f'<div class="rx rx-banner {estado}"><div class="rx-ico" style="flex:0 0 46px;height:46px;background:{fondo};'
        f'color:{OSCURO}">{ico}</div><div><div class="tt">{md(titulo)}</div><div class="dd">{md(detalle)}</div>'
        f'</div></div>', unsafe_allow_html=True)


def tarjeta_descarga(etiqueta, clase_etq, nombre, descripcion, archivo):
    st.markdown(
        f'<div class="rx rx-file"><span class="tag {clase_etq}">{esc(etiqueta)}</span>'
        f'<div class="nm">{esc(nombre)}</div><div class="ds">{esc(descripcion)}</div>'
        f'<div class="fn">{esc(archivo)}</div></div>', unsafe_allow_html=True)
