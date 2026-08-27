import streamlit as st

st.set_page_config(
    page_title="Rex+ Tools",
    page_icon="🛠️",
    layout="wide"
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

.header-bar {
    background: linear-gradient(135deg, #0f2d5e 0%, #1a4a8a 100%);
    padding: 18px 32px;
    border-radius: 12px;
    display: flex;
    align-items: center;
    gap: 16px;
    margin-bottom: 32px;
}
.header-bar h1 { color: white; font-size: 1.6rem; font-weight: 700; margin: 0; }
.header-bar span { color: #7eb3ff; font-size: 0.95rem; }

.tool-card {
    background: white;
    border: 1.5px solid #e8edf5;
    border-radius: 14px;
    padding: 28px;
    height: 100%;
    transition: box-shadow 0.2s;
    position: relative;
}
.tool-card:hover { box-shadow: 0 6px 24px rgba(15,45,94,0.10); }
.tool-icon { font-size: 2.4rem; margin-bottom: 14px; }
.tool-title { font-size: 1.1rem; font-weight: 700; color: #0f2d5e; margin-bottom: 8px; }
.tool-desc { font-size: 0.88rem; color: #555; line-height: 1.6; margin-bottom: 16px; }
.tag {
    display: inline-block;
    background: #eef4ff;
    color: #1a4a8a;
    font-size: 0.72rem;
    font-weight: 600;
    border-radius: 20px;
    padding: 3px 10px;
    margin-right: 6px;
    margin-bottom: 6px;
}
.section-title { font-size: 1.3rem; font-weight: 700; color: #0f2d5e; margin-bottom: 20px; }
</style>

<div class="header-bar">
    <div>
        <h1>🛠️ Rex+ Tools</h1>
        <span>Selecciona una herramienta para comenzar</span>
    </div>
</div>

<div class="section-title">Herramientas disponibles</div>
""", unsafe_allow_html=True)

col1, col2 = st.columns(2, gap="large")

with col1:
    st.markdown("""
    <div class="tool-card">
        <div class="tool-icon">🔄</div>
        <div class="tool-title">1 · Migración TeamWork → Rex+</div>
        <div class="tool-desc">
            Transforma el archivo de liquidaciones de TeamWork al formato
            de importación Rex+. Calcula bases AFP, CES, impuestos y genera
            el Excel listo para cargar.
        </div>
        <span class="tag">ACTIVA</span>
        <span class="tag">Excel</span>
        <span class="tag">Remuneraciones</span>
    </div>
    """, unsafe_allow_html=True)
    st.page_link("pages/1_Migracion_TW.py", label="Abrir herramienta →", use_container_width=True)

with col2:
    st.markdown("""
    <div class="tool-card">
        <div class="tool-icon">📊</div>
        <div class="tool-title">2 · LRE Detalle → Liquidaciones</div>
        <div class="tool-desc">
            Procesa el Libro de Remuneraciones Electrónico (LRE) y genera
            el archivo de liquidaciones en detalle listo para importar en Rex+.
        </div>
        <span class="tag">ACTIVA</span>
        <span class="tag">LRE</span>
        <span class="tag">Liquidaciones</span>
    </div>
    """, unsafe_allow_html=True)
    st.page_link("pages/2_LRE_Detalle.py", label="Abrir herramienta →", use_container_width=True)

st.divider()
st.caption("Rex+ Tools · Visma · Uso interno")
