"""
salir.py — Botón "Salir" de Rex+ Tools (barra lateral) y aviso al cerrar la pestaña.

- Mientras la app está abierta, si el usuario intenta cerrar la pestaña o la ventana, el navegador
  muestra su aviso de confirmación. (Chrome/Edge muestran su propio texto; no permiten uno personalizado.)
- El botón Salir quita ese aviso, borra los datos de la sesión (archivos subidos, ingresos manuales)
  y muestra "Ya puedes cerrar esta pestaña".
- En el computador (streamlit run) además detiene el programa. En Streamlit Cloud no lo detiene,
  porque cortaría la app para todos los usuarios.

Uso: llamar boton_salir() justo después de st.set_page_config() en cada página.
"""
import os
import threading

import streamlit as st

MENSAJE_CIERRE = "Use el botón Salir para cerrar la aplicación"
CLAVE_SALIR = "_rex_salir"

_JS_ACTIVAR_AVISO = f"""
<script>
(function () {{
  var w = window.parent;
  if (w.__rexAvisoCierre) return;
  w.__rexAvisoCierre = function (e) {{
    e.preventDefault();
    e.returnValue = "{MENSAJE_CIERRE}";
    return "{MENSAJE_CIERRE}";
  }};
  w.addEventListener("beforeunload", w.__rexAvisoCierre);
}})();
</script>
"""

_JS_QUITAR_AVISO = """
<script>
(function () {
  var w = window.parent;
  if (w.__rexAvisoCierre) {
    w.removeEventListener("beforeunload", w.__rexAvisoCierre);
    w.__rexAvisoCierre = null;
  }
})();
</script>
"""


def _inyectar_js(codigo):
    """Ejecuta JS en un iframe invisible (st.iframe en Streamlit nuevo; components.html en versiones antiguas)."""
    if hasattr(st, "iframe"):
        st.iframe(codigo, height=1)
    else:
        import streamlit.components.v1 as components
        components.html(codigo, height=0)


def _es_local():
    """True si corre en el computador; Streamlit Cloud monta el repositorio en /mount/src."""
    return not os.path.abspath(__file__).startswith("/mount/src")


def _detener_programa():
    os._exit(0)


def _pantalla_cerrada():
    _inyectar_js(_JS_QUITAR_AVISO)
    st.success("✅ Sesión cerrada. Los datos cargados se borraron de la memoria. Ya puedes cerrar esta pestaña.")
    if _es_local():
        st.info("El programa se detuvo en tu computador. Para volver a abrirlo: `streamlit run Home.py`")
        threading.Timer(2.0, _detener_programa).start()
    elif st.button("↩️ Volver a Rex+ Tools", key="btn_volver_rex"):
        st.session_state.clear()
        st.rerun()
    st.stop()


def boton_salir():
    """Agrega el botón Salir en la barra lateral y activa el aviso al cerrar la pestaña."""
    if st.session_state.get(CLAVE_SALIR):
        _pantalla_cerrada()
    _inyectar_js(_JS_ACTIVAR_AVISO)
    with st.sidebar:
        st.markdown("---")
        if st.button("🚪 Salir", type="primary", use_container_width=True, key="btn_salir_rex"):
            st.session_state.clear()
            st.session_state[CLAVE_SALIR] = True
            st.rerun()
        st.caption(f"⚠️ {MENSAJE_CIERRE}.")
