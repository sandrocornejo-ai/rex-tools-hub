"""
salir.py — Botón "Salir" de Rex+ Tools (barra lateral) y aviso al cerrar la pestaña.

- Mientras la app está abierta, si el usuario intenta cerrar la pestaña o la ventana, el navegador
  muestra su aviso de confirmación. (Chrome/Edge muestran su propio texto; no permiten uno personalizado.)
- El botón Salir quita ese aviso, borra los datos de la sesión (archivos subidos, ingresos manuales)
  y cierra la pestaña. Si el navegador no permite cerrarla (Chrome/Edge solo dejan cerrar pestañas
  abiertas por un script), la deja en blanco.
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

_JS_CERRAR = """
<script>
(function () {
  // El iframe de Streamlit está aislado (sandbox) y no puede cerrar ni navegar la página principal,
  // por eso el código se inserta como <script> en la página principal y corre desde ahí.
  var w = window.parent;
  var codigo = [
    "if (window.__rexAvisoCierre) {",
    "  window.removeEventListener('beforeunload', window.__rexAvisoCierre);",
    "  window.__rexAvisoCierre = null;",
    "}",
    "window.open('', '_self');",
    "window.close();",
    // Si el navegador no permite cerrar la pestaña, la deja en blanco (la aplicación ya no queda visible)
    "setTimeout(function () { window.location.replace('about:blank'); }, 300);"
  ].join("\\n");
  var sc = w.document.createElement("script");
  sc.textContent = codigo;
  w.document.body.appendChild(sc);
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


def _cerrar():
    """Quita el aviso de cierre, cierra la pestaña (o la deja en blanco) y, en el computador, detiene el programa."""
    _inyectar_js(_JS_CERRAR)
    if _es_local():
        threading.Timer(2.0, _detener_programa).start()
    st.stop()


def boton_salir():
    """Agrega el botón Salir en la barra lateral y activa el aviso al cerrar la pestaña."""
    if st.session_state.get(CLAVE_SALIR):
        _cerrar()
    _inyectar_js(_JS_ACTIVAR_AVISO)
    with st.sidebar:
        st.markdown("---")
        if st.button("🚪 Salir", type="primary", use_container_width=True, key="btn_salir_rex"):
            st.session_state.clear()
            st.session_state[CLAVE_SALIR] = True
            st.rerun()
        st.caption(f"⚠️ {MENSAJE_CIERRE}.")
