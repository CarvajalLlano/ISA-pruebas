import tempfile
from pathlib import Path
import streamlit as st

import logica_prefacturacion as lp

st.set_page_config(page_title="Prefacturación Excel", page_icon="📄")

st.title("Prefacturación de archivos Excel")

archivo_subido = st.file_uploader("Sube el archivo de pedidos", type=["xlsx"])

if archivo_subido:

    st.success(f"Archivo cargado: {archivo_subido.name}")

    if st.button("Procesar archivo"):

        try:
            with tempfile.TemporaryDirectory() as tmpdir:
                tmpdir = Path(tmpdir)

                ruta_pedidos = tmpdir / archivo_subido.name
                ruta_tarifas = Path("Valor kilo destino Colvanes 2026.xlsx")

                if not ruta_tarifas.exists():
                    st.error("No se encontró el archivo de tarifas")
                    st.stop()

                # Guardar archivo subido
                with open(ruta_pedidos, "wb") as f:
                    f.write(archivo_subido.getbuffer())

                # Inyectar rutas a tu lógica
                lp.ARCHIVO_PEDIDOS = str(ruta_pedidos)
                lp.SALIDA = str(ruta_pedidos)
                lp.ARCHIVO_TARIFAS = str(ruta_tarifas)

                with st.spinner("Procesando..."):
                    lp.prefacturar()
                    lp.prefacturar_paquete()
                    lp.prefacturar_documento()

                st.success("Proceso terminado")

                # Descargar resultado
                with open(ruta_pedidos, "rb") as f:
                    st.download_button(
                        "Descargar resultado",
                        data=f.read(),
                        file_name=f"resultado_{archivo_subido.name}",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

        except Exception as e:
            st.error(f"Error: {e}")
