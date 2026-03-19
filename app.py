import tempfile
from pathlib import Path
import streamlit as st
import motor_prefacturacion as motor

st.set_page_config(page_title="Prefacturación Excel", page_icon="📄", layout="centered")

st.title("Prefacturación de archivos Excel")
st.caption("Carga el archivo de pedidos, ejecuta la conciliación y descarga el resultado.")

with st.container():
    archivo_subido = st.file_uploader(
        "Archivo de pedidos",
        type=["xlsx"],
        help="Debe contener las hojas MERCANCIA, PAQUETE y/o DOCUMENTO según el caso."
    )

    ruta_tarifas = Path("Valor kilo destino Colvanes 2026.xlsx")

    st.markdown("### Procesos a ejecutar")
    proc_mercancia = st.checkbox("Mercancía", value=True)
    proc_paquete = st.checkbox("Paquete", value=True)
    proc_documento = st.checkbox("Documento", value=True)

    ejecutar = st.button("Procesar archivo", type="primary", use_container_width=True)

if ejecutar:
    if archivo_subido is None:
        st.error("Primero debes subir un archivo Excel.")
        st.stop()

    if not ruta_tarifas.exists():
        st.error("No se encontró el archivo de tarifas: `Valor kilo destino Colvanes 2026.xlsx`")
        st.stop()

    if not any([proc_mercancia, proc_paquete, proc_documento]):
        st.error("Selecciona al menos un proceso.")
        st.stop()

    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir = Path(tmpdir)
            ruta_pedidos = tmpdir / archivo_subido.name

            with open(ruta_pedidos, "wb") as f:
                f.write(archivo_subido.getbuffer())

            motor.ARCHIVO_PEDIDOS = str(ruta_pedidos)
            motor.SALIDA = str(ruta_pedidos)
            motor.ARCHIVO_TARIFAS = str(ruta_tarifas)

            tareas = []
            if proc_mercancia:
                tareas.append(("Mercancía", motor.prefacturar))
            if proc_paquete:
                tareas.append(("Paquete", motor.prefacturar_paquete))
            if proc_documento:
                tareas.append(("Documento", motor.prefacturar_documento))

            progreso = st.progress(0, text="Iniciando proceso...")
            total = len(tareas)

            for i, (nombre, funcion) in enumerate(tareas, start=1):
                progreso.progress((i - 1) / total, text=f"Procesando: {nombre}...")
                funcion()

            progreso.progress(1.0, text="Proceso completado")

            st.success("El archivo fue procesado correctamente.")

            with open(ruta_pedidos, "rb") as f:
                st.download_button(
                    label="Descargar resultado",
                    data=f.read(),
                    file_name=f"resultado_{archivo_subido.name}",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

    except KeyError as e:
        st.error(f"Falta una columna esperada en el Excel: {e}")
    except Exception as e:
        st.error(f"Ocurrió un error durante el procesamiento: {e}")
