from io import BytesIO
from pathlib import Path
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Analizador Datos Contractus - Cruce FCM",
    page_icon="📊",
    layout="wide",
)


def normalizar_cc(serie: pd.Series) -> pd.Series:
    """Normaliza cédulas e identificaciones eliminando apóstrofes, decimales, espacios y ceros iniciales."""
    return (
        serie.fillna("")
        .astype(str)
        .str.replace("'", "", regex=False)
        .str.replace(r"\.0$", "", regex=True)
        .str.strip()
        .str.lstrip("0")
    )


def generar_excel_bytes(df: pd.DataFrame, nombre_hoja: str = "SOLO_EN_FCM") -> bytes:
    """Exporta un DataFrame a un archivo Excel en memoria con ajuste automático de ancho de columnas."""
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=nombre_hoja)
        ws = writer.sheets[nombre_hoja]
        for col in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 45)
    return buffer.getvalue()


st.title("Analizador de Datos Contractus — Cruce FCM")
st.markdown(
    "Selecciona los tres archivos Excel para identificar los registros presentes en **REPORTE FCM (hoja MANDAMIENTO)** "
    "cuya **Identificación / Cédula** no aparece en **Sanciones** ni en **Cobro Coactivo**."
)

col1, col2, col3 = st.columns(3)

with col1:
    archivo_fcm = st.file_uploader(
        "1. Reporte FCM (.xlsx)",
        type=["xlsx"],
        help="Se leerá únicamente la hoja 'MANDAMIENTO' usando la columna 'Identificación'.",
    )

with col2:
    archivo_sanciones = st.file_uploader(
        "2. Sanciones pendientes firma (.xlsx)",
        type=["xlsx"],
        help="Se leerá la columna 'Cédula'.",
    )

with col3:
    archivo_cobro = st.file_uploader(
        "3. Cobro coactivo pendientes firma (.xlsx)",
        type=["xlsx"],
        help="Se leerá la columna 'Cédula'.",
    )

if archivo_fcm and archivo_sanciones and archivo_cobro:
    try:
        xls_fcm = pd.ExcelFile(archivo_fcm)
        hojas_fcm = xls_fcm.sheet_names
        indice_default = hojas_fcm.index("MANDAMIENTO") if "MANDAMIENTO" in hojas_fcm else 0

        hoja_seleccionada = st.selectbox(
            "Hoja a analizar en REPORTE FCM:",
            options=hojas_fcm,
            index=indice_default,
        )

        df_fcm = pd.read_excel(xls_fcm, sheet_name=hoja_seleccionada)
        df_sanciones = pd.read_excel(archivo_sanciones)
        df_cobro = pd.read_excel(archivo_cobro)

        # Validación de columnas requeridas
        if "Identificación" not in df_fcm.columns:
            st.error(f"La hoja '{hoja_seleccionada}' no contiene la columna 'Identificación'.")
            st.stop()
        if "Cédula" not in df_sanciones.columns:
            st.error("El archivo de Sanciones no contiene la columna 'Cédula'.")
            st.stop()
        if "Cédula" not in df_cobro.columns:
            st.error("El archivo de Cobro Coactivo no contiene la columna 'Cédula'.")
            st.stop()

        df_fcm["_cc_norm"] = normalizar_cc(df_fcm["Identificación"])
        cc_sanciones = set(normalizar_cc(df_sanciones["Cédula"])) - {""}
        cc_cobro = set(normalizar_cc(df_cobro["Cédula"])) - {""}
        cc_excluir = cc_sanciones | cc_cobro

        df_resultado = (
            df_fcm[~df_fcm["_cc_norm"].isin(cc_excluir)]
            .drop(columns=["_cc_norm"])
            .copy()
        )

        st.divider()
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total FCM (MANDAMIENTO)", len(df_fcm))
        m2.metric("Cédulas en Sanciones", len(cc_sanciones))
        m3.metric("Cédulas en Cobro Coactivo", len(cc_cobro))
        m4.metric("Exclusivos en FCM", f"{len(df_resultado)} filas ({df_resultado['Identificación'].nunique()} CC)")

        st.subheader("Vista previa de registros exclusivos en FCM")
        st.dataframe(df_resultado, use_container_width=True, hide_index=True)

        excel_bytes = generar_excel_bytes(df_resultado)
        st.download_button(
            label="Descargar resultado en Excel (.xlsx)",
            data=excel_bytes,
            file_name="fcm_no_en_pendientes.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

    except Exception as exc:
        st.error(f"Error procesando los archivos: {exc}")
else:
    st.info("Carga los tres archivos Excel arriba para ejecutar el cruce automáticamente.")

st.divider()
st.caption("Desarrollado por Roberto José Molina González — Ingeniería de Sistemas & Arquitectura de Software")