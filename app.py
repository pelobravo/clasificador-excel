import streamlit as st
import pandas as pd
from io import BytesIO

# =========================================================
# CONFIGURACIÓN GENERAL
# =========================================================

st.set_page_config(
    page_title="Clasificador Bancario - Grupo Bodeguita Oriente",
    page_icon="🏦",
    layout="wide"
)

# =========================================================
# ESTILOS
# =========================================================

st.markdown("""
<style>

.stApp {
    background-color: #ffffff;
}

.stButton > button {
    background-color: #1e3a5f;
    color: white;
    border-radius: 8px;
    padding: 10px 24px;
    font-weight: bold;
    border: none;
}

.stButton > button:hover {
    background-color: #2c5282;
}

h1, h2, h3 {
    color: #1e3a5f;
}

.footer {
    text-align: center;
    color: #666;
    padding: 20px;
    font-size: 14px;
}

</style>
""", unsafe_allow_html=True)

# =========================================================
# HEADER
# =========================================================

col_logo, col_titulo = st.columns([1, 5])

with col_logo:

    try:
        st.image("LOGO.jpeg", width=80)

    except Exception:

        st.image(
            "https://raw.githubusercontent.com/pelobravo/clasificador-excel/main/LOGO.jpeg",
            width=80
        )

with col_titulo:

    st.title("Clasificador Bancario")
    st.markdown("### Grupo Bodeguita Oriente")

st.markdown("---")

# =========================================================
# SIDEBAR
# =========================================================

with st.sidebar:

    st.image(
        "https://raw.githubusercontent.com/pelobravo/clasificador-excel/main/LOGO.jpeg",
        width=100
    )

    st.markdown("---")

    archivo = st.file_uploader(
        "📂 Cargar archivo Excel",
        type=["xlsx", "xls"]
    )

    st.markdown("---")

    procesar = st.button(
        "🚀 Procesar",
        type="primary",
        use_container_width=True
    )

# =========================================================
# FUNCIONES
# =========================================================

def convertir_monto(valor):

    try:

        if pd.isna(valor):
            return None

        if isinstance(valor, (int, float)):
            return float(valor)

        valor = str(valor).strip()

        valor = valor.replace(" ", "")
        valor = valor.replace("$", "")
        valor = valor.replace("Bs", "")
        valor = valor.replace("€", "")

        # 1.234,56
        if "." in valor and "," in valor:

            valor = valor.replace(".", "")
            valor = valor.replace(",", ".")

        elif "," in valor:

            valor = valor.replace(",", ".")

        return float(valor)

    except Exception:

        return None


def es_comision(texto):

    texto = str(texto).lower()

    palabras = [
        "comision",
        "comisión",
        "cargo bancario",
        "fee",
        "iva comisión"
    ]

    return any(p in texto for p in palabras)

# =========================================================
# PROCESAMIENTO MERCANTIL
# =========================================================

def procesar_archivo(df):

    ingresos = []
    egresos = []
    comisiones = []

    registros_procesados = set()

    for _, fila in df.iterrows():

        try:

            # =================================================
            # VALIDAR LONGITUD
            # =================================================

            if len(fila) < 8:
                continue

            # =================================================
            # COLUMNAS MERCANTIL
            # =================================================

            fecha = str(fila[3]).strip()

            tipo = str(fila[5]).strip().upper()

            descripcion = str(fila[6]).strip()

            monto = convertir_monto(fila[7])

            # =================================================
            # VALIDACIONES
            # =================================================

            if descripcion == "" or descripcion.lower() == "nan":
                continue

            if monto is None:
                continue

            if fecha == "" or fecha.lower() == "nan":
                continue

            # =================================================
            # EVITAR ENCABEZADOS
            # =================================================

            texto = descripcion.upper()

            palabras_invalidas = [
                "SALDO",
                "DESCRIPCION",
                "DESCRIPCIÓN",
                "REFERENCIA",
                "MOVIMIENTO",
                "FECHA"
            ]

            if texto in palabras_invalidas:
                continue

            # =================================================
            # REGISTRO
            # =================================================

            registro = {
                "FECHA": fecha,
                "DESCRIPCIÓN": descripcion,
                "MONTO": round(abs(monto), 2)
            }

            clave = (
                fecha,
                descripcion,
                monto,
                tipo
            )

            if clave in registros_procesados:
                continue

            registros_procesados.add(clave)

            # =================================================
            # COMISIONES
            # =================================================

            if es_comision(descripcion):

                comisiones.append(registro)

            # =================================================
            # INGRESOS
            # =================================================

            elif tipo == "NC":

                ingresos.append(registro)

            # =================================================
            # EGRESOS
            # =================================================

            elif tipo == "ND":

                egresos.append(registro)

        except Exception:
            pass

    return ingresos, egresos, comisiones

# =========================================================
# INTERFAZ PRINCIPAL
# =========================================================

if archivo:

    st.info(
        f"📄 Archivo: **{archivo.name}** "
        f"- {archivo.size/1024:.1f} KB"
    )

    try:

        # =====================================================
        # LEER EXCEL
        # =====================================================

        df_original = pd.read_excel(
            archivo,
            header=None
        )

        # =====================================================
        # PREVIEW
        # =====================================================

        with st.expander("👁️ Vista previa"):

            st.dataframe(
                df_original.head(20),
                use_container_width=True
            )

        # =====================================================
        # PROCESAR
        # =====================================================

        if procesar:

            with st.spinner("Procesando archivo..."):

                ingresos, egresos, comisiones = procesar_archivo(df_original)

            # =================================================
            # DATAFRAMES
            # =================================================

            df_ingresos = pd.DataFrame(ingresos)
            df_egresos = pd.DataFrame(egresos)
            df_comisiones = pd.DataFrame(comisiones)

            # =================================================
            # TOTALES
            # =================================================

            total_ingresos = (
                df_ingresos["MONTO"].sum()
                if not df_ingresos.empty else 0
            )

            total_egresos = (
                df_egresos["MONTO"].sum()
                if not df_egresos.empty else 0
            )

            total_comisiones = (
                df_comisiones["MONTO"].sum()
                if not df_comisiones.empty else 0
            )

            # =================================================
            # KPI
            # =================================================

            st.success(
                f"""
                ✅ Procesado correctamente |
                INGRESOS: {len(df_ingresos)} |
                EGRESOS: {len(df_egresos)} |
                COMISIONES: {len(df_comisiones)}
                """
            )

            col1, col2, col3 = st.columns(3)

            with col1:

                st.metric(
                    "💰 INGRESOS",
                    len(df_ingresos),
                    f"${total_ingresos:,.2f}"
                )

            with col2:

                st.metric(
                    "💸 EGRESOS",
                    len(df_egresos),
                    f"${total_egresos:,.2f}"
                )

            with col3:

                st.metric(
                    "💳 COMISIONES",
                    len(df_comisiones),
                    f"${total_comisiones:,.2f}"
                )

            # =================================================
            # RESULTADOS
            # =================================================

            st.subheader("📊 Resultados")

            tab1, tab2, tab3 = st.tabs([
                "📈 INGRESOS",
                "📉 EGRESOS",
                "💳 COMISIONES"
            ])

            # =================================================
            # INGRESOS
            # =================================================

            with tab1:

                if not df_ingresos.empty:

                    st.dataframe(
                        df_ingresos,
                        use_container_width=True,
                        height=400
                    )

                    st.success(
                        f"TOTAL INGRESOS: ${total_ingresos:,.2f}"
                    )

                else:

                    st.info("No se encontraron ingresos")

            # =================================================
            # EGRESOS
            # =================================================

            with tab2:

                if not df_egresos.empty:

                    st.dataframe(
                        df_egresos,
                        use_container_width=True,
                        height=400
                    )

                    st.error(
                        f"TOTAL EGRESOS: ${total_egresos:,.2f}"
                    )

                else:

                    st.info("No se encontraron egresos")

            # =================================================
            # COMISIONES
            # =================================================

            with tab3:

                if not df_comisiones.empty:

                    st.dataframe(
                        df_comisiones,
                        use_container_width=True,
                        height=400
                    )

                    st.warning(
                        f"TOTAL COMISIONES: ${total_comisiones:,.2f}"
                    )

                else:

                    st.info("No se encontraron comisiones")

            # =================================================
            # EXPORTAR EXCEL
            # =================================================

            output = BytesIO()

            with pd.ExcelWriter(
                output,
                engine="openpyxl"
            ) as writer:

                # INGRESOS
                if not df_ingresos.empty:

                    df_ingresos.to_excel(
                        writer,
                        sheet_name="INGRESOS",
                        index=False
                    )

                # EGRESOS
                if not df_egresos.empty:

                    df_egresos.to_excel(
                        writer,
                        sheet_name="EGRESOS",
                        index=False
                    )

                # COMISIONES
                if not df_comisiones.empty:

                    df_comisiones.to_excel(
                        writer,
                        sheet_name="COMISIONES",
                        index=False
                    )

                # RESUMEN
                resumen = pd.DataFrame([
                    {
                        "TIPO": "INGRESOS",
                        "CANTIDAD": len(df_ingresos),
                        "TOTAL": total_ingresos
                    },
                    {
                        "TIPO": "EGRESOS",
                        "CANTIDAD": len(df_egresos),
                        "TOTAL": total_egresos
                    },
                    {
                        "TIPO": "COMISIONES",
                        "CANTIDAD": len(df_comisiones),
                        "TOTAL": total_comisiones
                    }
                ])

                resumen.to_excel(
                    writer,
                    sheet_name="RESUMEN",
                    index=False
                )

            output.seek(0)

            # =================================================
            # DESCARGA
            # =================================================

            st.download_button(
                label="📥 Descargar Excel Clasificado",
                data=output.getvalue(),
                file_name=f"balance_{archivo.name}",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

    except Exception as e:

        st.error(f"❌ Error: {str(e)}")

# =========================================================
# PANTALLA INICIAL
# =========================================================

else:

    st.markdown("""

    ### 👋 Clasificador Bancario Inteligente

    ## FUNCIONES

    ✅ Clasifica automáticamente:
    - Ingresos
    - Egresos
    - Comisiones

    ✅ Genera:
    - KPIs
    - Tablas separadas
    - Totales automáticos
    - Exportación Excel

    ---

    ## INSTRUCCIONES

    1. Carga el Excel Mercantil
    2. Presiona PROCESAR
    3. Descarga el resultado

    """)

# =========================================================
# FOOTER
# =========================================================

st.markdown("---")

st.markdown(
    """
    <div class="footer">
        <strong>Grupo Bodeguita Oriente</strong>
        <br>
        Clasificador Bancario v15.0
    </div>
    """,
    unsafe_allow_html=True
)
