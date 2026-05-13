import streamlit as st
import pandas as pd
import re
from io import BytesIO
from datetime import datetime

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
        type=['xlsx', 'xls']
    )

    st.markdown("---")

    procesar = st.button(
        "🚀 Procesar",
        type="primary",
        use_container_width=True
    )

# =========================================================
# FUNCIONES AUXILIARES
# =========================================================

def es_fecha_valida(valor):
    """
    Determina si un valor parece una fecha
    """

    if pd.isna(valor):
        return False

    # Si pandas ya detectó fecha
    if isinstance(valor, (datetime, pd.Timestamp)):
        return True

    valor_str = str(valor).strip()

    patrones = [
        r'\d{2}/\d{2}/\d{4}',
        r'\d{2}-\d{2}-\d{4}',
        r'\d{4}/\d{2}/\d{2}',
        r'\d{2}\.\d{2}\.\d{4}',
    ]

    for patron in patrones:
        if re.search(patron, valor_str):
            return True

    return False


def convertir_fecha(valor):

    try:

        if isinstance(valor, (datetime, pd.Timestamp)):
            return valor.strftime("%d/%m/%Y")

        return str(valor).strip()

    except Exception:
        return ""


def convertir_monto(valor):
    """
    Convierte montos correctamente:
    1.234,56
    1,234.56
    1234,56
    """

    try:

        if pd.isna(valor):
            return None

        if isinstance(valor, (int, float)):
            return float(valor)

        valor = str(valor).strip()

        # eliminar espacios
        valor = valor.replace(" ", "")

        # eliminar símbolos
        valor = valor.replace("$", "")
        valor = valor.replace("Bs", "")
        valor = valor.replace("€", "")

        # FORMATO EUROPEO
        # 1.234,56
        if "." in valor and "," in valor:
            valor = valor.replace(".", "")
            valor = valor.replace(",", ".")

        # FORMATO 1234,56
        elif "," in valor:
            valor = valor.replace(",", ".")

        return float(valor)

    except Exception:
        return None


def es_monto_valido(valor):

    monto = convertir_monto(valor)

    if monto is None:
        return False

    return 0.01 <= abs(monto) <= 999999999


def es_descripcion_valida(valor):

    if pd.isna(valor):
        return False

    valor_str = str(valor).strip()

    if len(valor_str) < 5:
        return False

    return bool(re.search(r'[A-Za-zÁÉÍÓÚáéíóúÑñ]', valor_str))


def detectar_tipo_movimiento(texto):

    texto = str(texto).upper()

    # INGRESOS
    patrones_ingreso = [
        r'\bNC\b',
        r'\bCR\b',
        r'\bCREDITO\b',
        r'\bCRÉDITO\b',
        r'\bABONO\b',
        r'\bINGRESO\b',
        r'\bDEPOSITO\b',
        r'\bDEPÓSITO\b',
    ]

    # EGRESOS
    patrones_egreso = [
        r'\bND\b',
        r'\bDB\b',
        r'\bDR\b',
        r'\bDEBITO\b',
        r'\bDÉBITO\b',
        r'\bCARGO\b',
        r'\bEGRESO\b',
        r'\bPAGO\b',
        r'\bRETIRO\b',
    ]

    for patron in patrones_ingreso:
        if re.search(patron, texto):
            return "INGRESO"

    for patron in patrones_egreso:
        if re.search(patron, texto):
            return "EGRESO"

    return ""


def es_comision(texto):

    texto = str(texto).lower()

    palabras = [
        "comision",
        "comisión",
        "comis",
        "fee",
        "cargo bancario"
    ]

    return any(p in texto for p in palabras)

# =========================================================
# PROCESAMIENTO PRINCIPAL
# =========================================================

def procesar_archivo(df):

    ingresos = []
    egresos = []
    comisiones = []

    registros_procesados = set()

    for fila in df.itertuples(index=False):

        fecha = ""
        descripcion = ""
        monto = None
        tipo_movimiento = ""

        fila_texto = " ".join([
            str(x) for x in fila if pd.notna(x)
        ])

        # =====================================================
        # RECORRER COLUMNAS
        # =====================================================

        for valor in fila:

            if pd.isna(valor):
                continue

            valor_str = str(valor).strip()

            # -------------------------------------------------
            # FECHA
            # -------------------------------------------------

            if not fecha and es_fecha_valida(valor):
                fecha = convertir_fecha(valor)

            # -------------------------------------------------
            # DESCRIPCIÓN
            # -------------------------------------------------

            if not descripcion and es_descripcion_valida(valor):
                descripcion = valor_str

            # -------------------------------------------------
            # MONTO
            # -------------------------------------------------

            if es_monto_valido(valor):

                monto_temp = convertir_monto(valor)

                if monto_temp is not None:

                    # Tomar el monto de mayor magnitud
                    if monto is None:
                        monto = monto_temp

                    elif abs(monto_temp) > abs(monto):
                        monto = monto_temp

            # -------------------------------------------------
            # TIPO
            # -------------------------------------------------

            tipo_detectado = detectar_tipo_movimiento(valor_str)

            if tipo_detectado:
                tipo_movimiento = tipo_detectado

        # =====================================================
        # VALIDACIONES
        # =====================================================

        if not fecha:
            continue

        if not descripcion:
            continue

        if monto is None:
            continue

        # evitar encabezados
        palabras_invalidas = [
            "saldo",
            "balance",
            "fecha",
            "descripcion",
            "descripción",
            "detalle",
            "movimiento"
        ]

        if descripcion.lower() in palabras_invalidas:
            continue

        # =====================================================
        # CLASIFICACIÓN AUTOMÁTICA
        # =====================================================

        if not tipo_movimiento:

            if monto < 0:
                tipo_movimiento = "EGRESO"
            else:
                tipo_movimiento = "INGRESO"

        # =====================================================
        # REGISTRO
        # =====================================================

        registro = {
            "FECHA": fecha,
            "DESCRIPCIÓN": descripcion,
            "MONTO": round(monto, 2)
        }

        # evitar duplicados
        clave = (
            fecha,
            descripcion,
            round(monto, 2)
        )

        if clave in registros_procesados:
            continue

        registros_procesados.add(clave)

        # =====================================================
        # COMISIONES
        # =====================================================

        if es_comision(descripcion):

            comisiones.append(registro)

        elif tipo_movimiento == "INGRESO":

            ingresos.append(registro)

        elif tipo_movimiento == "EGRESO":

            egresos.append(registro)

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
        # LEER ARCHIVO
        # =====================================================

        df_original = pd.read_excel(
            archivo,
            header=None
        )

        # =====================================================
        # PREVIEW
        # =====================================================

        with st.expander("👁️ Vista previa del archivo"):

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

            st.success(
                f"✅ Procesado correctamente | "
                f"INGRESOS: {len(ingresos)} | "
                f"EGRESOS: {len(egresos)} | "
                f"COMISIONES: {len(comisiones)}"
            )

            # =================================================
            # MÉTRICAS
            # =================================================

            total_ingresos = sum(
                x["MONTO"] for x in ingresos
            )

            total_egresos = sum(
                x["MONTO"] for x in egresos
            )

            total_comisiones = sum(
                x["MONTO"] for x in comisiones
            )

            col1, col2, col3 = st.columns(3)

            col1.metric(
                "💰 INGRESOS",
                len(ingresos),
                f"${total_ingresos:,.2f}"
            )

            col2.metric(
                "💸 EGRESOS",
                len(egresos),
                f"${total_egresos:,.2f}"
            )

            col3.metric(
                "💳 COMISIONES",
                len(comisiones),
                f"${total_comisiones:,.2f}"
            )

            # =================================================
            # EXPORTAR EXCEL
            # =================================================

            output = BytesIO()

            with pd.ExcelWriter(
                output,
                engine='openpyxl'
            ) as writer:

                # INGRESOS
                if ingresos:
                    pd.DataFrame(ingresos).to_excel(
                        writer,
                        sheet_name='INGRESOS',
                        index=False
                    )

                # EGRESOS
                if egresos:
                    pd.DataFrame(egresos).to_excel(
                        writer,
                        sheet_name='EGRESOS',
                        index=False
                    )

                # COMISIONES
                if comisiones:
                    pd.DataFrame(comisiones).to_excel(
                        writer,
                        sheet_name='COMISIONES',
                        index=False
                    )

                # RESUMEN
                resumen = pd.DataFrame([
                    {
                        "TIPO": "INGRESOS",
                        "CANTIDAD": len(ingresos),
                        "MONTO_TOTAL": total_ingresos
                    },
                    {
                        "TIPO": "EGRESOS",
                        "CANTIDAD": len(egresos),
                        "MONTO_TOTAL": total_egresos
                    },
                    {
                        "TIPO": "COMISIONES",
                        "CANTIDAD": len(comisiones),
                        "MONTO_TOTAL": total_comisiones
                    }
                ])

                resumen.to_excel(
                    writer,
                    sheet_name='RESUMEN',
                    index=False
                )

            output.seek(0)

            # =================================================
            # TABS
            # =================================================

            st.subheader("📊 Resultados")

            tab1, tab2, tab3 = st.tabs([
                "📈 INGRESOS",
                "📉 EGRESOS",
                "💳 COMISIONES"
            ])

            with tab1:

                if ingresos:

                    st.dataframe(
                        pd.DataFrame(ingresos),
                        use_container_width=True
                    )

                else:
                    st.info("No se encontraron ingresos")

            with tab2:

                if egresos:

                    st.dataframe(
                        pd.DataFrame(egresos),
                        use_container_width=True
                    )

                else:
                    st.info("No se encontraron egresos")

            with tab3:

                if comisiones:

                    st.dataframe(
                        pd.DataFrame(comisiones),
                        use_container_width=True
                    )

                else:
                    st.info("No se encontraron comisiones")

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

        st.error(f"❌ Error procesando archivo: {str(e)}")

# =========================================================
# PANTALLA INICIAL
# =========================================================

else:

    st.markdown("""
    ### 👋 Clasificador Bancario Inteligente

    ## FUNCIONES

    ✅ Detecta automáticamente:
    - Fechas
    - Descripciones
    - Ingresos
    - Egresos
    - Comisiones

    ✅ Compatible con:
    - Estados de cuenta
    - Bancos distintos
    - Formatos variados

    ✅ Exporta:
    - INGRESOS
    - EGRESOS
    - COMISIONES
    - RESUMEN

    ---

    ## INSTRUCCIONES

    1. Carga el archivo Excel
    2. Presiona "Procesar"
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
        Clasificador Bancario v13.0
    </div>
    """,
    unsafe_allow_html=True
)
