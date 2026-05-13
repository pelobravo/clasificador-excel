import streamlit as st
import pandas as pd
from io import BytesIO

from openpyxl.styles import Font
from openpyxl.styles import PatternFill
from openpyxl.styles import Border
from openpyxl.styles import Side
from openpyxl.styles import Alignment
from openpyxl.drawing.image import Image

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

        # FORMATO 1.234,56
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

            if len(fila) < 8:
                continue

            fecha = str(fila[3]).strip()

            tipo = str(fila[5]).strip().upper()

            descripcion = str(fila[6]).strip()

            monto = convertir_monto(fila[7])

            # VALIDACIONES

            if descripcion == "" or descripcion.lower() == "nan":
                continue

            if monto is None:
                continue

            if fecha == "" or fecha.lower() == "nan":
                continue

            # EVITAR ENCABEZADOS

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

            # REGISTRO

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

            # CLASIFICACIÓN

            if es_comision(descripcion):

                comisiones.append(registro)

            elif tipo == "NC":

                ingresos.append(registro)

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
            # KPIs
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
            # EXPORTAR EXCEL PROFESIONAL
            # =================================================

            output = BytesIO()

            with pd.ExcelWriter(
                output,
                engine="openpyxl"
            ) as writer:

                workbook = writer.book

                # =================================================
                # CREAR HOJA PRINCIPAL
                # =================================================

                hoja = workbook.create_sheet(
                    title="REPORTE"
                )

                # ELIMINAR HOJA VACÍA
                if "Sheet" in workbook.sheetnames:

                    hoja_vacia = workbook["Sheet"]
                    workbook.remove(hoja_vacia)

                # =================================================
                # ESTILOS
                # =================================================

                rojo = PatternFill(
                    start_color="FF0000",
                    end_color="FF0000",
                    fill_type="solid"
                )

                amarillo = PatternFill(
                    start_color="FFF2CC",
                    end_color="FFF2CC",
                    fill_type="solid"
                )

                verde = PatternFill(
                    start_color="C6E0B4",
                    end_color="C6E0B4",
                    fill_type="solid"
                )

                blanco = Font(
                    color="FFFFFF",
                    bold=True
                )

                negro_bold = Font(
                    bold=True
                )

                borde = Border(
                    left=Side(style="thin"),
                    right=Side(style="thin"),
                    top=Side(style="thin"),
                    bottom=Side(style="thin")
                )

                centro = Alignment(
                    horizontal="center",
                    vertical="center"
                )

                # =================================================
                # LOGO
                # =================================================

                try:

                    logo = Image("LOGO.jpeg")
                    logo.width = 130
                    logo.height = 130

                    hoja.add_image(logo, "A1")

                except Exception:
                    pass

                # =================================================
                # TITULO
                # =================================================

                hoja.merge_cells("C7:G7")

                hoja["C7"] = "BANCO MERCANTIL II BOD ANZOATEGUI"

                hoja["C7"].font = Font(
                    bold=True,
                    size=16
                )

                hoja["C7"].alignment = centro

                # =================================================
                # ENCABEZADO SUPERIOR
                # =================================================

                hoja["D1"] = "FECHA"
                hoja["E1"] = pd.Timestamp.now().strftime("%d/%m/%Y")

                hoja["D2"] = "SALDO INICIAL"
                hoja["D3"] = "SALDO FINAL"
                hoja["D4"] = "TASA DEL DIA"

                hoja["E2"] = total_ingresos
                hoja["E3"] = total_egresos
                hoja["E4"] = 36.50

                hoja["E4"].font = Font(
                    bold=True,
                    color="FF0000",
                    size=14
                )

                # =================================================
                # FUNCION TABLAS
                # =================================================

                def crear_tabla(
                    titulo,
                    dataframe,
                    fila_inicio,
                    color_total
                ):

                    hoja.merge_cells(
                        start_row=fila_inicio,
                        start_column=1,
                        end_row=fila_inicio,
                        end_column=5
                    )

                    titulo_cell = hoja.cell(
                        row=fila_inicio,
                        column=1
                    )

                    titulo_cell.value = titulo
                    titulo_cell.fill = rojo
                    titulo_cell.font = blanco
                    titulo_cell.alignment = centro

                    headers = [
                        "FECHA",
                        "DESCRIPCIÓN",
                        "MONTO",
                        "STATUS",
                        "OBSERVACIÓN"
                    ]

                    fila_header = fila_inicio + 1

                    for col_num, header in enumerate(headers, 1):

                        cell = hoja.cell(
                            row=fila_header,
                            column=col_num
                        )

                        cell.value = header
                        cell.fill = rojo
                        cell.font = blanco
                        cell.border = borde
                        cell.alignment = centro

                    fila_data = fila_header + 1

                    for _, row in dataframe.iterrows():

                        hoja.cell(
                            row=fila_data,
                            column=1
                        ).value = row["FECHA"]

                        hoja.cell(
                            row=fila_data,
                            column=2
                        ).value = row["DESCRIPCIÓN"]

                        hoja.cell(
                            row=fila_data,
                            column=3
                        ).value = row["MONTO"]

                        hoja.cell(
                            row=fila_data,
                            column=3
                        ).number_format = '$#,##0.00'

                        for col in range(1, 6):

                            hoja.cell(
                                row=fila_data,
                                column=col
                            ).border = borde

                        fila_data += 1

                    # =================================================
                    # TOTAL
                    # =================================================

                    hoja.merge_cells(
                        start_row=fila_data,
                        start_column=1,
                        end_row=fila_data,
                        end_column=2
                    )

                    total_label = hoja.cell(
                        row=fila_data,
                        column=1
                    )

                    total_label.value = f"TOTAL {titulo}"

                    total_label.font = negro_bold
                    total_label.alignment = centro

                    total_monto = hoja.cell(
                        row=fila_data,
                        column=3
                    )

                    total_monto.value = dataframe["MONTO"].sum()

                    total_monto.number_format = '$#,##0.00'
                    total_monto.fill = color_total
                    total_monto.font = negro_bold

                    return fila_data + 4

                # =================================================
                # CREAR TABLAS
                # =================================================

                fila_actual = 10

                fila_actual = crear_tabla(
                    "INGRESOS",
                    df_ingresos,
                    fila_actual,
                    verde
                )

                fila_actual = crear_tabla(
                    "EGRESOS",
                    df_egresos,
                    fila_actual,
                    amarillo
                )

                fila_actual = crear_tabla(
                    "COMISIONES",
                    df_comisiones,
                    fila_actual,
                    amarillo
                )

                # =================================================
                # ANCHOS COLUMNAS
                # =================================================

                hoja.column_dimensions["A"].width = 18
                hoja.column_dimensions["B"].width = 60
                hoja.column_dimensions["C"].width = 18
                hoja.column_dimensions["D"].width = 20
                hoja.column_dimensions["E"].width = 35

                # =================================================
                # GUARDAR ARCHIVO
                # =================================================

                workbook.save(output)

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
    - Exportación profesional

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
        Clasificador Bancario v17.0
    </div>
    """,
    unsafe_allow_html=True
)
