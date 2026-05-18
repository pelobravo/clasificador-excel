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

        if "." in valor and "," in valor:

            valor = valor.replace(".", "")
            valor = valor.replace(",", ".")

        elif "," in valor:

            valor = valor.replace(",", ".")

        return float(valor)

    except Exception:

        return None

# =========================================================
# DETECTAR COMISIONES
# =========================================================

def es_comision(texto):

    texto = str(texto).lower()

    palabras = [

        "comision",
        "comisión",
        "cargo",
        "cargo bancario",
        "fee",
        "iva",
        "itf",
        "impuesto",

        "op.cred",
        "op cred",
        "credito directo",
        "transferencia de fondos",

        "comision por transferencia",
        "comision pago movil",
        "comisión pago movil",

        "servicio bancario",
        "gasto bancario",

        "mantenimiento de cuenta",
        "debito automatico bancario"
    ]

    return any(
        p in texto
        for p in palabras
    )

# =========================================================
# PROCESAMIENTO MERCANTIL
# =========================================================

def procesar_archivo(df):

    ingresos = []
    egresos = []
    comisiones = []

    registros_procesados = set()

    tipos_ingresos = [
        "NC",
        "C",
        "CREDITO",
        "ABONO"
    ]

    tipos_egresos = [
        "ND",
        "D",
        "DEBITO",
        "DEBIT"
    ]

    for _, fila in df.iterrows():

        try:

            if len(fila) < 9:
                continue

            # =================================================
            # COLUMNAS
            # =================================================

            fecha_raw = str(fila[3]).strip()

            if fecha_raw.lower() == "nan":
                continue

            fecha_raw = fecha_raw.replace(".0", "")

            if len(fecha_raw) == 7:

                dia = fecha_raw[0]
                mes = fecha_raw[1:3]
                anio = fecha_raw[3:]

                fecha = f"0{dia}/{mes}/{anio}"

            elif len(fecha_raw) == 8:

                dia = fecha_raw[0:2]
                mes = fecha_raw[2:4]
                anio = fecha_raw[4:]

                fecha = f"{dia}/{mes}/{anio}"

            else:

                fecha = fecha_raw

            tipo = str(
                fila[5]
            ).strip().upper()

            descripcion = str(
                fila[6]
            ).strip()

            referencia = str(
                fila[4]
            ).strip()

            monto_bs = convertir_monto(
                fila[7]
            )

            monto_usd = convertir_monto(
                fila[8]
            )

            # =================================================
            # VALIDACIONES
            # =================================================

            if descripcion == "" or descripcion.lower() == "nan":
                continue

            if monto_usd is None:
                continue

            texto = descripcion.upper()

            palabras_invalidas = [
                "SALDO",
                "DESCRIPCION",
                "DESCRIPCIÓN",
                "REFERENCIA",
                "MOVIMIENTO",
                "FECHA",
                "SALDO INICIAL",
                "SALDO FINAL"
            ]

            if texto in palabras_invalidas:
                continue

            # =================================================
            # REGISTRO
            # =================================================

            registro = {

                "FECHA": fecha,

                "REFERENCIA": referencia,

                "DESCRIPCIÓN": descripcion,

                "MONTO BS": round(
                    abs(monto_bs),
                    2
                ) if monto_bs else 0,

                "MONTO USD": round(
                    abs(monto_usd),
                    2
                ) if monto_usd else 0
            }

            clave = (
                fecha,
                referencia,
                descripcion,
                monto_usd,
                tipo
            )

            if clave in registros_procesados:
                continue

            registros_procesados.add(clave)

            # =================================================
            # CLASIFICACIÓN
            # =================================================

            if es_comision(descripcion):

                comisiones.append(registro)

            elif tipo in tipos_ingresos:

                ingresos.append(registro)

            elif tipo in tipos_egresos:

                egresos.append(registro)

        except Exception as e:

            st.warning(
                f"Error procesando fila: {e}"
            )

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
        # LEER SOLO LA PRIMERA HOJA
        # =====================================================

        df_original = pd.read_excel(
            archivo,
            sheet_name=0,
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

            df_ingresos = pd.DataFrame(ingresos)
            df_egresos = pd.DataFrame(egresos)
            df_comisiones = pd.DataFrame(comisiones)

            total_ingresos = (
                df_ingresos["MONTO USD"].sum()
                if not df_ingresos.empty else 0
            )

            total_egresos = (
                df_egresos["MONTO USD"].sum()
                if not df_egresos.empty else 0
            )

            total_comisiones = (
                df_comisiones["MONTO USD"].sum()
                if not df_comisiones.empty else 0
            )

            # =================================================
            # KPIs
            # =================================================

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

            with tab1:
                st.dataframe(df_ingresos, use_container_width=True)

            with tab2:
                st.dataframe(df_egresos, use_container_width=True)

            with tab3:
                st.dataframe(df_comisiones, use_container_width=True)

            # =================================================
            # EXPORTAR EXCEL
            # =================================================

            output = BytesIO()

            with pd.ExcelWriter(
                output,
                engine="openpyxl"
            ) as writer:

                workbook = writer.book

                hoja = workbook.create_sheet(
                    title="REPORTE"
                )

                if "Sheet" in workbook.sheetnames:

                    hoja_vacia = workbook["Sheet"]
                    workbook.remove(hoja_vacia)

                rojo = PatternFill(
                    start_color="FF0000",
                    end_color="FF0000",
                    fill_type="solid"
                )

                verde = PatternFill(
                    start_color="C6E0B4",
                    end_color="C6E0B4",
                    fill_type="solid"
                )

                amarillo = PatternFill(
                    start_color="FFF2CC",
                    end_color="FFF2CC",
                    fill_type="solid"
                )

                blanco = Font(
                    color="FFFFFF",
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

                except:
                    pass

                hoja.merge_cells("C7:H7")

                hoja["C7"] = "BANCO MERCANTIL II"

                hoja["C7"].font = Font(
                    bold=True,
                    size=16
                )

                hoja["C7"].alignment = centro

                # =================================================
                # FUNCIÓN TABLAS
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
                        end_column=7
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
                        "REFERENCIA",
                        "DESCRIPCIÓN",
                        "MONTO BS",
                        "MONTO USD",
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
                        ).value = row["REFERENCIA"]

                        hoja.cell(
                            row=fila_data,
                            column=3
                        ).value = row["DESCRIPCIÓN"]

                        hoja.cell(
                            row=fila_data,
                            column=4
                        ).value = row["MONTO BS"]

                        hoja.cell(
                            row=fila_data,
                            column=5
                        ).value = row["MONTO USD"]

                        hoja.cell(
                            row=fila_data,
                            column=4
                        ).number_format = '#,##0.00'

                        hoja.cell(
                            row=fila_data,
                            column=5
                        ).number_format = '$#,##0.00'

                        for col in range(1, 8):

                            hoja.cell(
                                row=fila_data,
                                column=col
                            ).border = borde

                        fila_data += 1

                    total_cell = hoja.cell(
                        row=fila_data,
                        column=3
                    )

                    total_cell.value = f"TOTAL {titulo}"

                    total_cell.font = Font(
                        bold=True
                    )

                    monto_total = hoja.cell(
                        row=fila_data,
                        column=5
                    )

                    monto_total.value = dataframe[
                        "MONTO USD"
                    ].sum()

                    monto_total.number_format = '$#,##0.00'
                    monto_total.fill = color_total

                    return fila_data + 4

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
                # AUTOAJUSTE COLUMNAS
                # =================================================

                for columna in hoja.columns:

                    max_length = 0

                    try:

                        columna_letra = (
                            columna[0].column_letter
                        )

                    except:
                        continue

                    for cell in columna:

                        try:

                            if len(str(cell.value)) > max_length:

                                max_length = len(
                                    str(cell.value)
                                )

                        except:
                            pass

                    adjusted_width = (
                        max_length + 5
                    )

                    hoja.column_dimensions[
                        columna_letra
                    ].width = adjusted_width

            output.seek(0)

            st.download_button(
                label="📥 Descargar Excel Clasificado",
                data=output.getvalue(),
                file_name=f"balance_{archivo.name}",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

    except Exception as e:

        st.error(f"❌ Error general: {str(e)}")

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
    - Exportación profesional

    """)
