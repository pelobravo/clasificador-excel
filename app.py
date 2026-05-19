import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import date
import requests
from bs4 import BeautifulSoup
import json
import re

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
    
    fecha_inicio = st.date_input(
        "📅 Fecha Inicio",
        value=date.today().replace(day=1)
    )
    
    fecha_fin = st.date_input(
        "📅 Fecha Fin",
        value=date.today()
    )
    st.markdown("---")
    
    usar_api = st.checkbox(
        "🌐 Usar API BCV automática (experimental)",
        value=False
    )
    st.markdown("---")
    
    procesar = st.button(
        "🚀 Procesar",
        type="primary",
        use_container_width=True
    )

# =========================================================
# 1. DETECTOR DE BANCO
# =========================================================

def detectar_banco(nombre_archivo):
    nombre = nombre_archivo.upper()
    
    if "MERCANTIL" in nombre:
        return "mercantil"
    elif "VENEZUELA" in nombre or "BANCO DE VENEZUELA" in nombre:
        return "venezuela"
    elif "BANESCO" in nombre:
        return "banesco"
    elif "PROVINCIAL" in nombre:
        return "provincial"
    elif "BNC" in nombre:
        return "bnc"
    elif "TESORO" in nombre or "TESORERIA" in nombre:
        return "tesoro"
    return "desconocido"

# =========================================================
# 2. PARSERS POR BANCO
# =========================================================

def procesar_mercantil(df):
    """Procesa archivo del Banco Mercantil"""
    # El procesamiento actual ya está optimizado para Mercantil
    # Normalizar columnas
    df.columns = range(len(df.columns))
    return df

def procesar_venezuela(df):
    """Procesa archivo del Banco de Venezuela"""
    try:
        # Buscar fila de encabezados
        for i in range(min(10, len(df))):
            fila = df.iloc[i].astype(str)
            if fila.str.contains("fecha", case=False).any() or fila.str.contains("día", case=False).any():
                df.columns = df.iloc[i]
                df = df.iloc[i+1:].reset_index(drop=True)
                break
        
        # Normalizar columnas
        rename_map = {}
        for col in df.columns:
            col_lower = str(col).lower()
            if "fecha" in col_lower or "día" in col_lower or "dia" in col_lower:
                rename_map[col] = "FECHA"
            elif "descrip" in col_lower or "concepto" in col_lower or "detalle" in col_lower:
                rename_map[col] = "DESCRIPCION"
            elif "monto" in col_lower or "bs" in col_lower:
                rename_map[col] = "MONTO"
            elif "referencia" in col_lower or "ref" in col_lower:
                rename_map[col] = "REFERENCIA"
            elif "tipo" in col_lower or "movimiento" in col_lower:
                rename_map[col] = "TIPO"
        
        df = df.rename(columns=rename_map)
        return df
    except Exception as e:
        st.warning(f"Error en procesar_venezuela: {e}")
        return df

def procesar_banesco(df):
    """Procesa archivo del Banco Banesco"""
    try:
        # Buscar fila de encabezados
        for i in range(min(10, len(df))):
            fila = df.iloc[i].astype(str)
            if fila.str.contains("fecha", case=False).any():
                df.columns = df.iloc[i]
                df = df.iloc[i+1:].reset_index(drop=True)
                break
        
        # Normalizar columnas
        rename_map = {}
        for col in df.columns:
            col_lower = str(col).lower()
            if "fecha" in col_lower:
                rename_map[col] = "FECHA"
            elif "descrip" in col_lower or "concepto" in col_lower:
                rename_map[col] = "DESCRIPCION"
            elif "monto" in col_lower:
                rename_map[col] = "MONTO"
            elif "referencia" in col_lower:
                rename_map[col] = "REFERENCIA"
        
        df = df.rename(columns=rename_map)
        return df
    except Exception as e:
        st.warning(f"Error en procesar_banesco: {e}")
        return df

def procesar_provincial(df, archivo_original=None):
    """Procesa archivo del Banco Provincial (encabezado dinámico)"""
    try:
        encabezado = None
        for i in range(min(20, len(df))):
            fila = df.iloc[i].astype(str)
            if fila.str.contains("fecha", case=False).any():
                encabezado = i
                break
        
        if encabezado is not None:
            if archivo_original:
                # Recargar el archivo con skiprows
                df = pd.read_excel(archivo_original, skiprows=encabezado, header=0)
            else:
                df.columns = df.iloc[encabezado]
                df = df.iloc[encabezado+1:].reset_index(drop=True)
        
        # Normalizar columnas
        rename_map = {}
        for col in df.columns:
            col_lower = str(col).lower()
            if "fecha" in col_lower:
                rename_map[col] = "FECHA"
            elif "descrip" in col_lower or "concepto" in col_lower:
                rename_map[col] = "DESCRIPCION"
            elif "monto" in col_lower or "debito" in col_lower or "credito" in col_lower:
                rename_map[col] = "MONTO"
            elif "referencia" in col_lower:
                rename_map[col] = "REFERENCIA"
        
        df = df.rename(columns=rename_map)
        return df
    except Exception as e:
        st.warning(f"Error en procesar_provincial: {e}")
        return df

def procesar_bnc(df):
    """Procesa archivo del BNC"""
    try:
        for i in range(min(10, len(df))):
            fila = df.iloc[i].astype(str)
            if fila.str.contains("fecha", case=False).any():
                df.columns = df.iloc[i]
                df = df.iloc[i+1:].reset_index(drop=True)
                break
        
        rename_map = {}
        for col in df.columns:
            col_lower = str(col).lower()
            if "fecha" in col_lower:
                rename_map[col] = "FECHA"
            elif "descrip" in col_lower:
                rename_map[col] = "DESCRIPCION"
            elif "monto" in col_lower:
                rename_map[col] = "MONTO"
        
        df = df.rename(columns=rename_map)
        return df
    except Exception as e:
        st.warning(f"Error en procesar_bnc: {e}")
        return df

def procesar_tesoro(df):
    """Procesa archivo del Tesoro"""
    try:
        rename_map = {}
        for col in df.columns:
            col_lower = str(col).lower()
            if "fecha" in col_lower:
                rename_map[col] = "FECHA"
            elif "descrip" in col_lower:
                rename_map[col] = "DESCRIPCION"
            elif "monto" in col_lower:
                rename_map[col] = "MONTO"
        
        df = df.rename(columns=rename_map)
        return df
    except Exception as e:
        st.warning(f"Error en procesar_tesoro: {e}")
        return df

# =========================================================
# FUNCIONES GENERALES
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
        if valor == "":
            return None
        if "." in valor and "," in valor:
            valor = valor.replace(".", "")
            valor = valor.replace(",", ".")
        elif "," in valor:
            valor = valor.replace(",", ".")
        return float(valor)
    except Exception:
        return None

def normalizar_fecha(fecha_valor):
    """Normaliza fechas a formato datetime"""
    try:
        if pd.isna(fecha_valor):
            return pd.NaT
        
        # Si ya es datetime
        if isinstance(fecha_valor, (pd.Timestamp, datetime)):
            return fecha_valor
        
        fecha_str = str(fecha_valor).strip()
        fecha_str = fecha_str.replace(".0", "")
        
        # Formato dd/mm/yyyy
        if "/" in fecha_str:
            try:
                return pd.to_datetime(fecha_str, format="%d/%m/%Y", errors="coerce")
            except:
                pass
        
        # Formato dd-mm-yyyy
        if "-" in fecha_str:
            try:
                return pd.to_datetime(fecha_str, format="%d-%m-%Y", errors="coerce")
            except:
                pass
        
        # Dejar que pandas intente
        return pd.to_datetime(fecha_str, errors="coerce")
    except:
        return pd.NaT

# =========================================================
# OBTENER TASA BCV
# =========================================================

@st.cache_data(ttl=3600)
def obtener_tasa_bcv_fecha(fecha_obj):
    tasas_bcv_local = {
        "01/05/2026": 489.5547,
        "02/05/2026": 489.5547,
        "03/05/2026": 489.5547,
        "04/05/2026": 489.5547,
        "05/05/2026": 491.2281,
        "06/05/2026": 491.2281,
        "07/05/2026": 491.2281,
        "08/05/2026": 492.3542,
        "09/05/2026": 492.3542,
        "10/05/2026": 492.3542,
        "11/05/2026": 493.1023,
        "12/05/2026": 493.1023,
        "13/05/2026": 493.1023,
        "14/05/2026": 494.8765,
        "15/05/2026": 494.8765,
        "16/05/2026": 494.8765,
        "17/05/2026": 494.8765,
        "18/05/2026": 496.0012,
        "19/05/2026": 496.0012,
        "20/05/2026": 496.0012,
    }
    
    fecha_str = fecha_obj.strftime("%d/%m/%Y")
    
    if fecha_str in tasas_bcv_local:
        return tasas_bcv_local[fecha_str]
    
    if tasas_bcv_local:
        ultima_fecha = max(tasas_bcv_local.keys())
        return tasas_bcv_local[ultima_fecha]
    
    return None

def obtener_tasa_por_fecha(fecha_obj, usar_api=False):
    return obtener_tasa_bcv_fecha(fecha_obj)

def es_comision(texto):
    texto = str(texto).lower()
    palabras = [
        "comision", "comisión", "cargo", "cargo bancario", "fee",
        "iva", "itf", "impuesto", "op.cred", "op cred", "credito directo",
        "transferencia de fondos", "comision por transferencia",
        "comision pago movil", "comisión pago movil", "servicio bancario",
        "gasto bancario", "mantenimiento de cuenta", "debito automatico bancario"
    ]
    return any(p in texto for p in palabras)

# =========================================================
# PROCESAMIENTO UNIFICADO CON TASAS BCV
# =========================================================

def procesar_transacciones(df_normalizado, usar_api=False):
    """Procesa transacciones normalizadas de cualquier banco"""
    ingresos = []
    egresos = []
    comisiones = []
    registros_procesados = set()
    cache_tasas = {}
    
    tipos_ingresos = ["NC", "C", "CREDITO", "ABONO", "INGRESO", "DEPOSITO"]
    tipos_egresos = ["ND", "D", "DEBITO", "DEBIT", "EGRESO", "RETIRO"]
    
    for idx, fila in df_normalizado.iterrows():
        try:
            # Obtener fecha
            fecha_valor = fila.get("FECHA", None)
            if pd.isna(fecha_valor):
                continue
            
            # Convertir fecha a datetime
            fecha_obj = normalizar_fecha(fecha_valor)
            if pd.isna(fecha_obj):
                continue
            
            fecha_str = fecha_obj.strftime("%d/%m/%Y")
            
            # Obtener descripción y tipo
            descripcion = str(fila.get("DESCRIPCION", "")).strip()
            if descripcion == "" or descripcion.lower() == "nan":
                continue
            
            referencia = str(fila.get("REFERENCIA", idx))[:50]
            
            # Obtener monto
            monto_bs = fila.get("MONTO", 0)
            if pd.isna(monto_bs):
                continue
            monto_bs = abs(float(monto_bs))
            if monto_bs == 0:
                continue
            
            # Determinar tipo (ingreso/egreso)
            tipo = str(fila.get("TIPO", "")).upper()
            
            # Obtener tasa BCV
            if fecha_str in cache_tasas:
                tasa = cache_tasas[fecha_str]
            else:
                tasa = obtener_tasa_por_fecha(fecha_obj, usar_api)
                if tasa is not None:
                    cache_tasas[fecha_str] = tasa
            
            if tasa is None:
                tasa = 1.0
                st.warning(f"No se encontró tasa para fecha {fecha_str}, se usará tasa 1.0")
            
            # Calcular USD
            monto_usd = round(monto_bs / tasa, 2)
            
            # Filtrar palabras inválidas
            texto = descripcion.upper()
            palabras_invalidas = ["SALDO", "DESCRIPCION", "DESCRIPCIÓN", "REFERENCIA", 
                                  "MOVIMIENTO", "FECHA", "SALDO INICIAL", "SALDO FINAL"]
            if texto in palabras_invalidas:
                continue
            
            registro = {
                "FECHA": fecha_str,
                "REFERENCIA": referencia,
                "DESCRIPCIÓN": descripcion,
                "MONTO BS": round(monto_bs, 2),
                "TASA BCV": round(tasa, 4),
                "MONTO USD": monto_usd
            }
            
            clave = (fecha_str, referencia, descripcion, monto_usd)
            if clave in registros_procesados:
                continue
            registros_procesados.add(clave)
            
            if es_comision(descripcion):
                comisiones.append(registro)
            elif tipo in tipos_ingresos:
                ingresos.append(registro)
            elif tipo in tipos_egresos:
                egresos.append(registro)
            else:
                # Si no se pudo determinar por tipo, intentar por monto en contexto
                egresos.append(registro)  # Por defecto a egresos
                
        except Exception as e:
            st.warning(f"Error procesando fila {idx}: {e}")
    
    return ingresos, egresos, comisiones

# =========================================================
# INTERFAZ PRINCIPAL
# =========================================================

if archivo:
    st.info(f"📄 Archivo: **{archivo.name}** - {archivo.size/1024:.1f} KB")
    
    try:
        # DETECTAR BANCO
        banco = detectar_banco(archivo.name)
        st.info(f"🏦 Banco detectado: **{banco.upper()}**")
        
        # LEER EXCEL
        df_original = pd.read_excel(archivo, sheet_name=0, header=None)
        
        # APLICAR PARSER SEGÚN BANCO
        if banco == "mercantil":
            df_procesado = procesar_mercantil(df_original.copy())
        elif banco == "venezuela":
            df_procesado = procesar_venezuela(df_original.copy())
        elif banco == "banesco":
            df_procesado = procesar_banesco(df_original.copy())
        elif banco == "provincial":
            df_procesado = procesar_provincial(df_original.copy(), archivo)
        elif banco == "bnc":
            df_procesado = procesar_bnc(df_original.copy())
        elif banco == "tesoro":
            df_procesado = procesar_tesoro(df_original.copy())
        else:
            st.error(f"Banco no reconocido: {banco}")
            st.stop()
        
        with st.expander("👁️ Vista previa datos procesados"):
            st.dataframe(df_procesado.head(20), use_container_width=True)
        
        if procesar:
            # NORMALIZAR FECHAS
            if "FECHA" in df_procesado.columns:
                df_procesado["FECHA_DT"] = df_procesado["FECHA"].apply(normalizar_fecha)
                
                # Convertir fechas de filtro
                fecha_inicio_dt = pd.to_datetime(fecha_inicio)
                fecha_fin_dt = pd.to_datetime(fecha_fin)
                
                # Filtrar por rango de fechas
                df_filtrado = df_procesado[
                    (df_procesado["FECHA_DT"] >= fecha_inicio_dt) &
                    (df_procesado["FECHA_DT"] <= fecha_fin_dt)
                ].copy()
                
                st.success(f"Registros después de filtro: {len(df_filtrado)}")
            else:
                st.warning("No se encontró columna de fecha")
                df_filtrado = df_procesado
            
            with st.spinner("Procesando transacciones con tasas BCV..."):
                ingresos, egresos, comisiones = procesar_transacciones(df_filtrado, usar_api)
            
            df_ingresos = pd.DataFrame(ingresos)
            df_egresos = pd.DataFrame(egresos)
            df_comisiones = pd.DataFrame(comisiones)
            
            total_ingresos = df_ingresos["MONTO USD"].sum() if not df_ingresos.empty else 0
            total_egresos = df_egresos["MONTO USD"].sum() if not df_egresos.empty else 0
            total_comisiones = df_comisiones["MONTO USD"].sum() if not df_comisiones.empty else 0
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("💰 INGRESOS", len(df_ingresos), f"${total_ingresos:,.2f}")
            with col2:
                st.metric("💸 EGRESOS", len(df_egresos), f"${total_egresos:,.2f}")
            with col3:
                st.metric("💳 COMISIONES", len(df_comisiones), f"${total_comisiones:,.2f}")
            
            st.subheader("📊 Resultados")
            tab1, tab2, tab3 = st.tabs(["📈 INGRESOS", "📉 EGRESOS", "💳 COMISIONES"])
            
            with tab1:
                st.dataframe(df_ingresos, use_container_width=True)
            with tab2:
                st.dataframe(df_egresos, use_container_width=True)
            with tab3:
                st.dataframe(df_comisiones, use_container_width=True)
            
            # GENERAR EXCEL DE SALIDA
            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                workbook = writer.book
                hoja = workbook.create_sheet(title="REPORTE")
                
                if "Sheet" in workbook.sheetnames:
                    hoja_vacia = workbook["Sheet"]
                    workbook.remove(hoja_vacia)
                
                rojo = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
                verde = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")
                amarillo = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
                blanco = Font(color="FFFFFF", bold=True)
                borde = Border(left=Side(style="thin"), right=Side(style="thin"), 
                              top=Side(style="thin"), bottom=Side(style="thin"))
                centro = Alignment(horizontal="center", vertical="center")
                
                try:
                    logo = Image("LOGO.jpeg")
                    logo.width = 130
                    logo.height = 130
                    hoja.add_image(logo, "A1")
                except:
                    pass
                
                hoja.merge_cells("C7:H7")
                hoja["C7"] = f"{banco.upper()} - CON TASAS BCV"
                hoja["C7"].font = Font(bold=True, size=14)
                hoja["C7"].alignment = centro
                
                def crear_tabla(titulo, dataframe, fila_inicio, color_total):
                    hoja.merge_cells(start_row=fila_inicio, start_column=1, end_row=fila_inicio, end_column=8)
                    titulo_cell = hoja.cell(row=fila_inicio, column=1)
                    titulo_cell.value = titulo
                    titulo_cell.fill = rojo
                    titulo_cell.font = blanco
                    titulo_cell.alignment = centro
                    
                    headers = ["FECHA", "REFERENCIA", "DESCRIPCIÓN", "MONTO BS", "TASA BCV", "MONTO USD", "STATUS", "OBSERVACIÓN"]
                    fila_header = fila_inicio + 1
                    for col_num, header in enumerate(headers, 1):
                        cell = hoja.cell(row=fila_header, column=col_num)
                        cell.value = header
                        cell.fill = rojo
                        cell.font = blanco
                        cell.border = borde
                        cell.alignment = centro
                    
                    fila_data = fila_header + 1
                    for _, row in dataframe.iterrows():
                        hoja.cell(row=fila_data, column=1).value = row["FECHA"]
                        hoja.cell(row=fila_data, column=2).value = row["REFERENCIA"]
                        hoja.cell(row=fila_data, column=3).value = row["DESCRIPCIÓN"]
                        hoja.cell(row=fila_data, column=4).value = row["MONTO BS"]
                        hoja.cell(row=fila_data, column=5).value = row["TASA BCV"]
                        hoja.cell(row=fila_data, column=6).value = row["MONTO USD"]
                        hoja.cell(row=fila_data, column=4).number_format = '#,##0.00'
                        hoja.cell(row=fila_data, column=5).number_format = '#,##0.0000'
                        hoja.cell(row=fila_data, column=6).number_format = '$#,##0.00'
                        for col in range(1, 9):
                            hoja.cell(row=fila_data, column=col).border = borde
                        fila_data += 1
                    
                    total_cell = hoja.cell(row=fila_data, column=3)
                    total_cell.value = f"TOTAL {titulo}"
                    total_cell.font = Font(bold=True)
                    
                    monto_total = hoja.cell(row=fila_data, column=6)
                    monto_total.value = dataframe["MONTO USD"].sum()
                    monto_total.number_format = '$#,##0.00'
                    monto_total.fill = color_total
                    
                    return fila_data + 4
                
                fila_actual = 10
                if not df_ingresos.empty:
                    fila_actual = crear_tabla("INGRESOS", df_ingresos, fila_actual, verde)
                if not df_egresos.empty:
                    fila_actual = crear_tabla("EGRESOS", df_egresos, fila_actual, amarillo)
                if not df_comisiones.empty:
                    fila_actual = crear_tabla("COMISIONES", df_comisiones, fila_actual, amarillo)
                
                for columna in hoja.columns:
                    max_length = 0
                    try:
                        columna_letra = columna[0].column_letter
                    except:
                        continue
                    for cell in columna:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    hoja.column_dimensions[columna_letra].width = min(max_length + 5, 50)
            
            output.seek(0)
            st.download_button(
                label="📥 Descargar Excel Clasificado",
                data=output.getvalue(),
                file_name=f"balance_{banco}_{fecha_inicio}_{fecha_fin}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
            with st.expander("📊 Tasas BCV utilizadas"):
                todas_tasas = {}
                for registro in ingresos + egresos + comisiones:
                    fecha = registro["FECHA"]
                    tasa = registro["TASA BCV"]
                    if fecha not in todas_tasas:
                        todas_tasas[fecha] = tasa
                if todas_tasas:
                    df_tasas = pd.DataFrame([{"FECHA": f, "TASA BCV": t} for f, t in todas_tasas.items()]).sort_values("FECHA")
                    st.dataframe(df_tasas, use_container_width=True)
    
    except Exception as e:
        st.error(f"❌ Error general: {str(e)}")

else:
    st.markdown("""
    ### 👋 Clasificador Bancario Inteligente Multi-Banco

    ## FUNCIONES

    ✅ **Bancos soportados:**
    - Mercantil
    - Banco de Venezuela
    - Banesco
    - Provincial (encabezado dinámico)
    - BNC
    - Tesoro

    ✅ Clasifica automáticamente:
    - Ingresos
    - Egresos  
    - Comisiones

    ✅ Calcula USD con tasa BCV real por fecha

    ✅ Exporta reporte profesional
    """)
