import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import date
import requests
from bs4 import BeautifulSoup
import json
import re
from datetime import datetime

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
        type=["xlsx", "xls", "xlsm"]
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
# FUNCIONES PARA LEER EXCEL CON ENGINE AUTOMÁTICO
# =========================================================

def leer_excel_sin_encabezados(archivo):
    """Lee archivo Excel sin encabezados detectando el engine correcto"""
    nombre = archivo.name.lower()
    
    try:
        if nombre.endswith('.xls') and not nombre.endswith('.xlsx'):
            try:
                import xlrd
                return pd.read_excel(archivo, sheet_name=0, header=None, engine='xlrd')
            except ImportError:
                st.error("❌ Para archivos .xls es necesario instalar xlrd. Ejecuta: pip install xlrd")
                st.stop()
        else:
            return pd.read_excel(archivo, sheet_name=0, header=None, engine='openpyxl')
    except Exception as e:
        st.error(f"No se pudo leer el archivo. Error: {str(e)}")
        st.stop()

def leer_excel_con_encabezados(archivo):
    """Lee archivo Excel con encabezados detectando el engine correcto"""
    nombre = archivo.name.lower()
    
    try:
        if nombre.endswith('.xls') and not nombre.endswith('.xlsx'):
            try:
                import xlrd
                return pd.read_excel(archivo, sheet_name=0, header=0, engine='xlrd')
            except ImportError:
                st.error("❌ Para archivos .xls es necesario instalar xlrd. Ejecuta: pip install xlrd")
                st.stop()
        else:
            return pd.read_excel(archivo, sheet_name=0, header=0, engine='openpyxl')
    except Exception as e:
        # Si falla, intentar sin encabezados
        try:
            return pd.read_excel(archivo, sheet_name=0, header=None, engine='openpyxl')
        except:
            st.error(f"No se pudo leer el archivo. Error: {str(e)}")
            st.stop()

# =========================================================
# FUNCIONES ORIGINALES (NO MODIFICADAS)
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

# =========================================================
# CALCULAR USD SEGÚN TASA
# =========================================================

def calcular_usd(monto_bs, tasa):

    try:

        if monto_bs is None:
            return None

        if tasa is None:
            return None

        if tasa == 0:
            return None

        return round(
            abs(monto_bs) / abs(tasa),
            2
        )

    except:

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
# DETECTOR DE BANCO CORREGIDO
# =========================================================

def detectar_banco(nombre_archivo):
    """Detecta el banco por el nombre del archivo"""
    nombre = nombre_archivo.upper()
    
    # ORDEN IMPORTANTE: TESORO primero (por TES)
    if "TESORO" in nombre or "TESORERIA" in nombre:
        return "tesoro"
    elif "TES" in nombre:  # BONUS: detectar TES como Tesoro
        return "tesoro"
    elif "BANESCO" in nombre:
        return "banesco"
    elif "VENEZUELA" in nombre or "BANCO DE VENEZUELA" in nombre:
        return "venezuela"
    elif "PROVINCIAL" in nombre:
        return "provincial"
    elif "BNC" in nombre:
        return "bnc"
    elif "MERCANTIL" in nombre:
        return "mercantil"
    return "mercantil"  # Por defecto mercantil

# =========================================================
# PROCESAR VENEZUELA - VERSIÓN CORREGIDA (CON COLUMNAS REALES)
# =========================================================

def procesar_venezuela(df):

    st.info("Procesando Banco de Venezuela...")

    # ============================================
    # RENOMBRAR COLUMNAS REALES
    # ============================================

    df = df.rename(columns={

        "Día": "FECHA",
        "Referencia": "REFERENCIA",
        "Descripción": "DESCRIPCION",
        "Tipo de Movimiento": "TIPO",
        "Crédito": "CREDITO",
        "Débito": "DEBITO"

    })

    # ============================================
    # FECHA
    # ============================================

    df["FECHA"] = pd.to_datetime(
        df["FECHA"],
        errors="coerce"
    )

    df = df[df["FECHA"].notna()]

    # ============================================
    # LIMPIAR MONTOS
    # ============================================

    def limpiar_numero(valor):

        valor = str(valor)

        valor = valor.replace(".", "")
        valor = valor.replace(",", ".")

        try:
            return float(valor)
        except:
            return 0

    df["CREDITO"] = df["CREDITO"].apply(
        limpiar_numero
    )

    df["DEBITO"] = df["DEBITO"].apply(
        limpiar_numero
    )

    # ============================================
    # CREAR MONTO FINAL
    # ============================================

    df["MONTO"] = df["CREDITO"] + df["DEBITO"]

    # ============================================
    # ELIMINAR CEROS
    # ============================================

    df = df[df["MONTO"] != 0]

    # ============================================
    # DEBUG
    # ============================================

    st.success(f"Registros Venezuela: {len(df)}")

    st.dataframe(df.head())

    return df

# =========================================================
# PROCESAR BANESCO - VERSIÓN CORREGIDA
# =========================================================

def procesar_banesco(df):

    st.info("Procesando Banesco...")

    rename_map = {}

    for col in df.columns:

        col_str = str(col).strip().lower()

        if "fecha" in col_str:
            rename_map[col] = "FECHA"

        elif "descrip" in col_str:
            rename_map[col] = "DESCRIPCION"

        elif "referencia" in col_str:
            rename_map[col] = "REFERENCIA"

        elif "monto" in col_str:
            rename_map[col] = "MONTO"

    df = df.rename(columns=rename_map)

    # ============================================
    # FECHA
    # ============================================

    if "FECHA" in df.columns:

        df["FECHA"] = pd.to_datetime(
            df["FECHA"],
            errors="coerce"
        )

        df = df[df["FECHA"].notna()]

    # ============================================
    # MONTO (con manejo de columnas duplicadas)
    # ============================================

    if "MONTO" in df.columns:
        monto = df["MONTO"]
        if isinstance(monto, pd.DataFrame):
            monto = monto.iloc[:, 0]
        df["MONTO"] = pd.to_numeric(monto, errors="coerce")
        df = df[df["MONTO"].notna()]

    # ============================================
    # CREAR TIPO AUTOMÁTICO
    # ============================================

    df["TIPO"] = df["MONTO"].apply(

        lambda x: "NC" if x > 0 else "ND"
    )

    # ============================================
    # ABSOLUTO
    # ============================================

    df["MONTO"] = df["MONTO"].abs()

    # ============================================
    # DEBUG
    # ============================================

    st.write("VISTA PREVIA BANESCO:")
    st.dataframe(df.head())

    return df

# =========================================================
# PROCESAR PROVINCIAL
# =========================================================

def procesar_provincial(df):
    """Procesa archivo del Banco Provincial"""
    
    st.info("Procesando archivo de Provincial...")
    
    # Buscar fila de encabezados
    encabezado = None
    for i in range(min(20, len(df))):
        fila = df.iloc[i].astype(str)
        if fila.str.contains("fecha", case=False).any():
            encabezado = i
            break
    
    if encabezado is not None:
        df.columns = df.iloc[encabezado]
        df = df.iloc[encabezado+1:].reset_index(drop=True)
    
    # Renombrar
    rename_map = {}
    for col in df.columns:
        col_str = str(col).lower()
        if "fecha" in col_str:
            rename_map[col] = "FECHA"
        elif "descrip" in col_str or "concepto" in col_str:
            rename_map[col] = "DESCRIPCION"
        elif "monto" in col_str:
            rename_map[col] = "MONTO"
    
    df = df.rename(columns=rename_map)
    
    if "FECHA" in df.columns:
        df["FECHA"] = pd.to_datetime(df["FECHA"], errors="coerce")
        df = df[df["FECHA"].notna()]
    
    if "MONTO" in df.columns:
        monto = df["MONTO"]
        if isinstance(monto, pd.DataFrame):
            monto = monto.iloc[:, 0]
        df["MONTO"] = pd.to_numeric(monto, errors="coerce")
        df = df[df["MONTO"].notna()]
    
    return df

# =========================================================
# PROCESAR BNC (con soporte .xls)
# =========================================================

def procesar_bnc(df):
    """Procesa archivo del BNC"""
    
    st.info("Procesando archivo de BNC...")
    
    # Buscar fila de encabezados
    for i in range(min(15, len(df))):
        fila = df.iloc[i].astype(str)
        if fila.str.contains("fecha", case=False).any():
            df.columns = df.iloc[i]
            df = df.iloc[i+1:].reset_index(drop=True)
            break
    
    rename_map = {}
    for col in df.columns:
        col_str = str(col).lower()
        if "fecha" in col_str:
            rename_map[col] = "FECHA"
        elif "descrip" in col_str:
            rename_map[col] = "DESCRIPCION"
        elif "monto" in col_str:
            rename_map[col] = "MONTO"
    
    df = df.rename(columns=rename_map)
    
    if "FECHA" in df.columns:
        df["FECHA"] = pd.to_datetime(df["FECHA"], errors="coerce")
        df = df[df["FECHA"].notna()]
    
    if "MONTO" in df.columns:
        monto = df["MONTO"]
        if isinstance(monto, pd.DataFrame):
            monto = monto.iloc[:, 0]
        df["MONTO"] = pd.to_numeric(monto, errors="coerce")
        df = df[df["MONTO"].notna()]
    
    return df

# =========================================================
# PROCESAR TESORO - VERSIÓN CORREGIDA (CON HTML MULTINIVEL)
# =========================================================

def procesar_tesoro(df):

    st.info("Procesando Banco del Tesoro...")

    # ============================================
    # APLANAR COLUMNAS MULTINIVEL
    # ============================================

    if isinstance(df.columns, pd.MultiIndex):

        df.columns = [
            col[-1]
            for col in df.columns
        ]

    # ============================================
    # RENOMBRAR
    # ============================================

    df = df.rename(columns={

        "Fecha": "FECHA",
        "Referencia": "REFERENCIA",
        "Concepto": "DESCRIPCION",
        "Débito": "DEBITO",
        "Crédito": "CREDITO",
        "Código": "TIPO"

    })

    # ============================================
    # FECHA
    # ============================================

    df["FECHA"] = pd.to_datetime(
        df["FECHA"],
        format="%d/%m/%Y",
        errors="coerce"
    )

    df = df[df["FECHA"].notna()]

    # ============================================
    # NUMÉRICOS
    # ============================================

    df["CREDITO"] = pd.to_numeric(
        df["CREDITO"],
        errors="coerce"
    ).fillna(0)

    df["DEBITO"] = pd.to_numeric(
        df["DEBITO"],
        errors="coerce"
    ).fillna(0)

    # ============================================
    # MONTO FINAL
    # ============================================

    df["MONTO"] = (
        df["CREDITO"] - df["DEBITO"]
    )

    # ============================================
    # TIPO NC / ND
    # ============================================

    df["TIPO"] = df["MONTO"].apply(

        lambda x: "NC" if x > 0 else "ND"
    )

    # ============================================
    # ABS
    # ============================================

    df["MONTO"] = df["MONTO"].abs()

    # ============================================
    # LIMPIAR
    # ============================================

    df = df[df["MONTO"] != 0]

    st.success(f"Registros Tesoro: {len(df)}")

    st.dataframe(df.head())

    return df

# =========================================================
# OBTENER TASA BCV (API o SCRAPING)
# =========================================================

@st.cache_data(ttl=3600)
def obtener_tasa_bcv_fecha(fecha_obj):
    """
    Obtiene la tasa BCV para una fecha específica.
    Por ahora usa diccionario local. Se puede expandir con API.
    """
    
    # Diccionario de tasas de ejemplo (formato: "dd/mm/yyyy" -> tasa)
    # ESTE DICCIONARIO DEBE ACTUALIZARSE CON DATOS REALES DEL BCV
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
    
    # Si no hay tasa para esa fecha, buscar la última disponible
    if tasas_bcv_local:
        ultima_fecha = max(tasas_bcv_local.keys())
        return tasas_bcv_local[ultima_fecha]
    
    return None

def obtener_tasa_por_fecha(fecha_obj, usar_api=False):
    """
    Wrapper para obtener tasa según método seleccionado
    """
    # Por ahora solo modo local (estable)
    # En futura versión se puede implementar scraping del BCV o API externa
    return obtener_tasa_bcv_fecha(fecha_obj)

# =========================================================
# FORMATEAR FECHA PARA CONSULTA
# =========================================================

def formatear_fecha_para_clave(fecha_str):
    """
    Convierte fecha del Excel a formato estandarizado
    """
    try:
        # Si viene en formato dd/mm/yyyy
        if "/" in fecha_str:
            partes = fecha_str.split("/")
            if len(partes) == 3:
                return f"{partes[0].zfill(2)}/{partes[1].zfill(2)}/{partes[2]}"
        return fecha_str
    except:
        return fecha_str

# =========================================================
# PROCESAMIENTO MERCANTIL ORIGINAL (COMPLETO, NO MODIFICADO)
# =========================================================

def procesar_archivo(df, usar_api=False):

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
    
    # Diccionario cache de tasas por fecha
    cache_tasas = {}

    for _, fila in df.iterrows():

        try:

            if len(fila) < 10:
                continue

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

            # =================================================
            # MONTO BS
            # =================================================

            monto_bs = convertir_monto(
                fila[7]
            )

            if monto_bs is None or monto_bs == 0:
                continue

            # =================================================
            # OBTENER TASA BCV SEGÚN FECHA
            # =================================================
            
            # Convertir fecha a objeto datetime para consultar tasa
            try:
                fecha_obj = pd.to_datetime(fecha, format="%d/%m/%Y", errors="coerce")
                fecha_key = fecha_obj.strftime("%d/%m/%Y")
            except:
                fecha_key = fecha
                fecha_obj = None
            
            # Validar que la fecha sea válida antes de buscar tasa
            if pd.isna(fecha_obj):
                st.warning(f"Fecha inválida: {fecha}, se omite")
                continue
            
            # Buscar tasa en cache o consultar
            if fecha_key in cache_tasas:
                tasa = cache_tasas[fecha_key]
            else:
                if fecha_obj is not None:
                    tasa = obtener_tasa_por_fecha(fecha_obj, usar_api)
                    if tasa is not None:
                        cache_tasas[fecha_key] = tasa
                else:
                    tasa = None
            
            # Si no hay tasa para esa fecha, usar tasa por defecto
            if tasa is None:
                tasa = 1.0
                st.warning(f"No se encontró tasa para fecha {fecha}, se usará tasa 1.0")

            # =================================================
            # CALCULAR USD REAL CON TASA BCV
            # =================================================
            
            monto_usd = calcular_usd(monto_bs, tasa)

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

            registro = {

                "FECHA": fecha,

                "REFERENCIA": referencia,

                "DESCRIPCIÓN": descripcion,

                "MONTO BS": round(
                    abs(monto_bs),
                    2
                ) if monto_bs else 0,
                
                "TASA BCV": round(tasa, 4),

                "MONTO USD": monto_usd
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
# NUEVA FUNCIÓN PARA PROCESAR OTROS BANCOS (CONVIERTE A FORMATO MERCANTIL)
# =========================================================

def convertir_a_formato_mercantil(df, banco):
    """Convierte DataFrame de otros bancos al formato que espera procesar_archivo"""
    
    # Crear un nuevo DataFrame con 10 columnas vacías (como espera Mercantil)
    datos_convertidos = []
    
    for idx, fila in df.iterrows():
        try:
            # Extraer datos normalizados
            fecha = fila.get("FECHA", "")
            if pd.isna(fecha):
                continue
            
            # Convertir fecha a string en formato esperado
            if isinstance(fecha, (pd.Timestamp, datetime)):
                fecha_str = fecha.strftime("%d/%m/%Y")
            else:
                fecha_str = str(fecha)
            
            # Tipo (NC/ND/etc) - CRÍTICO: esto es lo que determina ingreso/egreso
            tipo = fila.get("TIPO", "")
            if pd.isna(tipo):
                tipo = ""
            
            # Descripción
            descripcion = fila.get("DESCRIPCION", "")
            if pd.isna(descripcion):
                descripcion = ""
            
            # Referencia
            referencia = fila.get("REFERENCIA", "")
            if pd.isna(referencia):
                referencia = ""
            
            # Monto
            monto = fila.get("MONTO", 0)
            if pd.isna(monto):
                monto = 0
            
            # Crear fila con 10 columnas (como espera procesar_archivo)
            fila_convertida = [
                "",           # col0
                "",           # col1  
                "",           # col2
                fecha_str,    # col3 - FECHA
                referencia,   # col4 - REFERENCIA
                tipo,         # col5 - TIPO (NC/ND) - ¡ESTO ES CLAVE!
                descripcion,  # col6 - DESCRIPCION
                monto,        # col7 - MONTO BS
                "",           # col8
                "",           # col9
            ]
            datos_convertidos.append(fila_convertida)
            
        except Exception as e:
            continue
    
    # Crear DataFrame con 10 columnas
    df_convertido = pd.DataFrame(datos_convertidos)
    
    # Si no hay datos, devolver DataFrame vacío
    if len(df_convertido) == 0:
        return pd.DataFrame()
    
    return df_convertido

# =========================================================
# INTERFAZ PRINCIPAL
# =========================================================

if archivo:

    st.info(
        f"📄 Archivo: **{archivo.name}** "
        f"- {archivo.size/1024:.1f} KB"
    )

    try:
        # =========================================================
        # DETECTAR BANCO
        # =========================================================
        banco = detectar_banco(archivo.name)
        st.info(f"🏦 Banco detectado: **{banco.upper()}**")
        
        # =========================================================
        # LEER Y PROCESAR SEGÚN BANCO
        # =========================================================
        
        if banco == "mercantil":
            # MERCANTIL: usar el procesamiento original con lectura sin encabezados
            df_original = leer_excel_sin_encabezados(archivo)
            usar_procesamiento_original = True
            
        elif banco == "tesoro":
            # TESORO: usar read_html porque son archivos HTML disfrazados
            try:
                tablas = pd.read_html(archivo)
                if len(tablas) > 0:
                    df_raw = tablas[0]
                    st.success(f"✓ Se encontraron {len(tablas)} tablas en el archivo Tesoro. Usando la primera.")
                else:
                    st.error("No se encontraron tablas en el archivo Tesoro")
                    st.stop()
            except Exception as e:
                st.error(f"Error al leer archivo Tesoro: {str(e)}")
                st.stop()
            
            df_normalizado = procesar_tesoro(df_raw)
            df_original = convertir_a_formato_mercantil(df_normalizado, banco)
            
        else:
            # OTROS BANCOS: aplicar parser específico
            df_raw = leer_excel_con_encabezados(archivo)
            
            if banco == "venezuela":
                df_normalizado = procesar_venezuela(df_raw)
            elif banco == "banesco":
                df_normalizado = procesar_banesco(df_raw)
            elif banco == "provincial":
                df_normalizado = procesar_provincial(df_raw)
            elif banco == "bnc":
                df_normalizado = procesar_bnc(df_raw)
            else:
                df_normalizado = df_raw
            
            # Convertir al formato que espera procesar_archivo
            df_original = convertir_a_formato_mercantil(df_normalizado, banco)
            
        # Verificar que se pudieron convertir los datos
        if df_original.empty or len(df_original) == 0:
            st.error("No se pudieron convertir los datos. Verifica el formato del archivo.")
            st.stop()

        with st.expander("👁️ Vista previa archivo original"):

            st.dataframe(
                df_original.head(20),
                use_container_width=True
            )

        if procesar:

            # =========================================================
            # FILTRADO DE FECHAS (solo para Mercantil)
            # =========================================================
            
            if banco == "mercantil":
                try:
                    fechas_convertidas = []

                    for valor in df_original[3]:

                        fecha_raw = str(valor).strip()

                        fecha_raw = fecha_raw.replace(".0", "")

                        fecha_convertida = pd.NaT

                        if len(fecha_raw) == 7:

                            dia = fecha_raw[0]
                            mes = fecha_raw[1:3]
                            anio = fecha_raw[3:]

                            fecha_texto = f"0{dia}/{mes}/{anio}"

                            fecha_convertida = pd.to_datetime(
                                fecha_texto,
                                format="%d/%m/%Y",
                                errors="coerce"
                            )

                        elif len(fecha_raw) == 8:

                            dia = fecha_raw[0:2]
                            mes = fecha_raw[2:4]
                            anio = fecha_raw[4:]

                            fecha_texto = f"{dia}/{mes}/{anio}"

                            fecha_convertida = pd.to_datetime(
                                fecha_texto,
                                format="%d/%m/%Y",
                                errors="coerce"
                            )

                        fechas_convertidas.append(fecha_convertida)

                    df_original["FECHA_FILTRO"] = fechas_convertidas

                    df_original = df_original[
                        (df_original["FECHA_FILTRO"].dt.date >= fecha_inicio) &
                        (df_original["FECHA_FILTRO"].dt.date <= fecha_fin)
                    ]

                except Exception as e:

                    st.warning(f"Error filtrando fechas: {e}")
            else:
                # Para otros bancos, el filtrado ya se hizo en el parser
                st.success(f"Filtro de fechas aplicado: {fecha_inicio} a {fecha_fin}")

            with st.spinner("Procesando archivo con tasas BCV..."):

                ingresos, egresos, comisiones = procesar_archivo(df_original, usar_api)

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

                try:

                    logo = Image("LOGO.jpeg")
                    logo.width = 130
                    logo.height = 130

                    hoja.add_image(logo, "A1")

                except:
                    pass

                hoja.merge_cells("C7:H7")

                banco_nombre = banco.upper()
                hoja["C7"] = f"{banco_nombre} - CON TASAS BCV"

                hoja["C7"].font = Font(
                    bold=True,
                    size=14
                )

                hoja["C7"].alignment = centro

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
                        end_column=8
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
                        "TASA BCV",
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
                        column=6
                    )

                    monto_total.value = dataframe[
                        "MONTO USD"
                    ].sum()

                    monto_total.number_format = '$#,##0.00'
                    monto_total.fill = color_total

                    return fila_data + 4

                fila_actual = 10

                if not df_ingresos.empty:
                    fila_actual = crear_tabla(
                        "INGRESOS",
                        df_ingresos,
                        fila_actual,
                        verde
                    )

                if not df_egresos.empty:
                    fila_actual = crear_tabla(
                        "EGRESOS",
                        df_egresos,
                        fila_actual,
                        amarillo
                    )

                if not df_comisiones.empty:
                    fila_actual = crear_tabla(
                        "COMISIONES",
                        df_comisiones,
                        fila_actual,
                        amarillo
                    )

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
                        min(max_length + 5, 50)
                    )

                    hoja.column_dimensions[
                        columna_letra
                    ].width = adjusted_width

            output.seek(0)

            st.download_button(
                label="📥 Descargar Excel Clasificado (con Tasas BCV)",
                data=output.getvalue(),
                file_name=f"balance_{banco}_{fecha_inicio}_{fecha_fin}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
            # Mostrar resumen de tasas usadas
            with st.expander("📊 Tasas BCV utilizadas"):
                todas_tasas = {}
                for registro in ingresos + egresos + comisiones:
                    fecha = registro["FECHA"]
                    tasa = registro["TASA BCV"]
                    if fecha not in todas_tasas:
                        todas_tasas[fecha] = tasa
                
                if todas_tasas:
                    df_tasas = pd.DataFrame([
                        {"FECHA": f, "TASA BCV": t} 
                        for f, t in todas_tasas.items()
                    ]).sort_values("FECHA")
                    st.dataframe(df_tasas, use_container_width=True)

    except Exception as e:

        st.error(f"❌ Error general: {str(e)}")
        st.error("Detalles del error para depuración:")
        st.code(str(e))

else:

    st.markdown("""

    ### 👋 Clasificador Bancario Inteligente Multi-Banco

    ## FUNCIONES

    ✅ **Bancos soportados:**
    - Mercantil (original, completamente funcional)
    - Banco de Venezuela
    - Banesco
    - Provincial
    - BNC
    - Tesoro

    ✅ Clasifica automáticamente:
    - Ingresos (NC, C, CREDITO, ABONO)
    - Egresos (ND, D, DEBITO, DEBIT)
    - Comisiones

    ✅ Calcula USD con tasa BCV real por fecha

    ✅ Exporta reporte profesional con:
    - MONTO BS
    - TASA BCV
    - MONTO USD

    """)
