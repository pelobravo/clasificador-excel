import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import date
import requests
from bs4 import BeautifulSoup
import json
import re
from datetime import datetime
import unicodedata

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

# Inicialización de Estados para Evitar NameError y Permitir Acumulación
if "saldo_banesco" not in st.session_state: st.session_state.saldo_banesco = 0.0
if "saldo_bnc" not in st.session_state: st.session_state.saldo_bnc = 0.0
if "saldo_mercantil" not in st.session_state: st.session_state.saldo_mercantil = 0.0
if "saldo_venezuela" not in st.session_state: st.session_state.saldo_venezuela = 0.0
if "saldo_provincial" not in st.session_state: st.session_state.saldo_provincial = 0.0
if "saldo_bancamiga" not in st.session_state: st.session_state.saldo_bancamiga = 0.0
if "saldo_tesoro" not in st.session_state: st.session_state.saldo_tesoro = 0.0

# =========================================================
# ESTILOS
# =========================================================

st.markdown("""
<style>

/* Main Background and Fonts */
.stApp {
    background-color: #fafbfc;
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
}

/* Centered Titles and Headers */
h1, h2, h3, h4, h5, h6 {
    text-align: center;
    color: #1e3a5f;
    font-weight: 700;
}

/* Custom premium buttons with hover animations */
.stButton > button {
    background: linear-gradient(135deg, #1e3a5f 0%, #00a8cc 100%);
    color: white;
    border-radius: 8px;
    padding: 12px 28px;
    font-weight: bold;
    border: none;
    box-shadow: 0 4px 15px rgba(0, 168, 204, 0.2);
    transition: all 0.3s ease;
    width: 100%;
}

.stButton > button:hover {
    transform: translateY(-2px);
    box-shadow: 0 8px 25px rgba(0, 168, 204, 0.4);
    background: linear-gradient(135deg, #2c5282 0%, #00b5d8 100%);
}

.footer {
    text-align: center;
    color: #a0aec0;
    padding: 24px;
    font-size: 14px;
    border-top: 1px solid #e2e8f0;
    margin-top: 40px;
}

/* Premium KPIs Styles (Vivid Navy to Teal Gradient, Centered) */
.kpi-container {
    display: flex;
    gap: 24px;
    margin-bottom: 35px;
    flex-wrap: wrap;
    justify-content: center;
}
.kpi-card {
    background: linear-gradient(135deg, #1e3a5f 0%, #00b5d8 100%);
    color: white;
    padding: 24px;
    border-radius: 16px;
    box-shadow: 0 10px 30px rgba(0, 181, 216, 0.15);
    flex: 1;
    min-width: 290px;
    max-width: 380px;
    border: 1px solid rgba(255, 255, 255, 0.2);
    transition: all 0.3s cubic-bezier(0.25, 0.8, 0.25, 1);
    text-align: center;
}
.kpi-card:hover {
    transform: translateY(-5px);
    box-shadow: 0 15px 35px rgba(0, 181, 216, 0.3);
}
.kpi-title {
    font-size: 13px;
    text-transform: uppercase;
    letter-spacing: 1px;
    color: #e2e8f0;
    margin-bottom: 8px;
    font-weight: 600;
}
.kpi-value {
    font-size: 28px;
    font-weight: 800;
    color: #ffffff;
    text-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
}
.kpi-subtitle {
    font-size: 12px;
    color: #cbd5e0;
    margin-top: 6px;
    font-style: italic;
}

</style>
""", unsafe_allow_html=True)

# =========================================================
# 🔥 NUEVAS FUNCIONES PARA CONCILIACIÓN MULTIBANCO
# =========================================================

def formato_venezolano(valor):
    """Formatea un número al estilo de moneda venezolana (miles con punto, decimales con coma)"""
    if valor is None:
        return "0,00"
    try:
        parts = f"{float(valor):,.2f}".split(".")
        parts[0] = parts[0].replace(",", ".")
        return ",".join(parts)
    except:
        return "0,00"

def obtener_tasa_bcv():
    """Obtiene la tasa del día de forma segura"""
    tasa = obtener_tasa_por_fecha(date.today())
    if tasa is None:
        # Intentar obtener la tasa más reciente disponible en nuestro diccionario local
        tasa = 623.0223  # Fallback tasa del 30/06/2026
    return tasa

def obtener_saldo_final_banesco(df_raw):
    """Extrae el saldo final de la columna BALANCE del archivo de Banesco"""
    try:
        df_temp = df_raw.copy()
        if df_temp.shape[1] >= 5:
            balances = df_temp.iloc[:, 4].dropna()
            for val in reversed(balances.values):
                val_clean = convertir_monto(val)
                if val_clean is not None:
                    return val_clean
    except Exception as e:
        st.warning(f"No se pudo extraer el saldo final de Banesco: {e}")
    return 0.0

def obtener_saldo_final_bnc(df_raw, encabezado_idx):
    """Busca la columna de saldo en el reporte de BNC y extrae el último valor numérico"""
    try:
        df_temp = df_raw.iloc[encabezado_idx + 1:].copy()
        headers = df_raw.iloc[encabezado_idx].fillna("").astype(str).str.lower().tolist()
        saldo_col_idx = None
        for idx, h in enumerate(headers):
            if "saldo" in h or "balance" in h:
                saldo_col_idx = idx
                break
        if saldo_col_idx is not None:
            balances = df_temp.iloc[:, saldo_col_idx].dropna()
            for val in reversed(balances.values):
                val_clean = convertir_monto(val)
                if val_clean is not None:
                    return val_clean
    except Exception as e:
        st.warning(f"No se pudo extraer el saldo final de BNC: {e}")
    return 0.0

def obtener_saldo_final_mercantil(df_raw):
    """Extrae el saldo final de la columna BALANCE (columna 9) en Mercantil"""
    try:
        df_temp = df_raw.copy()
        if df_temp.shape[1] >= 9:
            balances = df_temp.iloc[:, 8].dropna()
            for val in reversed(balances.values):
                val_clean = convertir_monto(val)
                if val_clean is not None:
                    return val_clean
    except:
        pass
    return 0.0

def buscar_saldo_en_texto(df_raw):
    """Escanea todo el reporte en busca de celdas con la palabra 'SALDO' o 'DISPONIBLE' y obtiene el número"""
    try:
        # 1. Búsqueda específica de términos de saldo final
        terminos_finales = ["saldo disponible", "saldo actual", "saldo final", "saldo de la cuenta", "monto disponible"]
        for r_idx in range(df_raw.shape[0]):
            for c_idx in range(df_raw.shape[1]):
                val = str(df_raw.iloc[r_idx, c_idx]).lower()
                if any(term in val for term in terminos_finales):
                    # Eliminar fechas del texto antes de buscar números para evitar extraer el año (ej. 2026)
                    texto_sin_fechas = re.sub(r'\b\d{2}[/\-]\d{2}[/\-]\d{4}\b', '', val)
                    texto_sin_fechas = re.sub(r'\b\d{4}[/\-]\d{2}[/\-]\d{2}\b', '', texto_sin_fechas)
                    partes = re.findall(r'[\d\.\,]+', texto_sin_fechas)
                    if partes:
                        for p in reversed(partes):
                            num = convertir_monto(p)
                            if num is not None and num > 100:
                                return num
                    # Celda de la derecha
                    if c_idx + 1 < df_raw.shape[1]:
                        val_right = df_raw.iloc[r_idx, c_idx + 1]
                        num = convertir_monto(val_right)
                        if num is not None and num > 0:
                            return num
                    # Celda de abajo
                    if r_idx + 1 < df_raw.shape[0]:
                        val_below = df_raw.iloc[r_idx + 1, c_idx]
                        num = convertir_monto(val_below)
                        if num is not None and num > 0:
                            return num

        # 2. Búsqueda general si la específica falla (ignora inicial, anterior y promedio)
        for r_idx in range(df_raw.shape[0]):
            for c_idx in range(df_raw.shape[1]):
                val = str(df_raw.iloc[r_idx, c_idx]).lower()
                if ("saldo" in val or "disponible" in val) and "inicial" not in val and "anterior" not in val and "promedio" not in val:
                    # Eliminar fechas del texto antes de buscar números para evitar extraer el año (ej. 2026)
                    texto_sin_fechas = re.sub(r'\b\d{2}[/\-]\d{2}[/\-]\d{4}\b', '', val)
                    texto_sin_fechas = re.sub(r'\b\d{4}[/\-]\d{2}[/\-]\d{2}\b', '', texto_sin_fechas)
                    partes = re.findall(r'[\d\.\,]+', texto_sin_fechas)
                    if partes:
                        for p in reversed(partes):
                            num = convertir_monto(p)
                            if num is not None and num > 100:
                                return num
                    # Celda de la derecha
                    if c_idx + 1 < df_raw.shape[1]:
                        val_right = df_raw.iloc[r_idx, c_idx + 1]
                        num = convertir_monto(val_right)
                        if num is not None and num > 0:
                            return num
                    # Celda de abajo
                    if r_idx + 1 < df_raw.shape[0]:
                        val_below = df_raw.iloc[r_idx + 1, c_idx]
                        num = convertir_monto(val_below)
                        if num is not None and num > 0:
                            return num
    except:
        pass
    return 0.0

def buscar_saldo_inicial(df_raw):
    """Busca un saldo inicial en el texto (Saldo Inicial, Saldo Anterior, etc.)"""
    try:
        for r_idx in range(df_raw.shape[0]):
            for c_idx in range(df_raw.shape[1]):
                val = str(df_raw.iloc[r_idx, c_idx]).lower()
                if "saldo" in val and ("inicial" in val or "anterior" in val):
                    # Celda de la derecha
                    if c_idx + 1 < df_raw.shape[1]:
                        val_right = df_raw.iloc[r_idx, c_idx + 1]
                        num = convertir_monto(val_right)
                        if num is not None and num > 0:
                            return num
                    # Celda de abajo
                    if r_idx + 1 < df_raw.shape[0]:
                        val_below = df_raw.iloc[r_idx + 1, c_idx]
                        num = convertir_monto(val_below)
                        if num is not None and num > 0:
                            return num
    except:
        pass
    return 0.0

def obtener_saldo_final_tesoro(df_raw):
    """Calcula el saldo final de Banco del Tesoro sumando Créditos y Débitos o usando el neto"""
    # Intentar buscar por palabra 'Saldo' primero
    saldo_buscado = buscar_saldo_en_texto(df_raw)
    if saldo_buscado > 0:
        return saldo_buscado
        
    # Intentar buscar saldo inicial
    saldo_inicial = buscar_saldo_inicial(df_raw)
    
    # Si no tiene columna Saldo, sumamos todos los créditos y restamos débitos
    try:
        encabezado = None
        for i in range(min(30, len(df_raw))):
            fila = df_raw.iloc[i].fillna("").astype(str)
            texto = " ".join(fila.tolist()).lower()
            if "fecha" in texto and "referencia" in texto and "concepto" in texto:
                encabezado = i
                break
        if encabezado is None:
            return saldo_inicial
            
        df_temp = df_raw.iloc[encabezado + 1:].copy()
        headers = df_raw.iloc[encabezado].fillna("").astype(str).str.strip().tolist()
        
        col_debito_idx = None
        col_credito_idx = None
        for idx, col in enumerate(headers):
            c_clean = col.lower()
            if "débito" in c_clean or "debito" in c_clean:
                col_debito_idx = idx
            elif "crédito" in c_clean or "credito" in c_clean:
                col_credito_idx = idx
                
        total_creditos = 0.0
        total_debitos = 0.0
        
        if col_credito_idx is not None:
            creditos = df_temp.iloc[:, col_credito_idx].dropna()
            for val in creditos:
                num = convertir_monto(val)
                if num is not None:
                    total_creditos += num
                    
        if col_debito_idx is not None:
            debitos = df_temp.iloc[:, col_debito_idx].dropna()
            for val in debitos:
                num = convertir_monto(val)
                if num is not None:
                    total_debitos += num
                    
        # Retornar: Saldo Inicial + Créditos - Débitos
        saldo_calc = saldo_inicial + total_creditos - total_debitos
        # Si da negativo porque no hay saldo inicial en el archivo, tomar el neto absoluto de los movimientos
        if saldo_calc < 0 and saldo_inicial == 0.0:
            saldo_calc = abs(total_creditos - total_debitos)
        return saldo_calc
    except:
        pass
    return saldo_inicial

def obtener_saldo_final_columna_derecha(df_raw):
    """Busca en las últimas columnas del DataFrame (de derecha a izquierda) el último valor numérico válido"""
    try:
        df_temp = df_raw.copy()
        num_cols = df_temp.shape[1]
        for col_idx in range(num_cols - 1, 3, -1):
            balances = df_temp.iloc[:, col_idx].dropna()
            for val in reversed(balances.values):
                val_clean = convertir_monto(val)
                if val_clean is not None and val_clean > 0:
                    return val_clean
    except:
        pass
    return 0.0

def obtener_saldo_final_bancamiga(df_raw):
    """Busca la columna de saldo en Bancamiga y extrae el último valor numérico"""
    try:
        encabezado_idx = None
        for i in range(min(20, len(df_raw))):
            fila = df_raw.iloc[i].fillna("").astype(str)
            texto = " ".join(fila.tolist()).lower()
            if "fecha" in texto and "referencia" in texto and "saldo" in texto:
                encabezado_idx = i
                break
        if encabezado_idx is None:
            return buscar_saldo_en_texto(df_raw)
            
        df_temp = df_raw.iloc[encabezado_idx + 1:].copy()
        headers = df_raw.iloc[encabezado_idx].fillna("").astype(str).str.lower().tolist()
        saldo_col_idx = None
        for idx, h in enumerate(headers):
            if "saldo" in h:
                saldo_col_idx = idx
                break
        if saldo_col_idx is not None:
            balances = df_temp.iloc[:, saldo_col_idx].dropna()
            for val in reversed(balances.values):
                val_clean = convertir_monto(val)
                if val_clean is not None:
                    return val_clean
    except:
        pass
    return buscar_saldo_en_texto(df_raw)

def encontrar_fila_encabezado(df_raw):
    """Busca en las primeras filas una que contenga 'fecha' y ('descripcion' o 'descripción' o 'referencia')"""
    for i in range(min(40, len(df_raw))):
        fila = df_raw.iloc[i].fillna("").astype(str)
        texto = " ".join(fila.tolist()).lower()
        if "fecha" in texto and ("descri" in texto or "referencia" in texto or "monto" in texto):
            return i
    return None

def preparar_df_con_encabezado_dinamico(df_raw):
    """Encuentra el encabezado y limpia el DataFrame para que las columnas tengan los nombres correctos"""
    df_clean = df_raw.copy()
    idx_header = encontrar_fila_encabezado(df_clean)
    if idx_header is not None:
        cols = df_clean.iloc[idx_header].fillna("").astype(str).tolist()
        cols = [c.strip() for c in cols]
        df_clean.columns = cols
        df_clean = df_clean.iloc[idx_header + 1:].reset_index(drop=True)
    return df_clean

def obtener_saldo_banco(df_raw, banco, encabezado_idx=None):
    """Obtiene el saldo de un banco combinando extractores específicos y el escáner de texto"""
    if banco == "banesco":
        return obtener_saldo_final_banesco(df_raw) or obtener_saldo_final_columna_derecha(df_raw)
    elif banco == "bnc":
        if encabezado_idx is not None:
            return obtener_saldo_final_bnc(df_raw, encabezado_idx)
        return buscar_saldo_en_texto(df_raw) or obtener_saldo_final_columna_derecha(df_raw)
    elif banco == "mercantil":
        return obtener_saldo_final_mercantil(df_raw) or buscar_saldo_en_texto(df_raw) or obtener_saldo_final_columna_derecha(df_raw)
    elif banco == "tesoro":
        return obtener_saldo_final_tesoro(df_raw)
    elif banco == "bancamiga":
        return obtener_saldo_final_bancamiga(df_raw) or obtener_saldo_final_columna_derecha(df_raw)
    else:
        return buscar_saldo_en_texto(df_raw) or obtener_saldo_final_columna_derecha(df_raw)

def calcular_saldo_movimientos(df_convertido):
    """Suma los ingresos (NC) y resta los egresos (ND) de un DataFrame en formato unificado"""
    if df_convertido is None or df_convertido.empty:
        return 0.0
    try:
        tipo_col = "TIPO" if "TIPO" in df_convertido.columns else df_convertido.columns[5]
        monto_col = "MONTO BS" if "MONTO BS" in df_convertido.columns else df_convertido.columns[7]
        ingresado = df_convertido[df_convertido[tipo_col] == "NC"][monto_col].sum()
        egresado = df_convertido[df_convertido[tipo_col] == "ND"][monto_col].sum()
        return ingresado - egresado
    except:
        return 0.0

# =========================================================
# 🔥 FUNCIÓN PARA NORMALIZAR TEXTO (eliminar acentos)
# =========================================================

def normalizar_texto(texto):
    """Normaliza texto eliminando acentos y convirtiendo a minúsculas"""
    texto = str(texto)
    texto = (
        unicodedata.normalize("NFKD", texto)
        .encode("ascii", "ignore")
        .decode("utf-8")
    )
    return texto.lower()

# =========================================================
# HEADER
# =========================================================

col_l, col_c, col_r = st.columns([2, 1, 2])

with col_c:
    try:
        st.image("LOGO.jpeg", width=145)
    except Exception:
        st.image(
            "https://raw.githubusercontent.com/pelobravo/clasificador-excel/main/LOGO.jpeg",
            width=145
        )

st.markdown("<h1 style='text-align: center; color: #1e3a5f; margin-top: 15px; margin-bottom: 5px;'>Clasificador Bancario</h1>", unsafe_allow_html=True)
st.markdown("<h3 style='text-align: center; color: #00a8cc; margin-top: 0px; margin-bottom: 20px; font-weight: 600;'>Grupo Bodeguita Oriente</h3>", unsafe_allow_html=True)
st.markdown("---")

# Inicializar estado de navegación
if "seccion_activa" not in st.session_state:
    st.session_state.seccion_activa = "consolidado"

# Barra de Navegación Superior
col_nav1, col_nav2 = st.columns(2)
with col_nav1:
    if st.button("📊 CIERRE CONSOLIDADO MULTIBANCO", use_container_width=True, type="primary" if st.session_state.seccion_activa == "consolidado" else "secondary", key="nav_btn_consolidado"):
        st.session_state.seccion_activa = "consolidado"
        st.rerun()
with col_nav2:
    if st.button("🔍 CRUCE DE INFORMACIÓN (MONOBANCO)", use_container_width=True, type="primary" if st.session_state.seccion_activa == "cruce" else "secondary", key="nav_btn_cruce"):
        st.session_state.seccion_activa = "cruce"
        st.rerun()
st.markdown("---")

# =========================================================
# SIDEBAR DINÁMICO Y DECLARACIÓN DE VARIABLES
# =========================================================

# Declaración de variables para evitar NameError en cualquiera de los dos flujos
archivo_ipago = None
archivo_banesco = None
archivo_bnc = None
archivo_mercantil = None
archivo_venezuela = None
archivo_provincial = None
archivo_bancamiga = None
saldo_manual_tesoro = 0.0

archivo = None

with st.sidebar:
    st.image(
        "https://raw.githubusercontent.com/pelobravo/clasificador-excel/main/LOGO.jpeg",
        width=100
    )
    st.markdown("---")

    if st.session_state.get("seccion_activa", "consolidado") == "consolidado":
        st.markdown("### 📂 Cargar Archivos (Consolidado)")

        with st.expander("📊 iPago (Archivo Maestro)", expanded=True):
            archivo_ipago = st.file_uploader(
                "Cargar archivo iPago",
                type=["xlsx", "xls", "xlsm"],
                key="uploader_ipago"
            )

        st.markdown("#### 🏦 Bancos")

        with st.expander("🏦 Banesco", expanded=False):
            archivo_banesco = st.file_uploader(
                "Cargar Banesco",
                type=["xlsx", "xls", "xlsm"],
                accept_multiple_files=True,
                key="uploader_banesco"
            )

        with st.expander("🏦 BNC", expanded=False):
            archivo_bnc = st.file_uploader(
                "Cargar BNC",
                type=["xlsx", "xls", "xlsm"],
                accept_multiple_files=True,
                key="uploader_bnc"
            )

        with st.expander("🏦 Mercantil", expanded=False):
            archivo_mercantil = st.file_uploader(
                "Cargar Mercantil",
                type=["xlsx", "xls", "xlsm"],
                accept_multiple_files=True,
                key="uploader_mercantil"
            )

        with st.expander("🏦 Banco de Venezuela (BDV)", expanded=False):
            archivo_venezuela = st.file_uploader(
                "Cargar BDV",
                type=["xlsx", "xls", "xlsm"],
                accept_multiple_files=True,
                key="uploader_venezuela"
            )

        with st.expander("🏦 Provincial", expanded=False):
            archivo_provincial = st.file_uploader(
                "Cargar Provincial",
                type=["xlsx", "xls", "xlsm"],
                accept_multiple_files=True,
                key="uploader_provincial"
            )

        with st.expander("🏦 Bancamiga", expanded=False):
            archivo_bancamiga = st.file_uploader(
                "Cargar Bancamiga",
                type=["xlsx", "xls", "xlsm"],
                accept_multiple_files=True,
                key="uploader_bancamiga"
            )

        with st.expander("🏦 Banco del Tesoro", expanded=False):
            saldo_manual_tesoro = st.number_input(
                "Saldo manual Banco del Tesoro (VES)",
                min_value=0.0,
                value=0.0,
                step=100.0,
                key="saldo_manual_tesoro"
            )
    else:
        st.markdown("### 📂 Cargar Archivo Único (Monobanco)")

        archivo = st.file_uploader(
            "📂 Cargar archivo Excel",
            type=["xlsx", "xls", "xlsm"],
            key="uploader_monobanco"
        )

        archivo_ipago = st.file_uploader(
            "📂 Cargar archivo iPago",
            type=["xlsx", "xls", "xlsm"],
            key="uploader_ipago_mono"
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
            except Exception as e:
                st.warning(f"⚠️ Error leyendo como Excel, intentando como HTML: {str(e)}")
                archivo.seek(0)
                try:
                    tablas = pd.read_html(archivo)
                    if len(tablas) > 0:
                        return tablas[0]
                except Exception:
                    pass
                archivo.seek(0)
                contenido = archivo.read().decode("utf-8", errors="ignore")
                lineas = contenido.split("\n")
                datos = []
                for linea in lineas:
                    if linea.strip():
                        partes = linea.split("\t")
                        if len(partes) == 1:
                            partes = [p for p in linea.split(" ") if p.strip()]
                        if len(partes) > 0:
                            datos.append(partes)
                return pd.DataFrame(datos)
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
        try:
            return pd.read_excel(archivo, sheet_name=0, header=None, engine='openpyxl')
        except:
            st.error(f"No se pudo leer el archivo. Error: {str(e)}")
            st.stop()

# =========================================================
# 🔥 DETECCIÓN DE BANCO POR CONTENIDO DEL ARCHIVO
# =========================================================

def detectar_banco_por_contenido(archivo):
    """
    Detecta el banco leyendo el contenido del archivo, no solo el nombre.
    """
    try:
        pos = archivo.tell()
        archivo.seek(0)
        try:
            df_temp = pd.read_excel(archivo, nrows=20, header=None, engine='openpyxl')
            texto = " ".join(df_temp.astype(str).values.flatten()).upper()
            if "BANCAMIGA" in texto or "BANCAMIGA BANCO UNIVERSAL" in texto:
                archivo.seek(pos)
                return "bancamiga"
            elif "BANESCO" in texto:
                archivo.seek(pos)
                return "banesco"
            elif "PROVINCIAL" in texto:
                archivo.seek(pos)
                return "provincial"
            elif "BANCO DE VENEZUELA" in texto or "BDV" in texto:
                archivo.seek(pos)
                return "venezuela"
            elif "BNC" in texto:
                archivo.seek(pos)
                return "bnc"
            elif "MERCANTIL" in texto:
                archivo.seek(pos)
                return "mercantil"
            elif "TESORO" in texto or "BANCO DEL TESORO" in texto:
                archivo.seek(pos)
                return "tesoro"
        except Exception:
            pass
        archivo.seek(pos)
        return None
    except Exception:
        return None

def detectar_banco_por_nombre(nombre_archivo):
    """Detecta el banco por el nombre del archivo (fallback)"""
    nombre = nombre_archivo.upper()
    if "TESORO" in nombre or "TESORERIA" in nombre or "TES" in nombre:
        return "tesoro"
    elif "BANCAMIGA" in nombre or "BANCAAMIGA" in nombre:
        return "bancamiga"
    elif "BANESCO" in nombre or re.match(r"^J\d+", nombre_archivo):
        return "banesco"
    elif (
        "MOVIMIENTOS EN MONEDA NACIONAL" in nombre
        or "VENEZUELA" in nombre
        or "BANCO DE VENEZUELA" in nombre
        or "BDV" in nombre
        or "VZLA" in nombre
    ):
        return "venezuela"
    elif "PROVINCIAL" in nombre:
        return "provincial"
    elif "BNC" in nombre:
        return "bnc"
    elif "MERCANTIL" in nombre:
        return "mercantil"
    return "mercantil"

# =========================================================
# FUNCIONES DE CONVERSIÓN Y LIMPIEZA
# =========================================================

def convertir_monto(valor):
    try:
        if pd.isna(valor):
            return None
        if isinstance(valor, (int, float)):
            numero = float(valor)
            if isinstance(valor, int) and numero >= 100000:
                numero = numero / 100
            return numero
        valor_original = str(valor).strip()
        valor = valor_original
        valor = valor.replace(" ", "").replace("$", "").replace("Bs", "").replace("€", "")
        if valor == "":
            return None
        if "." in valor and "," in valor:
            valor = valor.replace(".", "").replace(",", ".")
        elif "," in valor:
            valor = valor.replace(",", ".")
        numero = float(valor)
        if "." not in valor_original and "," not in valor_original and numero >= 100000:
            numero = numero / 100
        return numero
    except Exception:
        return None

def calcular_usd(monto_bs, tasa):
    try:
        if monto_bs is None or tasa is None or tasa == 0:
            return None
        return round(abs(monto_bs) / abs(tasa), 2)
    except:
        return None

def es_comision(texto, proveedor=None):
    texto = normalizar_texto(texto).strip()
    if proveedor and str(proveedor).strip():
        return False
    if any(x in texto for x in [
        "comisiones sobre servicios contratados",
        "comision gerente comercial",
        "comision vendedor",
        "comision asesor",
        "comision ejecutivo",
        "comision supervisor",
        "comision ventas",
        "comisiones ventas",
        "comision comercial",
        "comisiones comerciales"
    ]):
        return False
    if any(x in texto for x in [
        "pago a proveedores",
        "pago de nomina",
        "nomina",
        "transf entre ctas",
        "transferencia a terceros",
        "pago movil comercial"
    ]):
        return False
    palabras_comision_bancaria = [
        "comision por transferencia",
        "comision pago movil",
        "comisión pago movil",
        "comision x pago de nomina",
        "comision x pago de nominas",
        "itf",
        "impuesto a las transacciones financieras",
        "cargo bancario",
        "mantenimiento de cuenta",
        "comision bancaria",
        "comisión bancaria",
        "cargo por servicio",
        "cargo por transaccion",
        "comision por",
        "comisión por",
        "comision pago movil comercial",
        "comision x pago de nominas mb",
        "com pago otras ctas",
        "com pago otr bcos",
        "comis",
        "comis. cr.i"
    ]
    for patron in palabras_comision_bancaria:
        if patron in texto:
            return True
    return False

# =========================================================
# ENRIQUECER EGRESOS CON IPAGO
# =========================================================

def enriquecer_egresos_con_ipago(df_egresos, df_ipago):
    if df_ipago is None or df_ipago.empty:
        return df_egresos
    df_resultado = df_egresos.copy()
    df_resultado["REFERENCIA_NORM"] = df_resultado["REFERENCIA"].astype(str).str.replace(".0", "", regex=False).str.strip()
    df_ipago["Referencia_Norm"] = df_ipago["Referencia"].astype(str).str.replace(".0", "", regex=False).str.strip()
    
    def generar_variantes_ref(ref):
        ref = str(ref).strip()
        if not ref or ref == "nan":
            return set()
        variantes = set()
        variantes.add(ref)
        ref_sin_ceros = ref.lstrip('0')
        if ref_sin_ceros != ref and ref_sin_ceros:
            variantes.add(ref_sin_ceros)
        if ref.endswith('X'):
            ref_sin_x = ref[:-1]
            if ref_sin_x:
                variantes.add(ref_sin_x)
                ref_sin_x_sin_ceros = ref_sin_x.lstrip('0')
                if ref_sin_x_sin_ceros:
                    variantes.add(ref_sin_x_sin_ceros)
        if 'X' in ref and not ref.endswith('X'):
            ref_sin_x = ref.replace('X', '')
            if ref_sin_x:
                variantes.add(ref_sin_x)
                ref_sin_x_sin_ceros = ref_sin_x.lstrip('0')
                if ref_sin_x_sin_ceros:
                    variantes.add(ref_sin_x_sin_ceros)
        if len(ref) >= 10 and ref.startswith('0'):
            ref_sin_cero_inicial = ref[1:]
            if ref_sin_cero_inicial:
                variantes.add(ref_sin_cero_inicial)
                variantes.add(ref_sin_cero_inicial.lstrip('0'))
        if '.' in ref:
            ref_sin_puntos = ref.replace('.', '')
            if ref_sin_puntos:
                variantes.add(ref_sin_puntos)
                variantes.add(ref_sin_puntos.lstrip('0'))
        return variantes

    ipago_dict = {}
    for _, row in df_ipago.iterrows():
        ref_original = str(row.get("Referencia", "")).strip()
        if not ref_original or ref_original == "nan":
            continue
        variantes = generar_variantes_ref(ref_original)
        for variante in variantes:
            if variante and variante not in ipago_dict:
                ipago_dict[variante] = {
                    "PROVEEDOR": row.get("Proveedor", ""),
                    "TIPO_EGRESO": row.get("Tipo de Egreso", ""),
                    "TIPO_PAGO": row.get("Tipo de Pago", ""),
                    "DESCRIPCION_IPAGO": row.get("Descripción", ""),
                    "FECHA_PAGO": row.get("Fecha Pago", ""),
                    "EMPRESA": row.get("Empresa", ""),
                    "MONTO_IPAGO": row.get("Monto", 0),
                    "MONTO_USD": row.get("Monto USD", 0),
                    "REFERENCIA_ORIGINAL": ref_original
                }

    for idx, row in df_resultado.iterrows():
        ref_banco = str(row.get("REFERENCIA", "")).strip()
        monto_banco = float(row.get("MONTO BS", 0))
        variantes_banco = generar_variantes_ref(ref_banco)
        coincide_ref = False
        datos_encontrados = None
        for variante in variantes_banco:
            if variante in ipago_dict:
                datos_encontrados = ipago_dict[variante]
                coincide_ref = True
                break
        if not coincide_ref:
            for clave, datos in ipago_dict.items():
                monto_ipago = float(datos.get("MONTO_IPAGO", 0))
                if monto_ipago > 0:
                    diferencia = abs(monto_banco - monto_ipago) / max(monto_banco, monto_ipago)
                    if diferencia < 0.01:
                        datos_encontrados = datos
                        coincide_ref = True
                        break
        if datos_encontrados:
            df_resultado.at[idx, "STATUS"] = datos_encontrados["PROVEEDOR"]
            df_resultado.at[idx, "OBSERVACIÓN"] = datos_encontrados["TIPO_EGRESO"]
            df_resultado.at[idx, "TIPO_PAGO"] = datos_encontrados["TIPO_PAGO"]
            df_resultado.at[idx, "PROVEEDOR_IPAGO"] = datos_encontrados["PROVEEDOR"]
            df_resultado.at[idx, "REFERENCIA_IPAGO"] = datos_encontrados.get("REFERENCIA_ORIGINAL", "")
            descripcion_ipago = datos_encontrados["DESCRIPCION_IPAGO"]
            if descripcion_ipago:
                df_resultado.at[idx, "DESCRIPCIÓN"] = descripcion_ipago
                df_resultado.at[idx, "DESCRIPCION_ORIGINAL"] = row.get("DESCRIPCIÓN", "")
            tipo = str(datos_encontrados["TIPO_EGRESO"]).upper()
            desc = str(datos_encontrados["DESCRIPCION_IPAGO"]).upper()
            if "COMISION" in tipo or "COMISION" in desc:
                df_resultado.at[idx, "ES_COMISION"] = True
            else:
                df_resultado.at[idx, "ES_COMISION"] = False
        else:
            df_resultado.at[idx, "STATUS"] = "SIN DATOS IPAGO"
            df_resultado.at[idx, "OBSERVACIÓN"] = "SIN CONCORDANCIA"
            df_resultado.at[idx, "TIPO_PAGO"] = ""
            df_resultado.at[idx, "PROVEEDOR_IPAGO"] = ""
            df_resultado.at[idx, "ES_COMISION"] = False
            
    df_resultado = df_resultado.drop(columns=["REFERENCIA_NORM"], errors="ignore")
    return df_resultado

# =========================================================
# PROCESADORES ESPECÍFICOS POR BANCO
# =========================================================

def procesar_banesco(df):
    st.info("Procesando Banesco...")
    try:
        df.columns = ["FECHA", "REFERENCIA", "DESCRIPCION", "MONTO_RAW", "BALANCE"]
        df.columns = [str(c).strip().upper() for c in df.columns]
        rename_map = {}
        for col in df.columns:
            c = str(col).lower()
            if "fecha" in c: rename_map[col] = "FECHA"
            elif "referencia" in c: rename_map[col] = "REFERENCIA"
            elif "descrip" in c: rename_map[col] = "DESCRIPCION"
            elif "monto" in c: rename_map[col] = "MONTO_RAW"
        df = df.rename(columns=rename_map)
        for col in ["FECHA", "REFERENCIA", "DESCRIPCION", "MONTO_RAW"]:
            if col not in df.columns:
                st.error(f"No existe columna: {col}")
                return pd.DataFrame()
        df["FECHA"] = pd.to_datetime(df["FECHA"], dayfirst=True, errors="coerce")
        df = df[df["FECHA"].notna()]
        df["TIPO"] = df["MONTO_RAW"].astype(str).apply(lambda x: "NC" if "+" in x else "ND")
        df["MONTO"] = (df["MONTO_RAW"].astype(str).str.replace("+", "", regex=False)
                       .str.replace("-", "", regex=False)
                       .str.replace(".", "", regex=False)
                       .str.replace(",", ".", regex=False)
                       .str.strip())
        df["MONTO"] = pd.to_numeric(df["MONTO"], errors="coerce")
        df = df[df["MONTO"].notna()]
        df = df[df["MONTO"] > 0]
        df = df[["FECHA", "REFERENCIA", "DESCRIPCION", "TIPO", "MONTO"]]
        st.success(f"Banesco OK: {len(df)} movimientos")
        return df
    except Exception as e:
        st.error(f"Error Banesco: {str(e)}")
        return pd.DataFrame()

def procesar_provincial(df):
    st.info("🔍 Procesando archivo de Provincial...")
    try:
        encabezado_idx = None
        for i in range(min(30, len(df))):
            fila = df.iloc[i]
            fila_str = [str(val) for val in fila.tolist()]
            texto_fila = " ".join(fila_str).upper()
            if "F. OPERACIÓN" in texto_fila or "F. VALOR" in texto_fila or "CONCEPTO" in texto_fila:
                encabezado_idx = i
                break
        if encabezado_idx is None:
            st.error("❌ No se encontró la fila de encabezados en el archivo Provincial.")
            return pd.DataFrame()
        headers = df.iloc[encabezado_idx].astype(str).str.strip().tolist()
        rename_map = {}
        for col in headers:
            col_clean = str(col).strip().upper()
            if "F. OPERACIÓN" in col_clean or "FECHA" in col_clean: rename_map[col] = "FECHA"
            elif "F. VALOR" in col_clean: rename_map[col] = "FECHA_VALOR"
            elif "CÓDIGO" in col_clean or "CODIGO" in col_clean: rename_map[col] = "CODIGO"
            elif "Nº. DOC" in col_clean or "NRO DOC" in col_clean or "DOC" in col_clean: rename_map[col] = "REFERENCIA"
            elif "CONCEPTO" in col_clean: rename_map[col] = "DESCRIPCION"
            elif "IMPORTE" in col_clean: rename_map[col] = "MONTO"
        df.columns = headers
        df = df.iloc[encabezado_idx + 1:].reset_index(drop=True)
        df = df.rename(columns=rename_map)
        if "FECHA" in df.columns:
            df["FECHA"] = df["FECHA"].astype(str).str.strip()
            df = df[~df["FECHA"].str.contains("FECHA|SALDO|Período", case=False, na=False)]
            df = df[df["FECHA"].str.match(r'^\d{2}-\d{2}-\d{4}$', na=False)]
            df["FECHA"] = pd.to_datetime(df["FECHA"], dayfirst=True, errors="coerce")
            df = df[df["FECHA"].notna()]
        else:
            return pd.DataFrame()
        if "MONTO" in df.columns:
            df["MONTO"] = df["MONTO"].astype(str).str.replace(" ", "", regex=False).str.replace(".", "", regex=False).str.replace(",", ".", regex=False).str.replace("'", "", regex=False)
            df["MONTO"] = pd.to_numeric(df["MONTO"], errors="coerce")
            df = df[df["MONTO"].notna()]
            df["TIPO"] = df["MONTO"].apply(lambda x: "NC" if x > 0 else "ND" if x < 0 else "")
            df["MONTO"] = df["MONTO"].abs()
            df = df[df["MONTO"] > 0]
        else:
            return pd.DataFrame()
        if "REFERENCIA" not in df.columns: df["REFERENCIA"] = ""
        else: df["REFERENCIA"] = df["REFERENCIA"].astype(str).str.strip().str.replace("'", "", regex=False)
        if "DESCRIPCION" not in df.columns: df["DESCRIPCION"] = ""
        else: df["DESCRIPCION"] = df["DESCRIPCION"].astype(str).str.strip()
        df["ES_COMISION"] = df["DESCRIPCION"].str.contains("COMIS", case=False, na=False)
        df_resultado = df[["FECHA", "REFERENCIA", "DESCRIPCION", "TIPO", "MONTO", "ES_COMISION"]].copy()
        st.success(f"✅ Provincial OK: {len(df_resultado)} movimientos")
        return df_resultado
    except Exception as e:
        st.error(f"❌ Error procesando Provincial: {str(e)}")
        return pd.DataFrame()

def procesar_bnc(df):
    st.info("Procesando archivo BNC...")
    try:
        encabezado = None
        for i in range(min(30, len(df))):
            fila = df.iloc[i].fillna("").astype(str)
            texto = " ".join(fila.tolist()).lower()
            if "fecha" in texto and ("descripcion" in texto or "descripción" in texto):
                encabezado = i
                break
        if encabezado is None:
            return pd.DataFrame()
        headers = []
        for idx, col in enumerate(df.iloc[encabezado]):
            col = str(col).strip().replace("\n", " ")
            if col == "" or col.lower() == "nan": col = f"COLUMNA_{idx}"
            headers.append(col)
        headers_unicos = []
        contador = {}
        for h in headers:
            if h in contador:
                contador[h] += 1
                nuevo = f"{h}_{contador[h]}"
            else:
                contador[h] = 0
                nuevo = h
            headers_unicos.append(nuevo)
        df.columns = headers_unicos
        df = df.iloc[encabezado + 1:].reset_index(drop=True)
        rename_map = {}
        for col in df.columns:
            col_str = str(col).strip().lower()
            if "fecha" in col_str: rename_map[col] = "FECHA"
            elif "descripcion" in col_str or "descripción" in col_str or "concepto" in col_str: rename_map[col] = "DESCRIPCION"
            elif "referencia" in col_str: rename_map[col] = "REFERENCIA"
            elif "credito" in col_str or "haber" in col_str: rename_map[col] = "CREDITO"
            elif "debito" in col_str or "debe" in col_str: rename_map[col] = "DEBITO"
        df = df.rename(columns=rename_map)
        if "FECHA" in df.columns:
            df["FECHA"] = pd.to_datetime(df["FECHA"], dayfirst=True, errors="coerce")
            df = df[df["FECHA"].notna()]
        df["CREDITO"] = pd.to_numeric(df.get("CREDITO", 0), errors="coerce").fillna(0) if "CREDITO" in df.columns else 0
        df["DEBITO"] = pd.to_numeric(df.get("DEBITO", 0), errors="coerce").fillna(0) if "DEBITO" in df.columns else 0
        df["MONTO"] = df["CREDITO"] - df["DEBITO"]
        df["TIPO"] = df["MONTO"].apply(lambda x: "NC" if x > 0 else "ND")
        df["MONTO"] = df["MONTO"].abs()
        df = df[df["MONTO"] != 0]
        st.success(f"Registros BNC OK: {len(df)}")
        return df
    except Exception as e:
        st.error(f"Error BNC: {e}")
        return pd.DataFrame()

def procesar_tesoro(df):
    st.info("Procesando Banco del Tesoro...")
    try:
        encabezado = None
        for i in range(min(20, len(df))):
            fila = df.iloc[i].astype(str)
            texto = " ".join(map(str, fila.tolist())).lower()
            if "fecha" in texto and "referencia" in texto and "concepto" in texto:
                encabezado = i
                break
        if encabezado is None:
            return pd.DataFrame()
        df.columns = df.iloc[encabezado]
        df = df.iloc[encabezado + 1:].reset_index(drop=True)
        df.columns = [str(c).strip() for c in df.columns]
        rename_map = {}
        for col in df.columns:
            c = str(col).strip().lower()
            if "fecha" in c: rename_map[col] = "FECHA"
            elif "referencia" in c: rename_map[col] = "REFERENCIA"
            elif "concepto" in c: rename_map[col] = "DESCRIPCION"
            elif "débito" in c or "debito" in c: rename_map[col] = "DEBITO"
            elif "crédito" in c or "credito" in c: rename_map[col] = "CREDITO"
            elif "código" in c or "codigo" in c: rename_map[col] = "TIPO"
        df = df.rename(columns=rename_map)
        if "FECHA" not in df.columns: return pd.DataFrame()
        df["FECHA"] = pd.to_datetime(df["FECHA"], dayfirst=True, errors="coerce")
        df = df[df["FECHA"].notna()]
        def limpiar_numero(valor):
            valor = str(valor).replace(".", "").replace(",", ".")
            try: return float(valor)
            except: return 0
        df["CREDITO"] = df.get("CREDITO", 0).apply(limpiar_numero) if "CREDITO" in df.columns else 0
        df["DEBITO"] = df.get("DEBITO", 0).apply(limpiar_numero) if "DEBITO" in df.columns else 0
        df["MONTO"] = df["CREDITO"] - df["DEBITO"]
        df["TIPO"] = df["MONTO"].apply(lambda x: "NC" if x > 0 else "ND")
        df["MONTO"] = df["MONTO"].abs()
        df = df[df["MONTO"] > 0]
        df = df[["FECHA", "REFERENCIA", "DESCRIPCION", "TIPO", "MONTO"]]
        st.success(f"Tesoro OK: {len(df)} registros")
        return df
    except Exception as e:
        st.error(f"Error Tesoro: {str(e)}")
        return pd.DataFrame()

def procesar_bancamiga(df):
    st.info("🔍 Procesando archivo de Bancamiga...")
    try:
        columnas = [str(c).strip().upper() for c in df.columns]
        if "FECHA" in columnas and "REFERENCIA" in columnas:
            rename_map = {
                "NRO.": "NRO", "NRO": "NRO", "FECHA": "FECHA", "REFERENCIA": "REFERENCIA",
                "CONCEPTO": "DESCRIPCION", "DÉBITO": "DEBITO", "DEBITO": "DEBITO",
                "CRÉDITO": "CREDITO", "CREDITO": "CREDITO", "SALDO": "SALDO"
            }
            df.columns = [rename_map.get(str(c).strip().upper(), str(c).strip().upper()) for c in df.columns]
        else:
            encabezado_idx = None
            for i in range(min(30, len(df))):
                fila = df.iloc[i]
                fila_str = [str(val) for val in fila.tolist()]
                texto_fila = " ".join(fila_str).upper()
                if "NRO" in texto_fila and "FECHA" in texto_fila and "REFERENCIA" in texto_fila:
                    encabezado_idx = i
                    break
            if encabezado_idx is None: return pd.DataFrame()
            headers = df.iloc[encabezado_idx].astype(str).str.strip().tolist()
            rename_map = {}
            for col in headers:
                col_clean = str(col).strip().upper()
                if "NRO" in col_clean or "Nº" in col_clean: rename_map[col] = "NRO"
                elif "FECHA" in col_clean: rename_map[col] = "FECHA"
                elif "REFERENCIA" in col_clean: rename_map[col] = "REFERENCIA"
                elif "CONCEPTO" in col_clean: rename_map[col] = "DESCRIPCION"
                elif "DÉBITO" in col_clean or "DEBITO" in col_clean: rename_map[col] = "DEBITO"
                elif "CRÉDITO" in col_clean or "CREDITO" in col_clean: rename_map[col] = "CREDITO"
                elif "SALDO" in col_clean: rename_map[col] = "SALDO"
            df.columns = headers
            df = df.iloc[encabezado_idx + 1:].reset_index(drop=True)
            df = df.rename(columns=rename_map)
            
        if "FECHA" not in df.columns: return pd.DataFrame()
        df["FECHA"] = df["FECHA"].astype(str).str.strip()
        df = df[~df["FECHA"].str.contains("FECHA|SALDO|TOTAL|CRÉDITO|CREDITO|DÉBITO|DEBITO", case=False, na=False)]
        df["FECHA"] = pd.to_datetime(df["FECHA"], format="%d/%m/%Y", errors="coerce")
        mask = df["FECHA"].isna()
        if mask.any():
            df.loc[mask, "FECHA"] = pd.to_datetime(df.loc[mask, "FECHA"].astype(str), dayfirst=True, errors="coerce")
        df = df[df["FECHA"].notna()]
        
        df["DEBITO"] = df["DEBITO"].astype(str).str.replace(" ", "", regex=False).str.replace(".", "", regex=False).str.replace(",", ".", regex=False) if "DEBITO" in df.columns else "0"
        df["DEBITO"] = pd.to_numeric(df["DEBITO"], errors="coerce").fillna(0)
        df["CREDITO"] = df["CREDITO"].astype(str).str.replace(" ", "", regex=False).str.replace(".", "", regex=False).str.replace(",", ".", regex=False) if "CREDITO" in df.columns else "0"
        df["CREDITO"] = pd.to_numeric(df["CREDITO"], errors="coerce").fillna(0)
        
        df["MONTO"] = df["CREDITO"] - df["DEBITO"]
        df["TIPO"] = df["MONTO"].apply(lambda x: "NC" if x > 0 else "ND" if x < 0 else "")
        df["MONTO"] = df["MONTO"].abs()
        df = df[df["MONTO"] > 0]
        if "REFERENCIA" not in df.columns: df["REFERENCIA"] = ""
        else: df["REFERENCIA"] = df["REFERENCIA"].astype(str).str.strip().str.replace("'", "", regex=False)
        if "DESCRIPCION" not in df.columns: df["DESCRIPCION"] = ""
        else: df["DESCRIPCION"] = df["DESCRIPCION"].astype(str).str.strip()
        df["ES_COMISION"] = df["DESCRIPCION"].str.contains("Comisi", case=False, na=False)
        df_resultado = df[["FECHA", "REFERENCIA", "DESCRIPCION", "TIPO", "MONTO", "ES_COMISION"]].copy()
        st.success(f"✅ Bancamiga OK: {len(df_resultado)} movimientos")
        return df_resultado
    except Exception as e:
        st.error(f"❌ Error Bancamiga: {str(e)}")
        return pd.DataFrame()

def procesar_venezuela_simple(df):
    st.info("🔍 Procesando Banco de Venezuela (MODO SIMPLE)...")
    try:
        col_fecha = 3
        col_ref = 1
        col_desc = 2
        col_tipo = 4
        col_credito = 5
        col_debito = 6
        
        movimientos = []
        for idx in range(1, len(df)):
            try:
                fila = df.iloc[idx]
                if pd.isna(fila[col_fecha]): continue
                fecha_raw = str(fila[col_fecha]).strip()
                fecha_val = pd.to_datetime(fecha_raw, format="%d/%m/%Y", errors="coerce")
                if pd.isna(fecha_val):
                    fecha_val = pd.to_datetime(fecha_raw, dayfirst=True, errors="coerce")
                if pd.isna(fecha_val): continue
                
                referencia = str(fila[col_ref]).strip() if pd.notna(fila[col_ref]) else ""
                descripcion = str(fila[col_desc]).strip() if pd.notna(fila[col_desc]) else ""
                tipo_mov = str(fila[col_tipo]).strip().upper() if pd.notna(fila[col_tipo]) else ""
                
                desc_upper = descripcion.upper()
                if desc_upper in ["SALDO INICIAL", "SALDO FINAL", "TOTALES"]: continue
                
                val_credito = 0
                if pd.notna(fila[col_credito]):
                    clean_cred = str(fila[col_credito]).strip().replace(".", "").replace(",", ".")
                    try: val_credito = float(clean_cred)
                    except: pass
                
                val_debito = 0
                if pd.notna(fila[col_debito]):
                    clean_deb = str(fila[col_debito]).strip().replace(".", "").replace(",", ".")
                    try: val_debito = float(clean_deb)
                    except: pass
                
                monto = 0
                tipo = ""
                if tipo_mov == "NC":
                    monto = abs(val_credito)
                    tipo = "NC"
                elif tipo_mov == "ND":
                    monto = abs(val_debito)
                    tipo = "ND"
                else:
                    if abs(val_credito) > 0:
                        monto = abs(val_credito)
                        tipo = "NC"
                    elif abs(val_debito) > 0:
                        monto = abs(val_debito)
                        tipo = "ND"
                    else: continue
                
                if monto <= 0: continue
                movimientos.append({
                    "FECHA": fecha_val.strftime("%d/%m/%Y"),
                    "FECHA_OBJ": fecha_val,
                    "REFERENCIA": referencia,
                    "DESCRIPCION": descripcion,
                    "TIPO": tipo,
                    "MONTO": monto
                })
            except:
                continue
        df_resultado = pd.DataFrame(movimientos)
        st.success(f"✅ Venezuela OK: {len(df_resultado)} movimientos")
        return df_resultado
    except Exception as e:
        st.error(f"❌ Error en BDV: {str(e)}")
        return pd.DataFrame()

# =========================================================
# CONVERTIDORES A FORMATO MERCANTIL
# =========================================================

def convertir_venezuela_a_formato_mercantil(df):
    datos_convertidos = []
    for idx, fila in df.iterrows():
        try:
            fecha = fila["FECHA_OBJ"] if "FECHA_OBJ" in fila else fila["FECHA"]
            if pd.isna(fecha): continue
            if isinstance(fecha, (pd.Timestamp, datetime)):
                fecha_str = fecha.strftime("%d/%m/%Y")
            else:
                fecha_str = str(fecha)
            
            tipo = fila.get("TIPO", "") or ""
            descripcion = fila.get("DESCRIPCION", "") or ""
            referencia = fila.get("REFERENCIA", "") or ""
            monto = fila.get("MONTO", 0) or 0
            
            fila_convertida = ["", "", "", fecha_str, referencia, tipo, descripcion, monto, "", False]
            datos_convertidos.append(fila_convertida)
        except:
            continue
    df_convertido = pd.DataFrame(datos_convertidos)
    return df_convertido if len(df_convertido) > 0 else pd.DataFrame()

def convertir_a_formato_mercantil(df, banco):
    datos_convertidos = []
    # Normalizar columnas del dataframe a mayúsculas para evitar problemas de casing
    df_temp = df.copy()
    df_temp.columns = [str(c).strip().upper() for c in df_temp.columns]
    
    for idx, fila in df_temp.iterrows():
        try:
            fecha = fila.get("FECHA", "")
            if pd.isna(fecha): continue
            if isinstance(fecha, (pd.Timestamp, datetime)):
                fecha_str = fecha.strftime("%d/%m/%Y")
            else:
                fecha_str = str(fecha)
            
            tipo = fila.get("TIPO", "") or ""
            descripcion = fila.get("DESCRIPCION", "") or ""
            referencia = fila.get("REFERENCIA", "") or ""
            monto = fila.get("MONTO", 0) or 0
            es_comision_flag = bool(fila.get("ES_COMISION", False))
            
            fila_convertida = ["", "", "", fecha_str, referencia, tipo, descripcion, monto, "", es_comision_flag]
            datos_convertidos.append(fila_convertida)
        except:
            continue
    df_convertido = pd.DataFrame(datos_convertidos)
    return df_convertido if len(df_convertido) > 0 else pd.DataFrame()

# =========================================================
# OBTENER TASA BCV HISTORICA LOCAL
# =========================================================

@st.cache_data(ttl=3600)
def obtener_tasa_bcv_fecha(fecha_obj):
    tasas_bcv_local = {
        "01/06/2026": 554.4258, "02/06/2026": 557.9741, "03/06/2026": 558.6436,
        "04/06/2026": 560.3753, "05/06/2026": 563.2892, "06/06/2026": 567.6828,
        "07/06/2026": 567.6828, "08/06/2026": 567.6828, "09/06/2026": 567.6828,
        "10/06/2026": 572.6784, "11/06/2026": 577.5461, "12/06/2026": 582.6862,
        "13/06/2026": 587.4059, "14/06/2026": 587.4059, "15/06/2026": 587.4059,
        "16/06/2026": 592.5163, "17/06/2026": 596.7824, "18/06/2026": 602.3324,
        "19/06/2026": 607.3919, "20/06/2026": 612.4332, "21/06/2026": 612.4332,
        "22/06/2026": 612.4332, "23/06/2026": 617.6388, "24/06/2026": 621.5299,
        "25/06/2026": 621.5299, "26/06/2026": 622.2135, "27/06/2026": 623.0223,
        "28/06/2026": 623.0223, "29/06/2026": 623.0223, "30/06/2026": 623.0223,
    }
    fecha_str = fecha_obj.strftime("%d/%m/%Y")
    return tasas_bcv_local.get(fecha_str, None)

def obtener_tasa_por_fecha(fecha_obj, usar_api=False):
    return obtener_tasa_bcv_fecha(fecha_obj)

# =========================================================
# 🔥 PROCESAMIENTO PRINCIPAL - CLASIFICACIÓN
# =========================================================

def procesar_archivo(df, usar_api=False, banco=""):
    ingresos = []
    egresos = []
    comisiones = []
    registros_procesados = set()
    
    tipos_ingresos = ["NC", "C", "CREDITO", "ABONO"]
    tipos_egresos = ["ND", "D", "DEBITO", "DEBIT"]
    cache_tasas = {}

    for _, fila in df.iterrows():
        try:
            if len(fila) < 10: continue
            fecha_raw = str(fila[3]).strip()
            if fecha_raw.lower() == "nan": continue
            fecha_raw = fecha_raw.replace(".0", "")
            if len(fecha_raw) == 7:
                fecha = f"0{fecha_raw[0]}/{fecha_raw[1:3]}/{fecha_raw[3:]}"
            elif len(fecha_raw) == 8:
                fecha = f"{fecha_raw[0:2]}/{fecha_raw[2:4]}/{fecha_raw[4:]}"
            else:
                fecha = fecha_raw

            tipo = str(fila[5]).strip().upper()
            descripcion = str(fila[6]).strip()
            referencia = str(fila[4]).strip()
            monto_bs = convertir_monto(fila[7])
            if monto_bs is None or monto_bs == 0: continue
            
            fecha_obj = pd.to_datetime(fecha, dayfirst=True, errors="coerce")
            if pd.isna(fecha_obj): continue
            
            fecha_key = fecha_obj.strftime("%d/%m/%Y")
            tasa = cache_tasas.get(fecha_key) or obtener_tasa_por_fecha(fecha_obj, usar_api) or 1.0
            cache_tasas[fecha_key] = tasa

            monto_usd = calcular_usd(monto_bs, tasa)
            if monto_usd is None: continue
            
            texto = descripcion.upper()
            if texto in ["SALDO", "DESCRIPCION", "DESCRIPCIÓN", "REFERENCIA", "MOVIMIENTO", "FECHA", "SALDO INICIAL", "SALDO FINAL"]:
                continue

            registro = {
                "FECHA": fecha, "REFERENCIA": referencia, "DESCRIPCIÓN": descripcion,
                "MONTO BS": round(abs(monto_bs), 2), "TASA BCV": round(tasa, 4), "MONTO USD": monto_usd,
                "STATUS": "", "OBSERVACIÓN": "", "TIPO_PAGO": "", "PROVEEDOR_IPAGO": "", "DESCRIPCION_ORIGINAL": ""
            }

            clave = (fecha, referencia, descripcion, monto_usd, tipo)
            if clave in registros_procesados: continue
            registros_procesados.add(clave)

            # Reglas específicas por banco
            es_comision_banco = False
            if banco == "provincial" and "COMIS" in texto: es_comision_banco = True
            elif banco == "bancamiga" and "COMISI" in texto: es_comision_banco = True
            elif banco == "mercantil":
                patrones = [
                    "OP.CRED.DIRT. CLTE-CLTE", "OP CRED DIRT CLTE CLTE", "COMISION PAGO MOVIL COMERCIAL",
                    "COMISION POR TRANSFERENCIA DE FONDOS", "COMISION X PAGO DE NOMINAS",
                    "COMISION PAGO MOVIL COMERCIAL INTERBANCARIO", "COMISION X PAGO DE NOMINAS MB",
                    "ITF", "IMPUESTO A LAS TRANSACCIONES FINANCIERAS", "CARGO BANCARIO",
                    "MANTENIMIENTO DE CUENTA", "COMISION POR TRANSFERENCIA"
                ]
                if any(p in texto for p in patrones): es_comision_banco = True
            
            # BDV / Venezuela comisiones por descripción o referencia
            if banco == "venezuela":
                patrones_bdv = [
                    "COM PAGO OTRAS CTAS", "COMISION PAGO A PROVEEDORES", "COM PAGO OTR BCOS",
                    "COM PAGO OTRAS CTAS JUR NAT", "COM PAGO OTRAS CTAS JUR JUR", "COMISION POR TRANSFERENCIA",
                    "COMISION PAGO MOVIL", "COMISIÓN PAGO MOVIL", "COMISION X PAGO DE NOMINA",
                    "COMISION X PAGO DE NOMINAS", "ITF", "IMPUESTO A LAS TRANSACCIONES FINANCIERAS",
                    "CARGO BANCARIO", "MANTENIMIENTO DE CUENTA", "COMISION BANCARIA", "COMISIÓN BANCARIA",
                    "CARGO POR SERVICIO", "CARGO POR TRANSACCION", "COMISION PAGO MOVIL COMERCIAL",
                    "COMISION X PAGO DE NOMINAS MB"
                ]
                if any(p in texto for p in patrones_bdv) or referencia.startswith(("970", "972", "067")) or (tipo == "ND" and "COM" in texto):
                    es_comision_banco = True

            if es_comision_banco or es_comision(descripcion):
                comisiones.append(registro)
            elif tipo in tipos_ingresos:
                ingresos.append(registro)
            elif tipo in tipos_egresos:
                egresos.append(registro)
        except:
            continue
    if st.session_state.seccion_activa == "consolidado":
        # ---------------------------------------------------------
        # FLUX 1: CIERRE CONSOLIDADO MULTIBANCO
        # ---------------------------------------------------------
        if not archivo_banesco: st.session_state.saldo_banesco = 0.0
        if not archivo_bnc: st.session_state.saldo_bnc = 0.0
        if not archivo_mercantil: st.session_state.saldo_mercantil = 0.0
        if not archivo_venezuela: st.session_state.saldo_venezuela = 0.0
        if not archivo_provincial: st.session_state.saldo_provincial = 0.0
        if not archivo_bancamiga: st.session_state.saldo_bancamiga = 0.0
        st.session_state.saldo_tesoro = st.session_state.get("saldo_manual_tesoro", 0.0)

        # Renderizado de KPIs
        tasa_dia = obtener_tasa_bcv()
        total_ves = (
            st.session_state.saldo_banesco + st.session_state.saldo_bnc + 
            st.session_state.saldo_mercantil + st.session_state.saldo_venezuela + 
            st.session_state.saldo_provincial + st.session_state.saldo_bancamiga + 
            st.session_state.saldo_tesoro
        )
        total_usd = total_ves / tasa_dia if tasa_dia > 0 else 0.0

        bancos_con_saldo = []
        if st.session_state.saldo_banesco > 0: bancos_con_saldo.append(f"Banesco: Bs. {formato_venezolano(st.session_state.saldo_banesco)}")
        if st.session_state.saldo_bnc > 0: bancos_con_saldo.append(f"BNC: Bs. {formato_venezolano(st.session_state.saldo_bnc)}")
        if st.session_state.saldo_mercantil > 0: bancos_con_saldo.append(f"Mercantil: Bs. {formato_venezolano(st.session_state.saldo_mercantil)}")
        if st.session_state.saldo_venezuela > 0: bancos_con_saldo.append(f"BDV: Bs. {formato_venezolano(st.session_state.saldo_venezuela)}")
        if st.session_state.saldo_provincial > 0: bancos_con_saldo.append(f"Provincial: Bs. {formato_venezolano(st.session_state.saldo_provincial)}")
        if st.session_state.saldo_bancamiga > 0: bancos_con_saldo.append(f"Bancamiga: Bs. {formato_venezolano(st.session_state.saldo_bancamiga)}")
        if st.session_state.saldo_tesoro > 0: bancos_con_saldo.append(f"Tesoro: Bs. {formato_venezolano(st.session_state.saldo_tesoro)}")

        kpi_subtitle_text = " | ".join(bancos_con_saldo) if bancos_con_saldo else "Sin saldos cargados"

        st.markdown(f"""
        <div class="kpi-container">
            <div class="kpi-card">
                <div class="kpi-title">Total Saldos Bancos (VES)</div>
                <div class="kpi-value">Bs. {formato_venezolano(total_ves)}</div>
                <div class="kpi-subtitle">{kpi_subtitle_text}</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-title">Tasa Oficial BCV del Día</div>
                <div class="kpi-value">{tasa_dia:.4f} VES/USD</div>
                <div class="kpi-subtitle">Tasa del Banco Central de Venezuela</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-title">Total Equivalente en Dólares (USD)</div>
                <div class="kpi-value">${total_usd:,.2f}</div>
                <div class="kpi-subtitle">Convertido al tipo de cambio oficial</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        # LEER ARCHIVO IPAGO
        df_ipago = None
        if archivo_ipago:
            try:
                df_ipago = pd.read_excel(archivo_ipago, engine="openpyxl")
                df_ipago.columns = [str(c).strip() for c in df_ipago.columns]
                st.success(f"✅ Archivo iPago cargado: {len(df_ipago)} registros")
            except Exception as e:
                st.error(f"❌ Error leyendo archivo iPago: {e}")

        list_df_convertidos = []
        bancos_procesados = []
        saldos_detalle_excel = []

        # 1. Banesco
        if archivo_banesco:
            st.session_state.saldo_banesco = 0.0
            for idx, arch in enumerate(archivo_banesco, 1):
                try:
                    nombre = arch.name.lower()
                    if nombre.endswith(".xlsx") or nombre.endswith(".xlsm"):
                        df_raw = pd.read_excel(arch, engine="openpyxl", header=None)
                    else:
                        df_raw = pd.read_html(arch)[0]
                    
                    saldo_arch = obtener_saldo_banco(df_raw, "banesco")
                    st.session_state.saldo_banesco += saldo_arch
                    
                    nombre_banco = f"Banesco - Cuenta {idx}" if len(archivo_banesco) > 1 else "Banesco"
                    saldos_detalle_excel.append((nombre_banco, saldo_arch))
                    
                    df_normalizado = procesar_banesco(df_raw)
                    df_convertido = convertir_a_formato_mercantil(df_normalizado, "banesco")
                    if not df_convertido.empty:
                        list_df_convertidos.append(df_convertido)
                        if "Banesco" not in bancos_procesados:
                            bancos_procesados.append("Banesco")
                except Exception as e:
                    st.error(f"⚠️ Error leyendo Banesco ({arch.name}): {e}")
        else:
            saldos_detalle_excel.append(("Banesco", 0.0))

        # 2. BNC
        if archivo_bnc:
            st.session_state.saldo_bnc = 0.0
            for idx, arch in enumerate(archivo_bnc, 1):
                try:
                    df_raw = leer_excel_con_encabezados(arch)
                    enc_idx = encontrar_fila_encabezado(df_raw)
                    
                    saldo_arch = obtener_saldo_banco(df_raw, "bnc", enc_idx)
                    st.session_state.saldo_bnc += saldo_arch
                    
                    nombre_banco = f"BNC - Cuenta {idx}" if len(archivo_bnc) > 1 else "BNC"
                    saldos_detalle_excel.append((nombre_banco, saldo_arch))
                    
                    df_normalizado = procesar_bnc(df_raw)
                    df_convertido = convertir_a_formato_mercantil(df_normalizado, "bnc")
                    if not df_convertido.empty:
                        list_df_convertidos.append(df_convertido)
                        if "BNC" not in bancos_procesados:
                            bancos_procesados.append("BNC")
                except Exception as e:
                    st.error(f"⚠️ Error leyendo BNC ({arch.name}): {e}")
        else:
            saldos_detalle_excel.append(("BNC", 0.0))

        # 3. Mercantil
        if archivo_mercantil:
            st.session_state.saldo_mercantil = 0.0
            for idx, arch in enumerate(archivo_mercantil, 1):
                try:
                    df_raw = leer_excel_sin_encabezados(arch)
                    
                    saldo_arch = obtener_saldo_banco(df_raw, "mercantil")
                    st.session_state.saldo_mercantil += saldo_arch
                    
                    nombre_banco = f"Mercantil - Cuenta {idx}" if len(archivo_mercantil) > 1 else "Mercantil"
                    saldos_detalle_excel.append((nombre_banco, saldo_arch))
                    
                    # Mercantil ya viene en su formato esperado directo
                    if not df_raw.empty:
                        list_df_convertidos.append(df_raw)
                        if "Mercantil" not in bancos_procesados:
                            bancos_procesados.append("Mercantil")
                except Exception as e:
                    st.error(f"⚠️ Error leyendo Mercantil ({arch.name}): {e}")
        else:
            saldos_detalle_excel.append(("Mercantil", 0.0))

        # 4. Venezuela (BDV)
        if archivo_venezuela:
            st.session_state.saldo_venezuela = 0.0
            for idx, arch in enumerate(archivo_venezuela, 1):
                try:
                    df_raw = leer_excel_sin_encabezados(arch)
                    
                    saldo_arch = obtener_saldo_banco(df_raw, "venezuela")
                    st.session_state.saldo_venezuela += saldo_arch
                    
                    nombre_banco = f"Banco de Venezuela - Cuenta {idx}" if len(archivo_venezuela) > 1 else "Banco de Venezuela"
                    saldos_detalle_excel.append((nombre_banco, saldo_arch))
                    
                    df_normalizado = procesar_venezuela_simple(df_raw)
                    df_convertido = convertir_venezuela_a_formato_mercantil(df_normalizado)
                    if not df_convertido.empty:
                        list_df_convertidos.append(df_convertido)
                        if "Venezuela" not in bancos_procesados:
                            bancos_procesados.append("Venezuela")
                except Exception as e:
                    st.error(f"⚠️ Error leyendo BDV ({arch.name}): {e}")
        else:
            saldos_detalle_excel.append(("Banco de Venezuela", 0.0))

        # 5. Provincial
        if archivo_provincial:
            st.session_state.saldo_provincial = 0.0
            for idx, arch in enumerate(archivo_provincial, 1):
                try:
                    df_raw = leer_excel_sin_encabezados(arch)
                    
                    saldo_arch = obtener_saldo_banco(df_raw, "provincial")
                    st.session_state.saldo_provincial += saldo_arch
                    
                    nombre_banco = f"Provincial - Cuenta {idx}" if len(archivo_provincial) > 1 else "Provincial"
                    saldos_detalle_excel.append((nombre_banco, saldo_arch))
                    
                    df_normalizado = procesar_provincial(df_raw)
                    df_convertido = convertir_a_formato_mercantil(df_normalizado, "provincial")
                    if not df_convertido.empty:
                        list_df_convertidos.append(df_convertido)
                        if "Provincial" not in bancos_procesados:
                            bancos_procesados.append("Provincial")
                except Exception as e:
                    st.error(f"⚠️ Error leyendo Provincial ({arch.name}): {e}")
        else:
            saldos_detalle_excel.append(("Provincial", 0.0))

        # 6. Bancamiga
        if archivo_bancamiga:
            st.session_state.saldo_bancamiga = 0.0
            for idx, arch in enumerate(archivo_bancamiga, 1):
                try:
                    nombre = arch.name.lower()
                    if nombre.endswith(".xlsx") or nombre.endswith(".xlsm"):
                        df_raw = pd.read_excel(arch, engine="openpyxl", header=None)
                    else:
                        try:
                            df_raw = pd.read_excel(arch, header=None)
                        except Exception:
                            arch.seek(0)
                            df_raw = pd.read_html(arch)[0]
                    
                    saldo_arch = obtener_saldo_banco(df_raw, "bancamiga")
                    st.session_state.saldo_bancamiga += saldo_arch
                    
                    nombre_banco = f"Bancamiga - Cuenta {idx}" if len(archivo_bancamiga) > 1 else "Bancamiga"
                    saldos_detalle_excel.append((nombre_banco, saldo_arch))
                    
                    df_normalizado = procesar_bancamiga(df_raw)
                    df_convertido = convertir_a_formato_mercantil(df_normalizado, "bancamiga")
                    if not df_convertido.empty:
                        list_df_convertidos.append(df_convertido)
                        if "Bancamiga" not in bancos_procesados:
                            bancos_procesados.append("Bancamiga")
                except Exception as e:
                    st.error(f"⚠️ Error leyendo Bancamiga ({arch.name}): {e}")
        else:
            saldos_detalle_excel.append(("Bancamiga", 0.0))

        # 7. Tesoro (Manual)
        saldos_detalle_excel.append(("Banco del Tesoro", st.session_state.saldo_tesoro))

        st.session_state.saldos_detalle_excel = saldos_detalle_excel

        if list_df_convertidos:
            df_original = pd.concat(list_df_convertidos, ignore_index=True)
            
            # Filtrar por fechas
            try:
                def parsear_fechas_consolidado(columna_fechas):
                    fechas = []
                    for val in columna_fechas:
                        val_str = str(val).strip().replace(".0", "")
                        if not val_str or val_str == "nan":
                            fechas.append(pd.NaT)
                            continue
                        
                        # Si es numérico de 8 dígitos (formato Mercantil ddmmyyyy)
                        if len(val_str) == 8 and val_str.isdigit():
                            dt = pd.to_datetime(val_str, format="%d%m%Y", errors="coerce")
                            if pd.notna(dt):
                                fechas.append(dt)
                                continue
                        
                        # Parseo flexible general
                        parsed = False
                        for fmt in ["%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%Y/%m/%d"]:
                            dt = pd.to_datetime(val_str, format=fmt, errors="coerce")
                            if pd.notna(dt):
                                fechas.append(dt)
                                parsed = True
                                break
                        if not parsed:
                            dt = pd.to_datetime(val_str, errors="coerce", dayfirst=True)
                            fechas.append(dt)
                    return pd.Series(fechas)
                    
                fechas_convertidas = parsear_fechas_consolidado(df_original.iloc[:, 3])
                fecha_inicio_dt = pd.to_datetime(fecha_inicio)
                fecha_fin_dt = pd.to_datetime(fecha_fin)
                df_original = df_original[(fechas_convertidas >= fecha_inicio_dt) & (fechas_convertidas <= fecha_fin_dt)]
                st.success(f"📅 Movimientos consolidados de {', '.join(bancos_procesados)} filtrados del {fecha_inicio} al {fecha_fin} ({len(df_original)} registros)")
            except Exception as e:
                st.warning(f"⚠️ Error filtrando fechas consolidadas: {e}")

            if df_original.empty:
                st.warning("⚠️ No se encontraron movimientos en el rango de fechas.")
            else:
                with st.expander("👁️ Vista previa de movimientos consolidados (Formato Unificado)"):
                    st.dataframe(df_original.head(20), use_container_width=True)

                if procesar:
                    with st.spinner("🚀 Conciliando y cruzando transacciones con iPago..."):
                        ingresos, egresos, comisiones = procesar_archivo(df_original, usar_api, banco="multibanco")
                        df_ingresos = pd.DataFrame(ingresos)
                        df_egresos = pd.DataFrame(egresos)
                        df_comisiones = pd.DataFrame(comisiones)

                        # Cruce con iPago
                        if df_ipago is not None and not df_egresos.empty:
                            df_egresos = enriquecer_egresos_con_ipago(df_egresos, df_ipago)
                            if "ES_COMISION" in df_egresos.columns:
                                mascara = df_egresos["ES_COMISION"] == True
                                if mascara.any():
                                    df_comisiones_extra = df_egresos[mascara].copy().drop(columns=["ES_COMISION", "REFERENCIA_IPAGO"], errors="ignore")
                                    df_comisiones = pd.concat([df_comisiones, df_comisiones_extra], ignore_index=True) if not df_comisiones.empty else df_comisiones_extra
                                    df_egresos = df_egresos[~mascara].copy()
                                    st.success(f"💳 Se identificaron {len(df_comisiones_extra)} comisiones adicionales vía iPago.")
                            df_egresos = df_egresos.drop(columns=["ES_COMISION", "REFERENCIA_IPAGO"], errors="ignore")
                            st.success(f"🎯 Cruce completado. Egresos conciliados con iPago: {len(df_egresos)} registros.")

                        for df_t in [df_ingresos, df_egresos, df_comisiones]:
                            if not df_t.empty:
                                for col in ["STATUS", "OBSERVACIÓN", "TIPO_PAGO", "PROVEEDOR_IPAGO", "DESCRIPCION_ORIGINAL"]:
                                    if col not in df_t.columns: df_t[col] = ""

                        total_ingresos = df_ingresos["MONTO USD"].sum() if not df_ingresos.empty else 0
                        total_egresos = df_egresos["MONTO USD"].sum() if not df_egresos.empty else 0
                        total_comisiones = df_comisiones["MONTO USD"].sum() if not df_comisiones.empty else 0
                        neto_procesado = total_ingresos - total_egresos - total_comisiones

                        # Mostrar métricas
                        col1_m, col2_m, col3_m, col4_m = st.columns(4)
                        with col1_m: st.metric("💰 TOTAL INGRESOS PROCESADOS", len(df_ingresos), f"${total_ingresos:,.2f}")
                        with col2_m: st.metric("💸 TOTAL EGRESOS PROCESADOS", len(df_egresos), f"${total_egresos:,.2f}")
                        with col3_m: st.metric("💳 TOTAL COMISIONES PROCESADAS", len(df_comisiones), f"${total_comisiones:,.2f}")
                        with col4_m: st.metric("⚖️ NETO PROCESADO (ING - EGR - COM)", "", f"${neto_procesado:,.2f}")

                        st.subheader("📊 Detalle de Transacciones Conciliadas")
                        tab1, tab2, tab3 = st.tabs(["📈 INGRESOS", "📉 EGRESOS", "💳 COMISIONES"])
                        with tab1: st.dataframe(df_ingresos, use_container_width=True)
                        with tab2: st.dataframe(df_egresos, use_container_width=True)
                        with tab3: st.dataframe(df_comisiones, use_container_width=True)

                        # =========================================================
                        # MOTOR DE REPORTES EXCEL OPENPYXL COMPLETO
                        # =========================================================
                        output = BytesIO()
                        with pd.ExcelWriter(output, engine="openpyxl") as writer:
                            workbook = writer.book
                            
                            # -----------------------------------------------------
                            # PESTAÑA: RESUMEN DE SALDOS CONSOLIDADO
                            # -----------------------------------------------------
                            hoja_resumen = workbook.create_sheet(title="RESUMEN", index=0)
                            
                            rojo = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
                            azul_oscuro = PatternFill(start_color="1E3A5F", end_color="1E3A5F", fill_type="solid")
                            verde_claro = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")
                            gris_claro = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
                            amarillo = PatternFill(start_color="FFE699", end_color="FFE699", fill_type="solid")
                            blanco_bold = Font(color="FFFFFF", bold=True, size=11)
                            negro_bold = Font(color="000000", bold=True, size=11)
                            borde_fino = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
                            alineacion_centro = Alignment(horizontal="center", vertical="center")
                            centro = alineacion_centro
                            alineacion_derecha = Alignment(horizontal="right", vertical="center")
                            alineacion_izquierda = Alignment(horizontal="left", vertical="center")

                            # Cabecera Resumen (A la izquierda)
                            hoja_resumen["B2"] = "GRUPO BODEGUITA ORIENTE"
                            hoja_resumen["B2"].font = Font(bold=True, size=14, color="1E3A5F")
                            hoja_resumen["B2"].alignment = alineacion_izquierda

                            # Cabecera adicional
                            hoja_resumen["B3"] = "CONCILIACIÓN BANCARIA - RESUMEN DE SALDOS"
                            hoja_resumen["B3"].font = Font(bold=True, size=11, color="555555")
                            hoja_resumen["B3"].alignment = alineacion_izquierda

                            # Datos de la Empresa (A la derecha)
                            hoja_resumen["E2"] = "Fecha de Cierre:"
                            hoja_resumen["E2"].font = Font(bold=True, size=11, color="1E3A5F")
                            hoja_resumen["E2"].alignment = alineacion_derecha
                            
                            hoja_resumen["F2"] = date.today().strftime("%d/%m/%Y")
                            hoja_resumen["F2"].font = Font(bold=False, size=11)
                            hoja_resumen["F2"].alignment = alineacion_izquierda

                            hoja_resumen["E3"] = "Tasa BCV del Día:"
                            hoja_resumen["E3"].font = Font(bold=True, size=11, color="1E3A5F")
                            hoja_resumen["E3"].alignment = alineacion_derecha
                            
                            hoja_resumen["F3"] = tasa_dia
                            hoja_resumen["F3"].font = Font(bold=False, size=11)
                            hoja_resumen["F3"].number_format = '#,##0.0000'
                            hoja_resumen["F3"].alignment = alineacion_izquierda

                            # Tabla Resumen de Cuentas
                            hoja_resumen["B5"] = "BANCOS"
                            hoja_resumen["B5"].font = blanco_bold
                            hoja_resumen["B5"].fill = azul_oscuro
                            hoja_resumen["B5"].alignment = alineacion_centro
                            hoja_resumen["B5"].border = borde_fino

                            hoja_resumen["C5"] = "SALDO BS"
                            hoja_resumen["C5"].font = blanco_bold
                            hoja_resumen["C5"].fill = azul_oscuro
                            hoja_resumen["C5"].alignment = alineacion_centro
                            hoja_resumen["C5"].border = borde_fino

                            hoja_resumen["D5"] = "EQUIVALENTE USD"
                            hoja_resumen["D5"].font = blanco_bold
                            hoja_resumen["D5"].fill = azul_oscuro
                            hoja_resumen["D5"].alignment = alineacion_centro
                            hoja_resumen["D5"].border = borde_fino

                            fila_r = 6
                            for nombre_b, saldo_b in st.session_state.saldos_detalle_excel:
                                if saldo_b > 0 or nombre_b in ["Banco del Tesoro", "Banesco", "BNC", "Mercantil", "Banco de Venezuela", "Provincial", "Bancamiga"]:
                                    cell_n = hoja_resumen.cell(row=fila_r, column=2, value=nombre_b)
                                    cell_n.border = borde_fino
                                    cell_n.alignment = alineacion_izquierda
                                    if fila_r % 2 == 1:
                                        cell_n.fill = gris_claro

                                    cell_s = hoja_resumen.cell(row=fila_r, column=3, value=saldo_b)
                                    cell_s.border = borde_fino
                                    cell_s.number_format = '#,##0.00'
                                    cell_s.alignment = alineacion_derecha
                                    if fila_r % 2 == 1:
                                        cell_s.fill = gris_claro

                                    cell_u = hoja_resumen.cell(row=fila_r, column=4, value=(saldo_b / tasa_dia if tasa_dia > 0 else 0.0))
                                    cell_u.border = borde_fino
                                    cell_u.number_format = '$#,,##0.00'
                                    cell_u.alignment = alineacion_derecha
                                    if fila_r % 2 == 1:
                                        cell_u.fill = gris_claro

                                    fila_r += 1

                            # Totales
                            cell_total_lbl = hoja_resumen.cell(row=fila_r, column=2, value="TOTAL CONSOLIDADO")
                            cell_total_lbl.font = negro_bold
                            cell_total_lbl.border = borde_fino
                            cell_total_lbl.fill = verde_claro
                            
                            cell_total_ves = hoja_resumen.cell(row=fila_r, column=3, value=total_ves)
                            cell_total_ves.font = negro_bold
                            cell_total_ves.border = borde_fino
                            cell_total_ves.number_format = '#,##0.00'
                            cell_total_ves.fill = verde_claro
                            cell_total_ves.alignment = alineacion_derecha
                            
                            cell_total_usd = hoja_resumen.cell(row=fila_r, column=4, value=total_usd)
                            cell_total_usd.font = negro_bold
                            cell_total_usd.border = borde_fino
                            cell_total_usd.number_format = '$#,##0.00'
                            cell_total_usd.fill = verde_claro
                            cell_total_usd.alignment = alineacion_derecha

                            for columna in hoja_resumen.columns:
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
                                adjusted_width = min(max_length + 5, 50)
                                hoja_resumen.column_dimensions[columna_letra].width = adjusted_width

                            # -----------------------------------------------------
                            # PESTAÑA: DETALLE DE CONCILIACIÓN (REPORTE)
                            # -----------------------------------------------------
                            hoja = workbook.create_sheet(title="REPORTE")
                            if "Sheet" in workbook.sheetnames:
                                workbook.remove(workbook["Sheet"])

                            try:
                                logo = Image("LOGO.jpeg")
                                logo.width = 130
                                logo.height = 130
                                hoja.add_image(logo, "A1")
                            except:
                                pass

                            hoja.merge_cells("C7:H7")
                            hoja["C7"] = "REPORTE CONSOLIDADO DE CONCILIACIÓN MULTIBANCO"
                            hoja["C7"].font = Font(bold=True, size=14)
                            hoja["C7"].alignment = alineacion_centro

                            def crear_tabla(titulo, dataframe, fila_inicio, color_total):
                                hoja.merge_cells(start_row=fila_inicio, start_column=1, end_row=fila_inicio, end_column=10)
                                titulo_cell = hoja.cell(row=fila_inicio, column=1)
                                titulo_cell.value = titulo
                                titulo_cell.fill = rojo
                                titulo_cell.font = blanco_bold
                                titulo_cell.alignment = alineacion_centro

                                headers = [
                                    "FECHA", "REFERENCIA", "DESCRIPCIÓN", "DESCRIPCION_ORIGINAL",
                                    "MONTO BS", "TASA BCV", "MONTO USD", 
                                    "PROVEEDOR (iPago)", "TIPO EGRESO (iPago)", "TIPO PAGO (iPago)"
                                ]
                                fila_header = fila_inicio + 1

                                for col_num, header in enumerate(headers, 1):
                                    cell = hoja.cell(row=fila_header, column=col_num)
                                    cell.value = header
                                    cell.fill = rojo
                                    cell.font = blanco_bold
                                    cell.border = borde_fino
                                    cell.alignment = alineacion_centro

                                fila_data = fila_header + 1

                                for _, row in dataframe.iterrows():
                                    hoja.cell(row=fila_data, column=1).value = row.get("FECHA", "")
                                    hoja.cell(row=fila_data, column=2).value = row.get("REFERENCIA", "")
                                    hoja.cell(row=fila_data, column=3).value = row.get("DESCRIPCIÓN", "")
                                    hoja.cell(row=fila_data, column=4).value = row.get("DESCRIPCION_ORIGINAL", "")
                                    hoja.cell(row=fila_data, column=5).value = row.get("MONTO BS", 0)
                                    hoja.cell(row=fila_data, column=6).value = row.get("TASA BCV", 0)
                                    hoja.cell(row=fila_data, column=7).value = row.get("MONTO USD", 0)
                                    hoja.cell(row=fila_data, column=8).value = row.get("PROVEEDOR_IPAGO", row.get("STATUS", ""))
                                    hoja.cell(row=fila_data, column=9).value = row.get("OBSERVACIÓN", "")
                                    hoja.cell(row=fila_data, column=10).value = row.get("TIPO_PAGO", "")

                                    hoja.cell(row=fila_data, column=5).number_format = '#,##0.00'
                                    hoja.cell(row=fila_data, column=6).number_format = '#,##0.0000'
                                    hoja.cell(row=fila_data, column=7).number_format = '$#,##0.00'

                                    for col in range(1, 11):
                                         hoja.cell(row=fila_data, column=col).border = borde_fino
                                    fila_data += 1

                                total_cell = hoja.cell(row=fila_data, column=4)
                                total_cell.value = f"TOTAL {titulo}"
                                total_cell.font = Font(bold=True)

                                total_bs_cell = hoja.cell(row=fila_data, column=5)
                                total_bs_cell.value = dataframe["MONTO BS"].sum() if not dataframe.empty else 0
                                total_bs_cell.number_format = '#,##0.00'
                                total_bs_cell.fill = color_total

                                monto_total = hoja.cell(row=fila_data, column=7)
                                monto_total.value = dataframe["MONTO USD"].sum() if not dataframe.empty else 0
                                monto_total.number_format = '$#,##0.00'
                                monto_total.fill = color_total

                                return fila_data + 4

                            fila_actual = 10
                            if not df_ingresos.empty:
                                fila_actual = crear_tabla("INGRESOS", df_ingresos, fila_actual, verde_claro)
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
                                adjusted_width = min(max_length + 5, 50)
                                hoja.column_dimensions[columna_letra].width = adjusted_width

                        output.seek(0)

                        st.download_button(
                            label="📥 Descargar Excel Clasificado Consolidado (BCV + iPago)",
                            data=output.getvalue(),
                            file_name=f"cierre_consolidado_{fecha_inicio}_{fecha_fin}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                        
                        with st.expander("📊 Tasas BCV utilizadas en el proceso"):
                            todas_tasas = {}
                            for registro in ingresos + egresos + comisiones:
                                fecha_r = registro["FECHA"]
                                tasa_r = registro["TASA BCV"]
                                if fecha_r not in todas_tasas:
                                    todas_tasas[fecha_r] = tasa_r
                            if todas_tasas:
                                df_tasas = pd.DataFrame([
                                    {"FECHA": f, "TASA BCV": t} 
                                    for f, t in todas_tasas.items()
                                ]).sort_values("FECHA")
                                st.dataframe(df_tasas, use_container_width=True)

        else:
            st.markdown("""
            ### 👋 Conciliador Bancario Inteligente Multibanco
            
            Carga los archivos de tus cuentas bancarias y el archivo maestro de iPago en el menú de la izquierda para comenzar el proceso de conciliación automatizado.
            
            **Características:**
            - Soporte simultáneo para múltiples cuentas.
            - Cálculo automático de saldo consolidado en Bolívares (VES) y Dólares (USD).
            - Cruce inteligente y trazabilidad con iPago.
            - Generación de reportes de cierre en formato Excel profesional.
            """)

    else:
        # ---------------------------------------------------------
        # FLUX 2: CRUCE DE INFORMACIÓN (MONOBANCO)
        # ---------------------------------------------------------
        df_ipago = None

        if archivo:
            st.info(f"📄 Archivo: **{archivo.name}** - {archivo.size/1024:.1f} KB")

            try:
                # 1. Detectar banco
                banco = detectar_banco_por_nombre(archivo.name)
                if banco == "mercantil":
                    banco_contenido = detectar_banco_por_contenido(archivo)
                    if banco_contenido:
                        banco = banco_contenido

                st.success(f"🏦 **Banco detectado:** {banco.upper()}")

                if banco == "mercantil":
                    df_original = leer_excel_sin_encabezados(archivo)
                elif banco == "banesco":
                    try:
                        nombre = archivo.name.lower()
                        if nombre.endswith(".xlsx") or nombre.endswith(".xlsm"):
                            df_raw = pd.read_excel(archivo, engine="openpyxl", header=None)
                        else:
                            df_raw = pd.read_html(archivo)[0]
                        df_normalizado = procesar_banesco(df_raw)
                        df_original = convertir_a_formato_mercantil(df_normalizado, banco)
                    except Exception as e:
                        st.error(f"Error leyendo Banesco: {str(e)}")
                        st.stop()
                elif banco == "tesoro":
                    try:
                        df_raw = pd.read_excel(archivo, engine="openpyxl")
                        df_normalizado = procesar_tesoro(df_raw)
                        df_original = convertir_a_formato_mercantil(df_normalizado, banco)
                    except Exception as e:
                        st.error(f"Error leyendo Tesoro: {str(e)}")
                        st.stop()
                elif banco == "bancamiga":
                    try:
                        nombre = archivo.name.lower()
                        if nombre.endswith(".xlsx") or nombre.endswith(".xlsm"):
                            df_raw = pd.read_excel(archivo, engine="openpyxl", header=None)
                        else:
                            try:
                                df_raw = pd.read_excel(archivo, header=None)
                            except Exception:
                                archivo.seek(0)
                                df_raw = pd.read_html(archivo)[0]
                        df_normalizado = procesar_bancamiga(df_raw)
                        if df_normalizado.empty:
                            st.error("No se pudieron procesar los datos de Bancamiga.")
                            st.stop()
                        df_original = convertir_a_formato_mercantil(df_normalizado, banco)
                    except Exception as e:
                        st.error(f"Error leyendo Bancamiga: {str(e)}")
                        st.stop()
                elif banco == "provincial":
                    try:
                        df_raw = leer_excel_sin_encabezados(archivo)
                        df_normalizado = procesar_provincial(df_raw)
                        if df_normalizado.empty:
                            st.error("No se pudieron procesar los datos de Provincial.")
                            st.stop()
                        df_original = convertir_a_formato_mercantil(df_normalizado, banco)
                    except Exception as e:
                        st.error(f"Error leyendo Provincial: {str(e)}")
                        st.stop()
                elif banco == "venezuela":
                    df_raw = leer_excel_sin_encabezados(archivo)
                    df_normalizado = procesar_venezuela_simple(df_raw)
                    if df_normalizado.empty:
                        st.stop()
                    df_original = convertir_venezuela_a_formato_mercantil(df_normalizado)
                    fechas_convertidas = pd.to_datetime(df_normalizado["FECHA"], dayfirst=True, errors="coerce")
                elif banco == "bnc":
                    df_raw = leer_excel_con_encabezados(archivo)
                    df_normalizado = procesar_bnc(df_raw)
                    df_original = convertir_a_formato_mercantil(df_normalizado, banco)
                else:
                    df_raw = leer_excel_con_encabezados(archivo)
                    df_original = convertir_a_formato_mercantil(df_raw, banco)

                if df_original.empty:
                    st.error("No se detectaron movimientos para procesar.")
                    st.stop()

                try:
                    if banco == "mercantil":
                        fechas_convertidas = pd.to_datetime(df_original[3].astype(str).str.zfill(8), format="%d%m%Y", errors="coerce")
                    elif banco == "venezuela":
                        pass
                    else:
                        fechas_convertidas = pd.to_datetime(df_original.iloc[:, 3], errors="coerce", dayfirst=True)

                    fecha_inicio_dt = pd.to_datetime(fecha_inicio)
                    fecha_fin_dt = pd.to_datetime(fecha_fin)

                    if banco == "venezuela":
                        mask = (fechas_convertidas >= fecha_inicio_dt) & (fechas_convertidas <= fecha_fin_dt)
                        df_original = df_original[mask]
                    else:
                        df_original = df_original[(fechas_convertidas >= fecha_inicio_dt) & (fechas_convertidas <= fecha_fin_dt)]
                    st.success(f"Filtro de fechas aplicado: {fecha_inicio} a {fecha_fin}")
                except Exception as e:
                    st.warning(f"Error filtrando fechas: {e}")

                if df_original.empty or len(df_original) == 0:
                    st.error("❌ No se encontraron movimientos válidos en el rango de fechas.")
                    st.stop()

                with st.expander("👁️ Vista previa archivo original"):
                    st.dataframe(df_original.head(20), use_container_width=True)

                # Leer archivo iPago
                if archivo_ipago:
                    try:
                        df_ipago = pd.read_excel(archivo_ipago, engine="openpyxl")
                        df_ipago.columns = [str(c).strip() for c in df_ipago.columns]
                        st.success(f"Archivo iPago cargado: {len(df_ipago)} registros")
                        st.dataframe(df_ipago.head())
                    except Exception as e:
                        st.error(f"Error leyendo archivo iPago: {e}")

                if procesar:
                    with st.spinner("Procesando archivo con tasas BCV..."):
                        if banco == "venezuela":
                            ingresos = []
                            egresos = []
                            comisiones = []

                            for _, row in df_normalizado.iterrows():
                                fecha_obj = pd.to_datetime(row["FECHA"], dayfirst=True, errors="coerce")
                                tasa = obtener_tasa_por_fecha(fecha_obj, usar_api) or 1.0
                                monto_bs = float(row["MONTO"])
                                monto_usd = calcular_usd(monto_bs, tasa)

                                registro = {
                                    "FECHA": row["FECHA"], "REFERENCIA": row["REFERENCIA"], "DESCRIPCIÓN": row["DESCRIPCION"],
                                    "MONTO BS": monto_bs, "TASA BCV": tasa, "MONTO USD": monto_usd,
                                    "STATUS": "", "OBSERVACIÓN": "", "TIPO_PAGO": "", "PROVEEDOR_IPAGO": "", "DESCRIPCION_ORIGINAL": ""
                                }
                                tipo = str(row["TIPO"]).strip().upper()
                                descripcion = str(row["DESCRIPCION"]).strip().upper()
                                referencia = str(row["REFERENCIA"]).strip()

                                es_comision_venezuela = False
                                palabras_comision_bdv = [
                                    "COM PAGO OTRAS CTAS", "COMISION PAGO A PROVEEDORES", "COM PAGO OTR BCOS",
                                    "COM PAGO OTRAS CTAS JUR NAT", "COM PAGO OTRAS CTAS JUR JUR", "COMISION POR TRANSFERENCIA",
                                    "COMISION PAGO MOVIL", "COMISIÓN PAGO MOVIL", "COMISION X PAGO DE NOMINA",
                                    "COMISION X PAGO DE NOMINAS", "ITF", "IMPUESTO A LAS TRANSACCIONES FINANCIERAS",
                                    "CARGO BANCARIO", "MANTENIMIENTO DE CUENTA", "COMISION BANCARIA", "COMISIÓN BANCARIA",
                                    "CARGO POR SERVICIO", "CARGO POR TRANSACCION", "COMISION PAGO MOVIL COMERCIAL",
                                    "COMISION X PAGO DE NOMINAS MB"
                                ]
                                for patron in palabras_comision_bdv:
                                    if patron in descripcion:
                                        es_comision_venezuela = True
                                        break
                                if not es_comision_venezuela and referencia.startswith(("970", "972", "067")):
                                    if any(palabra in descripcion for palabra in ["COM", "PAGO OTRAS", "PAGO OTR", "COMISION"]):
                                        es_comision_venezuela = True
                                if not es_comision_venezuela and tipo == "ND":
                                    if "COM" in descripcion or "PAGO OTR" in descripcion:
                                        es_comision_venezuela = True
                                if not es_comision_venezuela and tipo == "ND" and monto_bs < 1000:
                                    if "COM" in descripcion or "PAGO OTR" in descripcion:
                                        es_comision_venezuela = True

                                if es_comision_venezuela:
                                    comisiones.append(registro)
                                elif tipo in ["NC", "C", "CREDITO", "ABONO"]:
                                    ingresos.append(registro)
                                else:
                                    egresos.append(registro)
                        else:
                            ingresos, egresos, comisiones = procesar_archivo(df_original, usar_api, banco=banco)

                    df_ingresos = pd.DataFrame(ingresos)
                    df_egresos = pd.DataFrame(egresos)
                    df_comisiones = pd.DataFrame(comisiones)

                    if archivo_ipago and not df_egresos.empty:
                        df_egresos = enriquecer_egresos_con_ipago(df_egresos, df_ipago)
                        if "ES_COMISION" in df_egresos.columns:
                            mascara_comisiones_ipago = df_egresos["ES_COMISION"] == True
                            if mascara_comisiones_ipago.any():
                                df_comisiones_extra = df_egresos[mascara_comisiones_ipago].copy().drop(columns=["ES_COMISION", "REFERENCIA_IPAGO"], errors="ignore")
                                df_comisiones = pd.concat([df_comisiones, df_comisiones_extra], ignore_index=True) if not df_comisiones.empty else df_comisiones_extra
                                df_egresos = df_egresos[~mascara_comisiones_ipago].copy()
                                st.success(f"💳 Se movieron {len(df_comisiones_extra)} comisiones desde iPago a la sección de COMISIONES")
                        df_egresos = df_egresos.drop(columns=["ES_COMISION", "REFERENCIA_IPAGO"], errors="ignore")
                        st.success(f"🎯 Egresos enriquecidos con iPago: {len(df_egresos)} registros")

                    for df_t in [df_ingresos, df_egresos, df_comisiones]:
                        if not df_t.empty:
                            for col in ["STATUS", "OBSERVACIÓN", "TIPO_PAGO", "PROVEEDOR_IPAGO", "DESCRIPCION_ORIGINAL"]:
                                if col not in df_t.columns: df_t[col] = ""

                    total_ingresos = df_ingresos["MONTO USD"].sum() if not df_ingresos.empty else 0
                    total_egresos = df_egresos["MONTO USD"].sum() if not df_egresos.empty else 0
                    total_comisiones = df_comisiones["MONTO USD"].sum() if not df_comisiones.empty else 0

                    col1, col2, col3 = st.columns(3)
                    with col1: st.metric("💰 INGRESOS", len(df_ingresos), f"${total_ingresos:,.2f}")
                    with col2: st.metric("💸 EGRESOS", len(df_egresos), f"${total_egresos:,.2f}")
                    with col3: st.metric("💳 COMISIONES", len(df_comisiones), f"${total_comisiones:,.2f}")

                    st.subheader("📊 Resultados")
                    tab1, tab2, tab3 = st.tabs(["📈 INGRESOS", "📉 EGRESOS", "💳 COMISIONES"])
                    with tab1: st.dataframe(df_ingresos, use_container_width=True)
                    with tab2: st.dataframe(df_egresos, use_container_width=True)
                    with tab3: st.dataframe(df_comisiones, use_container_width=True)

                    # excel openpyxl output
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine="openpyxl") as writer:
                        workbook = writer.book
                        hoja = workbook.create_sheet(title="REPORTE")
                        if "Sheet" in workbook.sheetnames:
                            workbook.remove(workbook["Sheet"])

                        rojo = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
                        verde = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")
                        amarillo = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
                        blanco = Font(color="FFFFFF", bold=True)
                        borde = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
                        centro = Alignment(horizontal="center", vertical="center")

                        try:
                            logo = Image("LOGO.jpeg")
                            logo.width = 130
                            logo.height = 130
                            hoja.add_image(logo, "A1")
                        except:
                            pass

                        hoja.merge_cells("C7:H7")
                        banco_nombre = banco.upper()
                        hoja["C7"] = f"{banco_nombre} - REPORTE DE CONCILIACIÓN"
                        hoja["C7"].font = Font(bold=True, size=14)
                        hoja["C7"].alignment = centro

                        def crear_tabla(titulo, dataframe, fila_inicio, color_total):
                            hoja.merge_cells(start_row=fila_inicio, start_column=1, end_row=fila_inicio, end_column=10)
                            titulo_cell = hoja.cell(row=fila_inicio, column=1)
                            titulo_cell.value = titulo
                            titulo_cell.fill = rojo
                            titulo_cell.font = blanco
                            titulo_cell.alignment = centro

                            headers = [
                                "FECHA", "REFERENCIA", "DESCRIPCIÓN", "DESCRIPCIÓN ORIGINAL",
                                "MONTO BS", "TASA BCV", "MONTO USD", 
                                "PROVEEDOR (iPago)", "TIPO EGRESO (iPago)", "TIPO PAGO (iPago)"
                            ]
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
                                hoja.cell(row=fila_data, column=1).value = row.get("FECHA", "")
                                hoja.cell(row=fila_data, column=2).value = row.get("REFERENCIA", "")
                                hoja.cell(row=fila_data, column=3).value = row.get("DESCRIPCIÓN", "")
                                hoja.cell(row=fila_data, column=4).value = row.get("DESCRIPCION_ORIGINAL", "")
                                hoja.cell(row=fila_data, column=5).value = row.get("MONTO BS", 0)
                                hoja.cell(row=fila_data, column=6).value = row.get("TASA BCV", 0)
                                hoja.cell(row=fila_data, column=7).value = row.get("MONTO USD", 0)
                                hoja.cell(row=fila_data, column=8).value = row.get("PROVEEDOR_IPAGO", row.get("STATUS", ""))
                                hoja.cell(row=fila_data, column=9).value = row.get("OBSERVACIÓN", "")
                                hoja.cell(row=fila_data, column=10).value = row.get("TIPO_PAGO", "")

                                hoja.cell(row=fila_data, column=5).number_format = '#,##0.00'
                                hoja.cell(row=fila_data, column=6).number_format = '#,##0.0000'
                                hoja.cell(row=fila_data, column=7).number_format = '$#,##0.00'

                                for col in range(1, 11):
                                     hoja.cell(row=fila_data, column=col).border = borde
                                fila_data += 1

                            total_cell = hoja.cell(row=fila_data, column=4)
                            total_cell.value = f"TOTAL {titulo}"
                            total_cell.font = Font(bold=True)

                            total_bs_cell = hoja.cell(row=fila_data, column=5)
                            total_bs_cell.value = dataframe["MONTO BS"].sum() if not dataframe.empty else 0
                            total_bs_cell.number_format = '#,##0.00'
                            total_bs_cell.fill = color_total

                            monto_total = hoja.cell(row=fila_data, column=7)
                            monto_total.value = dataframe["MONTO USD"].sum() if not dataframe.empty else 0
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
                            adjusted_width = min(max_length + 5, 50)
                            hoja.column_dimensions[columna_letra].width = adjusted_width

                    output.seek(0)

                    st.download_button(
                        label="📥 Descargar Excel Clasificado (con Tasas BCV e iPago)",
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
                            df_tasas = pd.DataFrame([
                                {"FECHA": f, "TASA BCV": t} 
                                for f, t in todas_tasas.items()
                            ]).sort_values("FECHA")
                            st.dataframe(df_tasas, use_container_width=True)

            except Exception as e:
                st.error(f"❌ Error general: {str(e)}")
                st.code(str(e))
        else:
            st.markdown("""
            ### 👋 Cruce de Información (Monobanco)
            
            Carga un estado de cuenta bancario y el archivo maestro de iPago en el menú de la izquierda para clasificar transacciones y realizar la conciliación individual.
            
            **Características:**
            - Soporte simultáneo para múltiples cuentas.
            - Cálculo automático de saldo consolidado en Bolívares (VES) y Dólares (USD).
            - Cruce inteligente y trazabilidad con iPago.
            - Generación de reportes de cierre en formato Excel profesional.
            """)
