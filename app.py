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
import unicodedata  # 🔥 Para normalizar texto y eliminar espacios ocultos

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

    archivo_ipago = st.file_uploader(
        "📂 Cargar archivo iPago",
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
                # Intentar leer con xlrd
                return pd.read_excel(archivo, sheet_name=0, header=None, engine='xlrd')
            except Exception as e:
                # Si falla, intentar leer como texto/CSV
                st.warning(f"⚠️ Error leyendo como Excel, intentando como texto: {str(e)}")
                archivo.seek(0)
                # Leer como texto y procesar manualmente
                contenido = archivo.read().decode('utf-8', errors='ignore')
                lineas = contenido.split('\n')
                # Crear DataFrame con las líneas
                datos = []
                for linea in lineas:
                    if linea.strip():
                        # Dividir por tabuladores o espacios múltiples
                        partes = linea.split('\t')
                        if len(partes) == 1:
                            partes = [p for p in linea.split(' ') if p.strip()]
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
# FUNCIONES ORIGINALES (NO MODIFICADAS)
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

        numero = float(valor)

        if "." not in valor_original and "," not in valor_original and numero >= 100000:
            numero = numero / 100

        return numero

    except Exception:
        return None

# =========================================================
# CALCULAR USD SEGÚN TASA
# =========================================================

def calcular_usd(monto_bs, tasa):
    try:
        if monto_bs is None or tasa is None or tasa == 0:
            return None
        return round(abs(monto_bs) / abs(tasa), 2)
    except:
        return None

# =========================================================
# 🔥 DETECTAR COMISIONES - VERSIÓN MEJORADA CON PALABRAS CLAVE
# =========================================================

def es_comision(texto, proveedor=None):
    """
    Detecta si un movimiento es una comisión bancaria.
    
    REGLAS:
    - Si tiene proveedor asociado → NO es comisión bancaria (es un pago a terceros)
    - Si es pago a personal (nómina, comisiones de ventas) → NO es comisión bancaria
    - Solo son comisiones bancarias: cargos del banco (ITF, mantenimiento, comisión por transferencia, etc.)
    """
    texto = normalizar_texto(texto).strip()
    texto_upper = texto.upper()
    
    # 🔥 REGLA 1: Si tiene proveedor asociado, NO es comisión bancaria
    if proveedor and str(proveedor).strip():
        return False
    
    # 🔥 REGLA 2: COMISIONES PAGADAS A PERSONAL = EGRESO (no comisión bancaria)
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
    
    # 🔥 REGLA 3: NUNCA SON COMISIONES BANCARIAS
    if any(x in texto for x in [
        "pago a proveedores",
        "pago de nomina",
        "nomina",
        "transf entre ctas",
        "transferencia a terceros",
        "pago movil comercial"
    ]):
        return False
    
    # 🔥 REGLA 4: SOLO SON COMISIONES BANCARIAS si coinciden con estas palabras
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
    
    # Verificar si coincide con alguna comisión bancaria
    for patron in palabras_comision_bancaria:
        if patron in texto:
            return True
    
    # Si contiene "comision" pero no coincide con las reglas anteriores, NO es comisión bancaria
    if "comision" in texto or "comisión" in texto:
        return False
    
    return False

# =========================================================
# DETECTOR DE BANCO CORREGIDO - CON EL CAMBIO SOLICITADO
# =========================================================

def detectar_banco(nombre_archivo):
    nombre = nombre_archivo.upper()

    if "TESORO" in nombre or "TESORERIA" in nombre or "TES" in nombre:
        return "tesoro"
    elif "BANESCO" in nombre or re.match(r"^J\d+", nombre_archivo):
        return "banesco"
    elif "BANCAMIGA" in nombre or "MOVIMIENTOS_" in nombre:
        return "bancamiga"
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
# 🔥 FUNCIÓN MEJORADA: ENRIQUECER EGRESOS CON IPAGO (CRUCE FLEXIBLE)
# =========================================================

def enriquecer_egresos_con_ipago(df_egresos, df_ipago):
    """
    Enriquece los egresos del banco con datos de iPago.
    Usa múltiples estrategias de cruce:
    1. Coincidencia exacta de referencia
    2. Coincidencia parcial (quitando ceros o X)
    3. Coincidencia por monto + fecha
    """
    if df_ipago is None or df_ipago.empty:
        return df_egresos
    
    # Hacer una copia para no modificar el original
    df_resultado = df_egresos.copy()
    
    # 🔥 NORMALIZAR REFERENCIAS DEL BANCO
    df_resultado["REFERENCIA_NORM"] = (
        df_resultado["REFERENCIA"]
        .astype(str)
        .str.replace(".0", "", regex=False)
        .str.strip()
    )
    
    # 🔥 NORMALIZAR REFERENCIAS DE IPAGO
    df_ipago["Referencia_Norm"] = (
        df_ipago["Referencia"]
        .astype(str)
        .str.replace(".0", "", regex=False)
        .str.strip()
    )
    
    # 🔥 CREAR VARIANTES DE REFERENCIA PARA CRUCE FLEXIBLE
    def generar_variantes_ref(ref):
        ref = str(ref).strip()
        if not ref or ref == "nan":
            return set()
        
        variantes = set()
        variantes.add(ref)  # Original
        
        # Quitar ceros a la izquierda
        ref_sin_ceros = ref.lstrip('0')
        if ref_sin_ceros != ref and ref_sin_ceros:
            variantes.add(ref_sin_ceros)
        
        # Quitar X al final
        if ref.endswith('X'):
            ref_sin_x = ref[:-1]
            if ref_sin_x:
                variantes.add(ref_sin_x)
                # También sin ceros
                ref_sin_x_sin_ceros = ref_sin_x.lstrip('0')
                if ref_sin_x_sin_ceros:
                    variantes.add(ref_sin_x_sin_ceros)
        
        # Si tiene X pero no es el final (caso raro)
        if 'X' in ref and not ref.endswith('X'):
            ref_sin_x = ref.replace('X', '')
            if ref_sin_x:
                variantes.add(ref_sin_x)
                ref_sin_x_sin_ceros = ref_sin_x.lstrip('0')
                if ref_sin_x_sin_ceros:
                    variantes.add(ref_sin_x_sin_ceros)
        
        # 🔥 NUEVO PARA VENEZUELA: Si la referencia tiene 11 dígitos y comienza con 0
        if len(ref) >= 10 and ref.startswith('0'):
            ref_sin_cero_inicial = ref[1:]
            if ref_sin_cero_inicial:
                variantes.add(ref_sin_cero_inicial)
                variantes.add(ref_sin_cero_inicial.lstrip('0'))
        
        # 🔥 NUEVO PARA VENEZUELA: Si la referencia tiene formato numérico con puntos
        if '.' in ref:
            ref_sin_puntos = ref.replace('.', '')
            if ref_sin_puntos:
                variantes.add(ref_sin_puntos)
                variantes.add(ref_sin_puntos.lstrip('0'))
        
        return variantes
    
    # 🔥 CREAR DICCIONARIO DE IPAGO CON TODAS LAS VARIANTES
    ipago_dict = {}
    for _, row in df_ipago.iterrows():
        ref_original = str(row.get("Referencia", "")).strip()
        
        # Si la referencia es NaN o está vacía, usar monto como clave
        if not ref_original or ref_original == "nan":
            continue
        
        # Generar todas las variantes de esta referencia
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
    
    # 🔥 ENRIQUECER CADA EGRESO
    for idx, row in df_resultado.iterrows():
        ref_banco = str(row.get("REFERENCIA", "")).strip()
        monto_banco = float(row.get("MONTO BS", 0))
        fecha_banco = str(row.get("FECHA", ""))
        
        # GENERAR VARIANTES DE LA REFERENCIA DEL BANCO
        variantes_banco = generar_variantes_ref(ref_banco)
        
        # BUSCAR COINCIDENCIA POR REFERENCIA
        coincide_ref = False
        datos_encontrados = None
        
        for variante in variantes_banco:
            if variante in ipago_dict:
                datos_encontrados = ipago_dict[variante]
                coincide_ref = True
                break
        
        # SI NO COINCIDE POR REFERENCIA, INTENTAR POR MONTO + FECHA
        if not coincide_ref:
            # Buscar en iPago por monto similar (con margen de 1%)
            for clave, datos in ipago_dict.items():
                monto_ipago = float(datos.get("MONTO_IPAGO", 0))
                if monto_ipago > 0:
                    diferencia = abs(monto_banco - monto_ipago) / max(monto_banco, monto_ipago)
                    if diferencia < 0.01:  # 1% de margen
                        datos_encontrados = datos
                        coincide_ref = True
                        break
        
        # APLICAR DATOS ENCONTRADOS
        if datos_encontrados:
            df_resultado.at[idx, "STATUS"] = datos_encontrados["PROVEEDOR"]
            df_resultado.at[idx, "OBSERVACIÓN"] = datos_encontrados["TIPO_EGRESO"]
            df_resultado.at[idx, "TIPO_PAGO"] = datos_encontrados["TIPO_PAGO"]
            df_resultado.at[idx, "PROVEEDOR_IPAGO"] = datos_encontrados["PROVEEDOR"]
            df_resultado.at[idx, "REFERENCIA_IPAGO"] = datos_encontrados.get("REFERENCIA_ORIGINAL", "")
            
            # 🔥 Reemplazar descripción con la de iPago
            descripcion_ipago = datos_encontrados["DESCRIPCION_IPAGO"]
            if descripcion_ipago:
                df_resultado.at[idx, "DESCRIPCIÓN"] = descripcion_ipago
                df_resultado.at[idx, "DESCRIPCION_ORIGINAL"] = row.get("DESCRIPCIÓN", "")
            
            # 🔥 Si es comisión, marcarlo como tal
            tipo = str(datos_encontrados["TIPO_EGRESO"]).upper()
            desc = str(datos_encontrados["DESCRIPCION_IPAGO"]).upper()
            if "COMISION" in tipo or "COMISION" in desc:
                df_resultado.at[idx, "ES_COMISION"] = True
            else:
                df_resultado.at[idx, "ES_COMISION"] = False
        else:
            # Si no hay coincidencia, mantener los valores actuales
            df_resultado.at[idx, "STATUS"] = "SIN DATOS IPAGO"
            df_resultado.at[idx, "OBSERVACIÓN"] = "SIN CONCORDANCIA"
            df_resultado.at[idx, "TIPO_PAGO"] = ""
            df_resultado.at[idx, "PROVEEDOR_IPAGO"] = ""
            df_resultado.at[idx, "ES_COMISION"] = False
    
    # Eliminar columnas auxiliares
    df_resultado = df_resultado.drop(
        columns=["REFERENCIA_NORM"], 
        errors="ignore"
    )
    
    return df_resultado

# =========================================================
# PROCESAR BANESCO
# =========================================================

def procesar_banesco(df):
    st.info("Procesando Banesco...")
    try:
        df.columns = ["FECHA", "REFERENCIA", "DESCRIPCION", "MONTO_RAW", "BALANCE"]
        df.columns = [str(c).strip().upper() for c in df.columns]

        rename_map = {}
        for col in df.columns:
            c = str(col).lower()
            if "fecha" in c:
                rename_map[col] = "FECHA"
            elif "referencia" in c:
                rename_map[col] = "REFERENCIA"
            elif "descrip" in c:
                rename_map[col] = "DESCRIPCION"
            elif "monto" in c:
                rename_map[col] = "MONTO_RAW"

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
        st.dataframe(df.head())
        return df

    except Exception as e:
        st.error(f"Error Banesco: {str(e)}")
        return pd.DataFrame()

# =========================================================
# PROCESAR PROVINCIAL - VERSIÓN MEJORADA PARA FORMATO ESPECÍFICO
# =========================================================

def procesar_provincial(df):
    """
    Procesa archivo de Provincial con formato específico.
    El archivo tiene un formato de texto con columnas:
    F. Operación | F. Valor | Código | Nº. Doc. | Concepto | Importe | Oficina
    """
    st.info("🔍 Procesando archivo de Provincial (formato especial)...")
    
    try:
        # Mostrar información del archivo
        st.write("📊 **Información del archivo:**")
        st.write(f"- Número de filas: {len(df)}")
        st.write(f"- Número de columnas: {len(df.columns)}")
        
        # Mostrar primeras filas para debug
        st.write("👁️ **Primeras 15 filas del archivo:**")
        st.dataframe(df.head(15))
        
        # Buscar la fila que contiene los encabezados
        encabezado_idx = None
        for i in range(min(30, len(df))):
            fila = df.iloc[i]
            # Convertir TODOS los valores a string para evitar errores
            fila_str = [str(val) for val in fila.tolist()]
            texto_fila = " ".join(fila_str).upper()
            
            # Buscar columnas que contengan "F. Operación" o "Concepto" o "Importe"
            if "F. OPERACIÓN" in texto_fila or "F. VALOR" in texto_fila or "CONCEPTO" in texto_fila:
                encabezado_idx = i
                break
        
        if encabezado_idx is None:
            st.error("❌ No se encontró la fila de encabezados en el archivo Provincial.")
            return pd.DataFrame()
        
        st.write(f"✅ Encabezados encontrados en la fila {encabezado_idx}")
        
        # Obtener los encabezados
        headers = df.iloc[encabezado_idx].astype(str).str.strip().tolist()
        st.write("📋 **Encabezados detectados:**", headers)
        
        # Limpiar y mapear encabezados
        rename_map = {}
        for col in headers:
            col_clean = str(col).strip().upper()
            if "F. OPERACIÓN" in col_clean or "FECHA" in col_clean:
                rename_map[col] = "FECHA"
            elif "F. VALOR" in col_clean:
                rename_map[col] = "FECHA_VALOR"
            elif "CÓDIGO" in col_clean or "CODIGO" in col_clean:
                rename_map[col] = "CODIGO"
            elif "Nº. DOC" in col_clean or "NRO DOC" in col_clean or "DOC" in col_clean:
                rename_map[col] = "REFERENCIA"
            elif "CONCEPTO" in col_clean:
                rename_map[col] = "DESCRIPCION"
            elif "IMPORTE" in col_clean:
                rename_map[col] = "MONTO"
            elif "OFICINA" in col_clean:
                rename_map[col] = "OFICINA"
        
        st.write("📋 **Mapeo de columnas:**", rename_map)
        
        # Asignar encabezados al DataFrame
        df.columns = headers
        df = df.iloc[encabezado_idx + 1:].reset_index(drop=True)
        
        # Renombrar columnas
        df = df.rename(columns=rename_map)
        
        # Verificar columnas necesarias
        if "FECHA" not in df.columns:
            # Intentar encontrar fecha en otra columna
            for col in df.columns:
                if "FECHA" in str(col).upper():
                    df = df.rename(columns={col: "FECHA"})
                    break
        
        if "FECHA" in df.columns:
            # Procesar fechas - convertir a string primero
            df["FECHA"] = df["FECHA"].astype(str).str.strip()
            # Eliminar filas con fechas vacías o que sean encabezados
            df = df[~df["FECHA"].str.contains("FECHA|SALDO|Período", case=False, na=False)]
            # Eliminar filas con fechas que sean números o NaN
            df = df[df["FECHA"].str.match(r'^\d{2}-\d{2}-\d{4}$', na=False)]
            
            # Convertir fechas (formato DD-MM-YYYY)
            df["FECHA"] = pd.to_datetime(df["FECHA"], dayfirst=True, errors="coerce")
            df = df[df["FECHA"].notna()]
        else:
            st.error("❌ No se encontró columna FECHA en el archivo Provincial.")
            return pd.DataFrame()
        
        # Procesar el monto
        if "MONTO" in df.columns:
            # Limpiar el monto (quitar espacios, puntos, comas) - convertir a string primero
            df["MONTO"] = df["MONTO"].astype(str).str.replace(" ", "", regex=False)
            df["MONTO"] = df["MONTO"].str.replace(".", "", regex=False)
            df["MONTO"] = df["MONTO"].str.replace(",", ".", regex=False)
            df["MONTO"] = df["MONTO"].str.replace("'", "", regex=False)
            
            # Eliminar filas con monto vacío o que no sean numéricos
            df = df[df["MONTO"].str.match(r'^[\d\.]+$', na=False)]
            
            # Convertir a numérico
            df["MONTO"] = pd.to_numeric(df["MONTO"], errors="coerce")
            
            # Si el monto es negativo, es un ND (débito), si es positivo es NC (crédito)
            df["TIPO"] = df["MONTO"].apply(lambda x: "NC" if x > 0 else "ND" if x < 0 else "")
            
            # Tomar valor absoluto
            df["MONTO"] = df["MONTO"].abs()
            
            # Eliminar filas con monto 0 o NaN
            df = df[df["MONTO"].notna()]
            df = df[df["MONTO"] > 0]
        else:
            st.error("❌ No se encontró columna MONTO en el archivo Provincial.")
            return pd.DataFrame()
        
        # Asegurar que existe columna REFERENCIA
        if "REFERENCIA" not in df.columns:
            df["REFERENCIA"] = ""
        else:
            df["REFERENCIA"] = df["REFERENCIA"].astype(str).str.strip()
            # Limpiar referencias (quitar comillas simples)
            df["REFERENCIA"] = df["REFERENCIA"].str.replace("'", "", regex=False)
        
        # Asegurar que existe columna DESCRIPCION
        if "DESCRIPCION" not in df.columns:
            df["DESCRIPCION"] = ""
        else:
            df["DESCRIPCION"] = df["DESCRIPCION"].astype(str).str.strip()
        
        # 🔥 DETECTAR COMISIONES DE PROVINCIAL - CREAR COLUMNA ESPECIAL
        df["ES_COMISION"] = df["DESCRIPCION"].str.contains("COMIS", case=False, na=False)
        
        # 🔥 DEBUG: Mostrar cuántas comisiones se detectaron
        num_comisiones = df["ES_COMISION"].sum()
        st.info(f"💳 Se detectaron {num_comisiones} comisiones en el archivo Provincial")
        
        # Mostrar las comisiones detectadas
        if num_comisiones > 0:
            st.write("📋 **Comisiones detectadas:**")
            st.dataframe(df[df["ES_COMISION"] == True][["FECHA", "REFERENCIA", "DESCRIPCION", "MONTO"]])
        
        # Seleccionar solo las columnas necesarias
        df_resultado = df[["FECHA", "REFERENCIA", "DESCRIPCION", "TIPO", "MONTO", "ES_COMISION"]].copy()
        
        # Mostrar resultados
        st.success(f"✅ Provincial OK: {len(df_resultado)} movimientos detectados")
        st.dataframe(df_resultado.head(10))
        
        return df_resultado
        
    except Exception as e:
        st.error(f"❌ Error procesando Provincial: {str(e)}")
        import traceback
        st.code(traceback.format_exc())
        return pd.DataFrame()

# =========================================================
# PROCESAR BNC
# =========================================================

def procesar_bnc(df):
    st.info("Procesando archivo BNC...")
    encabezado = None

    for i in range(min(30, len(df))):
        fila = df.iloc[i].fillna("").astype(str)
        texto = " ".join(fila.tolist()).lower()
        if "fecha" in texto and ("descripcion" in texto or "descripción" in texto):
            encabezado = i
            break

    if encabezado is None:
        st.error("No se encontró encabezado válido en BNC")
        return pd.DataFrame()

    headers = []
    for idx, col in enumerate(df.iloc[encabezado]):
        col = str(col).strip().replace("\n", " ")
        if col == "" or col.lower() == "nan":
            col = f"COLUMNA_{idx}"
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
        if "fecha" in col_str:
            rename_map[col] = "FECHA"
        elif "descripcion" in col_str or "descripción" in col_str or "concepto" in col_str:
            rename_map[col] = "DESCRIPCION"
        elif "referencia" in col_str:
            rename_map[col] = "REFERENCIA"
        elif "credito" in col_str or "haber" in col_str:
            rename_map[col] = "CREDITO"
        elif "debito" in col_str or "debe" in col_str:
            rename_map[col] = "DEBITO"

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

    st.success(f"Registros BNC: {len(df)}")
    st.dataframe(df.head())
    return df

# =========================================================
# PROCESAR TESORO
# =========================================================

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
            st.error("No se encontró encabezado válido en Tesoro")
            return pd.DataFrame()

        df.columns = df.iloc[encabezado]
        df = df.iloc[encabezado + 1:].reset_index(drop=True)
        df.columns = [str(c).strip() for c in df.columns]

        rename_map = {}
        for col in df.columns:
            c = str(col).strip().lower()
            if "fecha" in c:
                rename_map[col] = "FECHA"
            elif "referencia" in c:
                rename_map[col] = "REFERENCIA"
            elif "concepto" in c:
                rename_map[col] = "DESCRIPCION"
            elif "débito" in c or "debito" in c:
                rename_map[col] = "DEBITO"
            elif "crédito" in c or "credito" in c:
                rename_map[col] = "CREDITO"
            elif "código" in c or "codigo" in c:
                rename_map[col] = "TIPO"

        df = df.rename(columns=rename_map)

        if "FECHA" not in df.columns:
            st.error("No existe columna FECHA")
            return pd.DataFrame()

        df["FECHA"] = pd.to_datetime(df["FECHA"], dayfirst=True, errors="coerce")
        df = df[df["FECHA"].notna()]

        def limpiar_numero(valor):
            valor = str(valor).replace(".", "").replace(",", ".")
            try:
                return float(valor)
            except:
                return 0

        df["CREDITO"] = df.get("CREDITO", 0).apply(limpiar_numero) if "CREDITO" in df.columns else 0
        df["DEBITO"] = df.get("DEBITO", 0).apply(limpiar_numero) if "DEBITO" in df.columns else 0

        df["MONTO"] = df["CREDITO"] - df["DEBITO"]
        df["TIPO"] = df["MONTO"].apply(lambda x: "NC" if x > 0 else "ND")
        df["MONTO"] = df["MONTO"].abs()
        df = df[df["MONTO"] > 0]

        df = df[["FECHA", "REFERENCIA", "DESCRIPCION", "TIPO", "MONTO"]]
        st.success(f"Tesoro OK: {len(df)} registros")
        st.dataframe(df.head())
        return df

    except Exception as e:
        st.error(f"Error Tesoro: {str(e)}")
        return pd.DataFrame()

# =========================================================
# PROCESAR BANCAMIGA - CORREGIDO
# =========================================================

def procesar_bancamiga(df):
    st.info("Procesando Bancamiga...")
    try:
        encabezado = None
        for i in range(min(15, len(df))):
            fila = df.iloc[i].fillna("").astype(str)
            texto = " ".join(fila.tolist()).lower()
            if "fecha" in texto and "referencia" in texto and "concepto" in texto:
                encabezado = i
                break

        if encabezado is not None:
            df.columns = df.iloc[encabezado]
            df = df.iloc[encabezado + 1:].reset_index(drop=True)
            df.columns = [str(c).strip() for c in df.columns]

        if isinstance(df.columns, pd.MultiIndex):
            df.columns = [col[-1] for col in df.columns]

        df = df.rename(columns={
            "Fecha": "FECHA",
            "Referencia": "REFERENCIA",
            "Concepto": "DESCRIPCION",
            "Débito": "DEBITO",
            "Crédito": "CREDITO"
        })

        if "FECHA" not in df.columns:
            st.error("No se encontró columna FECHA")
            return pd.DataFrame()

        df["FECHA"] = pd.to_datetime(df["FECHA"], dayfirst=True, errors="coerce")
        df = df[df["FECHA"].notna()]

        def limpiar_numero_bancamiga(valor):
            if pd.isna(valor):
                return 0
            try:
                if isinstance(valor, (int, float)):
                    numero = float(valor)
                    if isinstance(valor, int) and numero >= 100000:
                        numero = numero / 100
                    return numero

                valor_original = str(valor).strip()
                valor = valor_original.replace(" ", "").replace("$", "").replace("Bs", "").replace("€", "")
                if valor == "":
                    return 0

                if "." in valor and "," in valor:
                    valor = valor.replace(".", "").replace(",", ".")
                elif "," in valor:
                    valor = valor.replace(",", ".")

                numero = float(valor)
                if "." not in valor_original and "," not in valor_original and numero >= 100000:
                    numero = numero / 100
                return numero
            except:
                return 0

        df["CREDITO"] = df["CREDITO"].apply(limpiar_numero_bancamiga)
        df["DEBITO"] = df["DEBITO"].apply(limpiar_numero_bancamiga)

        df["MONTO"] = df["CREDITO"] - df["DEBITO"]
        df["TIPO"] = df["MONTO"].apply(lambda x: "NC" if x > 0 else "ND")
        df["MONTO"] = df["MONTO"].abs()
        df = df[df["MONTO"] > 0]

        df = df[["FECHA", "REFERENCIA", "DESCRIPCION", "TIPO", "MONTO"]]
        st.success(f"Bancamiga OK: {len(df)} registros")
        st.dataframe(df.head())
        return df

    except Exception as e:
        st.error(f"Error Bancamiga: {str(e)}")
        return pd.DataFrame()

# =========================================================
# OBTENER TASA BCV
# =========================================================

@st.cache_data(ttl=3600)
def obtener_tasa_bcv_fecha(fecha_obj):
    tasas_bcv_local = {
        "01/06/2026": 554.4258,
        "02/06/2026": 557.9741,
        "03/06/2026": 558.6436,
        "04/06/2026": 560.3753,
        "05/06/2026": 563.2892,
        "06/06/2026": 567.6828,
        "07/06/2026": 567.6828,
        "08/06/2026": 567.6828,
        "09/06/2026": 567.6828,
        "10/06/2026": 572.6784,
        "11/06/2026": 577.5461,
        "12/06/2026": 582.6862,
        "13/06/2026": 587.4059,
        "14/06/2026": 587.4059,
        "15/06/2026": 587.4059,
        "16/06/2026": 592.5163,
        "17/06/2026": 596.7824,
        "18/06/2026": 602.3324,
        "19/06/2026": 607.3919,
        "20/06/2026": 612.4332,
        "21/06/2026": 612.4332,
        "22/06/2026": 612.4332,
        "23/06/2026": 617.6388,
    }
    fecha_str = fecha_obj.strftime("%d/%m/%Y")
    return tasas_bcv_local.get(fecha_str, None)

def obtener_tasa_por_fecha(fecha_obj, usar_api=False):
    return obtener_tasa_bcv_fecha(fecha_obj)

# =========================================================
# CONVERTIR A FORMATO MERCANTIL - INCLUYE FLAG DE COMISIONES
# =========================================================

def convertir_a_formato_mercantil(df, banco):
    """Convierte DataFrame de otros bancos al formato que espera procesar_archivo"""
    datos_convertidos = []
    
    for idx, fila in df.iterrows():
        try:
            fecha = fila.get("FECHA", "")
            if pd.isna(fecha):
                continue
            
            if isinstance(fecha, (pd.Timestamp, datetime)):
                fecha_str = fecha.strftime("%d/%m/%Y")
            else:
                fecha_str = str(fecha)
            
            tipo = fila.get("TIPO", "") or ""
            descripcion = fila.get("DESCRIPCION", "") or ""
            referencia = fila.get("REFERENCIA", "") or ""
            monto = fila.get("MONTO", 0) or 0
            
            # 🔥 OBTENER FLAG DE COMISIÓN
            es_comision = fila.get("ES_COMISION", False)
            if isinstance(es_comision, (bool, np.bool_)):
                es_comision = bool(es_comision)
            else:
                es_comision = False
            
            fila_convertida = [
                "",           # col0
                "",           # col1  
                "",           # col2
                fecha_str,    # col3 - FECHA
                referencia,   # col4 - REFERENCIA
                tipo,         # col5 - TIPO (NC/ND)
                descripcion,  # col6 - DESCRIPCION
                monto,        # col7 - MONTO BS
                "",           # col8
                es_comision,  # col9 - ES_COMISION (flag para identificar comisiones)
            ]
            datos_convertidos.append(fila_convertida)
            
        except Exception as e:
            continue
    
    df_convertido = pd.DataFrame(datos_convertidos)
    return df_convertido if len(df_convertido) > 0 else pd.DataFrame()

# =========================================================
# 🔥 PROCESAR VENEZUELA - VERSIÓN MEJORADA (SIN FILTROS EXCESIVOS)
# =========================================================

def procesar_venezuela_simple(df):
    """Procesa el archivo del BDV usando índices fijos (sin encabezados)"""
    st.info("🔍 Procesando Banco de Venezuela (MODO SIMPLE)...")
    
    try:
        # Mostrar información del archivo
        st.write("📊 **Información del archivo:**")
        st.write(f"- Número de filas: {len(df)}")
        st.write(f"- Número de columnas: {len(df.columns)}")
        
        # Mostrar primeras filas
        st.write("👁️ **Primeras 10 filas del archivo (sin encabezados):**")
        st.dataframe(df.head(10))
        
        # Índices fijos para el formato BDV
        col_fecha = 3
        col_ref = 1
        col_desc = 2
        col_tipo = 4
        col_credito = 5
        col_debito = 6
        
        movimientos = []
        filas_procesadas = 0
        
        # Empezar desde la fila 1 (saltar encabezados)
        for idx in range(1, len(df)):
            try:
                fila = df.iloc[idx]
                
                # Verificar que existe fecha
                if pd.isna(fila[col_fecha]):
                    continue
                
                # Obtener fecha
                fecha_raw = str(fila[col_fecha]).strip()
                
                # Intentar parsear fecha
                fecha_val = None
                try:
                    fecha_val = pd.to_datetime(fecha_raw, format="%d/%m/%Y", errors="coerce")
                except:
                    pass
                
                if pd.isna(fecha_val):
                    try:
                        fecha_val = pd.to_datetime(fecha_raw, dayfirst=True, errors="coerce")
                    except:
                        pass
                
                if pd.isna(fecha_val):
                    continue
                
                # Obtener datos
                referencia = str(fila[col_ref]).strip() if pd.notna(fila[col_ref]) else ""
                descripcion = str(fila[col_desc]).strip() if pd.notna(fila[col_desc]) else ""
                tipo_mov = str(fila[col_tipo]).strip().upper() if pd.notna(fila[col_tipo]) else ""
                
                # 🔥 SOLO FILTRAR POR SALDO INICIAL/FINAL EXPLÍCITO
                desc_upper = descripcion.upper()
                if desc_upper in ["SALDO INICIAL", "SALDO FINAL", "TOTALES"]:
                    continue
                
                # Procesar crédito
                val_credito = 0
                if pd.notna(fila[col_credito]):
                    clean_cred = str(fila[col_credito]).strip()
                    clean_cred = clean_cred.replace(".", "").replace(",", ".")
                    try:
                        val_credito = float(clean_cred) if clean_cred else 0
                    except:
                        pass
                
                # Procesar débito
                val_debito = 0
                if pd.notna(fila[col_debito]):
                    clean_deb = str(fila[col_debito]).strip()
                    clean_deb = clean_deb.replace(".", "").replace(",", ".")
                    try:
                        val_debito = float(clean_deb) if clean_deb else 0
                    except:
                        pass
                
                # Determinar tipo y monto
                monto = 0
                tipo = ""
                
                if tipo_mov == "NC":
                    monto = abs(val_credito)
                    tipo = "NC"
                elif tipo_mov == "ND":
                    monto = abs(val_debito)
                    tipo = "ND"
                else:
                    # Si no tiene tipo definido, determinar por crédito/débito
                    if abs(val_credito) > 0:
                        monto = abs(val_credito)
                        tipo = "NC"
                    elif abs(val_debito) > 0:
                        monto = abs(val_debito)
                        tipo = "ND"
                    else:
                        continue
                
                if monto <= 0:
                    continue
                
                filas_procesadas += 1
                
                # Mostrar primeras 5 filas procesadas
                if filas_procesadas <= 5:
                    st.write(f"✅ **Fila {idx} procesada:** Fecha={fecha_val.strftime('%d/%m/%Y')}, Ref={referencia}, Tipo={tipo}, Monto={monto:,.2f}, Desc={descripcion[:50]}")
                
                movimientos.append({
                    "FECHA": fecha_val.strftime("%d/%m/%Y"),
                    "FECHA_OBJ": fecha_val,
                    "REFERENCIA": referencia,
                    "DESCRIPCION": descripcion,
                    "TIPO": tipo,
                    "MONTO": monto
                })
                
            except Exception as e:
                continue
        
        st.write(f"📊 **Filas procesadas exitosamente:** {filas_procesadas}")
        
        df_resultado = pd.DataFrame(movimientos)
        
        if df_resultado.empty:
            st.error("❌ No se encontraron movimientos válidos en el archivo de Venezuela.")
            return pd.DataFrame()
        
        st.success(f"✅ Venezuela OK: {len(df_resultado)} movimientos detectados")
        st.dataframe(df_resultado.head(10))
        return df_resultado
        
    except Exception as e:
        st.error(f"❌ Error general en procesar_venezuela_simple: {str(e)}")
        import traceback
        st.code(traceback.format_exc())
        return pd.DataFrame()


def convertir_venezuela_a_formato_mercantil(df):
    """Convierte DataFrame de Venezuela al formato Mercantil - SIN AFECTAR A MERCANTIL"""
    datos_convertidos = []
    
    for idx, fila in df.iterrows():
        try:
            # Obtener fecha
            fecha = None
            if "FECHA_OBJ" in fila and pd.notna(fila["FECHA_OBJ"]):
                fecha = fila["FECHA_OBJ"]
            elif "FECHA" in fila and pd.notna(fila["FECHA"]):
                fecha = fila["FECHA"]
            else:
                continue
            
            if pd.isna(fecha):
                continue
            
            # Convertir a string
            if isinstance(fecha, (pd.Timestamp, datetime)):
                fecha_str = fecha.strftime("%d/%m/%Y")
            else:
                try:
                    fecha_dt = pd.to_datetime(fecha, dayfirst=True, errors="coerce")
                    if pd.notna(fecha_dt):
                        fecha_str = fecha_dt.strftime("%d/%m/%Y")
                    else:
                        fecha_str = str(fecha)
                except:
                    fecha_str = str(fecha)
            
            tipo = fila.get("TIPO", "") or ""
            descripcion = fila.get("DESCRIPCION", "") or ""
            referencia = fila.get("REFERENCIA", "") or ""
            monto = fila.get("MONTO", 0) or 0
            
            fila_convertida = [
                "",           # col0
                "",           # col1  
                "",           # col2
                fecha_str,    # col3 - FECHA
                referencia,   # col4 - REFERENCIA
                tipo,         # col5 - TIPO (NC/ND)
                descripcion,  # col6 - DESCRIPCION
                monto,        # col7 - MONTO BS
                "",           # col8
                False,        # col9 - ES_COMISION (siempre False para Venezuela)
            ]
            datos_convertidos.append(fila_convertida)
            
        except Exception as e:
            continue
    
    df_convertido = pd.DataFrame(datos_convertidos)
    return df_convertido if len(df_convertido) > 0 else pd.DataFrame()

# =========================================================
# 🔥 PROCESAMIENTO PRINCIPAL - CON DETECCIÓN DIRECTA PARA PROVINCIAL
# =========================================================

def procesar_archivo(df, usar_api=False, banco=""):
    """
    Procesa el archivo y clasifica movimientos en ingresos, egresos y comisiones.
    
    Args:
        df: DataFrame con los movimientos en formato Mercantil
        usar_api: Booleano para usar API BCV
        banco: Nombre del banco (para aplicar reglas específicas)
    """
    ingresos = []
    egresos = []
    comisiones = []
    registros_procesados = set()
    
    tipos_ingresos = ["NC", "C", "CREDITO", "ABONO"]
    tipos_egresos = ["ND", "D", "DEBITO", "DEBIT"]
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
                fecha = f"0{fecha_raw[0]}/{fecha_raw[1:3]}/{fecha_raw[3:]}"
            elif len(fecha_raw) == 8:
                fecha = f"{fecha_raw[0:2]}/{fecha_raw[2:4]}/{fecha_raw[4:]}"
            else:
                fecha = fecha_raw

            tipo = str(fila[5]).strip().upper()
            descripcion = str(fila[6]).strip()
            referencia = str(fila[4]).strip()

            monto_bs = convertir_monto(fila[7])
            if monto_bs is None or monto_bs == 0:
                continue
            
            fecha_obj = pd.to_datetime(fecha, dayfirst=True, errors="coerce")
            if pd.isna(fecha_obj):
                continue
            
            fecha_key = fecha_obj.strftime("%d/%m/%Y")
            if fecha_key in cache_tasas:
                tasa = cache_tasas[fecha_key]
            else:
                tasa = obtener_tasa_por_fecha(fecha_obj, usar_api) or 1.0
                if tasa is not None:
                    cache_tasas[fecha_key] = tasa

            monto_usd = calcular_usd(monto_bs, tasa)
            if monto_usd is None:
                continue

            texto = descripcion.upper()
            if texto in ["SALDO", "DESCRIPCION", "DESCRIPCIÓN", "REFERENCIA", "MOVIMIENTO", "FECHA", "SALDO INICIAL", "SALDO FINAL"]:
                continue

            registro = {
                "FECHA": fecha,
                "REFERENCIA": referencia,
                "DESCRIPCIÓN": descripcion,
                "MONTO BS": round(abs(monto_bs), 2) if monto_bs else 0,
                "TASA BCV": round(tasa, 4),
                "MONTO USD": monto_usd,
                "STATUS": "",
                "OBSERVACIÓN": "",
                "TIPO_PAGO": "",
                "PROVEEDOR_IPAGO": "",
                "DESCRIPCION_ORIGINAL": ""
            }

            clave = (fecha, referencia, descripcion, monto_usd, tipo)
            if clave in registros_procesados:
                continue
            registros_procesados.add(clave)

            # =========================================================
            # 🔥 REGLA ESPECIAL PARA PROVINCIAL - DETECCIÓN DIRECTA
            # =========================================================
            es_comision_provincial = False
            
            if banco == "provincial":
                descripcion_upper = descripcion.upper()
                # Detectar comisiones de Provincial por "COMIS" en la descripción
                if "COMIS" in descripcion_upper:
                    es_comision_provincial = True
                    st.write(f"🔍 Comisión detectada: {descripcion} - Monto: {monto_bs}")
                
                if es_comision_provincial:
                    comisiones.append(registro)
                    continue

            # =========================================================
            # 🔥 REGLA ESPECIAL PARA MERCANTIL - TODAS LAS COMISIONES
            # =========================================================
            es_comision_mercantil = False
            
            if banco == "mercantil":
                descripcion_upper = descripcion.upper()
                
                # 🔥 Lista completa de patrones de comisiones de Mercantil
                patrones_comision_mercantil = [
                    "OP.CRED.DIRT. CLTE-CLTE",
                    "OP CRED DIRT CLTE CLTE",
                    "COMISION PAGO MOVIL COMERCIAL",
                    "COMISION POR TRANSFERENCIA DE FONDOS",
                    "COMISION X PAGO DE NOMINAS",
                    "COMISION PAGO MOVIL COMERCIAL INTERBANCARIO",
                    "COMISION X PAGO DE NOMINAS MB",
                    "ITF",
                    "IMPUESTO A LAS TRANSACCIONES FINANCIERAS",
                    "CARGO BANCARIO",
                    "MANTENIMIENTO DE CUENTA",
                    "COMISION POR TRANSFERENCIA"
                ]
                
                for patron in patrones_comision_mercantil:
                    if patron in descripcion_upper:
                        es_comision_mercantil = True
                        break
                
                # Si es comisión de Mercantil, la clasificamos como tal
                if es_comision_mercantil:
                    comisiones.append(registro)
                    continue

            # 🔥 PASAR EL PROVEEDOR A LA FUNCIÓN es_comision
            proveedor = fila.get("Proveedor") if isinstance(fila, dict) else None
            if es_comision(descripcion, proveedor):
                comisiones.append(registro)
            elif tipo in tipos_ingresos:
                ingresos.append(registro)
            elif tipo in tipos_egresos:
                egresos.append(registro)

        except Exception as e:
            continue

    return ingresos, egresos, comisiones

# =========================================================
# INTERFAZ PRINCIPAL
# =========================================================

df_ipago = None

if archivo:
    st.info(f"📄 Archivo: **{archivo.name}** - {archivo.size/1024:.1f} KB")

    try:
        archivo.seek(0)
        primeros_bytes = archivo.read(100)
        archivo.seek(0)
        
        if b"<table" in primeros_bytes.lower():
            banco = "banesco"
        else:
            banco = detectar_banco(archivo.name)
        
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
                try:
                    df_raw = pd.read_html(archivo)[0]
                except:
                    archivo.seek(0)
                    if archivo.name.lower().endswith(".xls"):
                        df_raw = pd.read_excel(archivo, engine="xlrd", header=5)
                    else:
                        df_raw = pd.read_excel(archivo, engine="openpyxl", header=0)
                df_normalizado = procesar_bancamiga(df_raw)
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
            # 🔥 LEER SIN ENCABEZADOS (igual que Mercantil)
            df_raw = leer_excel_sin_encabezados(archivo)
            df_normalizado = procesar_venezuela_simple(df_raw)
            if df_normalizado.empty:
                st.stop()
            
            # Venezuela trabaja directamente con el dataframe normalizado
            df_original = convertir_venezuela_a_formato_mercantil(df_normalizado)
            
            # Fechas para Venezuela
            fechas_convertidas = pd.to_datetime(
                df_normalizado["FECHA"],
                dayfirst=True,
                errors="coerce"
            )
            
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
                fechas_convertidas = pd.to_datetime(
                    df_original[3].astype(str).str.zfill(8),
                    format="%d%m%Y",
                    errors="coerce"
                )
            elif banco == "venezuela":
                # Ya tenemos fechas_convertidas definidas arriba
                pass
            else:
                fechas_convertidas = pd.to_datetime(
                    df_original.iloc[:, 3],
                    errors="coerce",
                    dayfirst=True
                )

            fecha_inicio_dt = pd.to_datetime(fecha_inicio)
            fecha_fin_dt = pd.to_datetime(fecha_fin)

            # Aplicar filtro según el banco
            if banco == "venezuela":
                # Usar el dataframe original para el filtro
                mask = (fechas_convertidas >= fecha_inicio_dt) & (fechas_convertidas <= fecha_fin_dt)
                # Filtrar df_original usando la máscara
                df_original = df_original[mask]
            else:
                df_original = df_original[
                    (fechas_convertidas >= fecha_inicio_dt) & 
                    (fechas_convertidas <= fecha_fin_dt)
                ]
            
            st.success(f"Filtro de fechas aplicado: {fecha_inicio} a {fecha_fin}")
        except Exception as e:
            st.warning(f"Error filtrando fechas: {e}")
            
        if df_original.empty or len(df_original) == 0:
            st.error("❌ No se encontraron movimientos válidos en el rango de fechas.")
            st.stop()

        with st.expander("👁️ Vista previa archivo original"):
            st.dataframe(df_original.head(20), use_container_width=True)

        # =========================================================
        # LEER ARCHIVO IPAGO
        # =========================================================
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
                            "FECHA": row["FECHA"],
                            "REFERENCIA": row["REFERENCIA"],
                            "DESCRIPCIÓN": row["DESCRIPCION"],
                            "MONTO BS": monto_bs,
                            "TASA BCV": tasa,
                            "MONTO USD": monto_usd,
                            "STATUS": "",
                            "OBSERVACIÓN": "",
                            "TIPO_PAGO": "",
                            "PROVEEDOR_IPAGO": "",
                            "DESCRIPCION_ORIGINAL": ""
                        }

                        tipo = str(row["TIPO"]).strip().upper()
                        descripcion = str(row["DESCRIPCION"]).strip().upper()
                        referencia = str(row["REFERENCIA"]).strip()
                        
                        # 🔥 DETECCIÓN MEJORADA DE COMISIONES PARA VENEZUELA
                        es_comision_venezuela = False
                        
                        # 🔥 REGLA 1: Detectar por descripción
                        palabras_comision_bdv = [
                            "COM PAGO OTRAS CTAS",
                            "COMISION PAGO A PROVEEDORES",
                            "COM PAGO OTR BCOS",
                            "COM PAGO OTRAS CTAS JUR NAT",
                            "COM PAGO OTRAS CTAS JUR JUR",
                            "COMISION POR TRANSFERENCIA",
                            "COMISION PAGO MOVIL",
                            "COMISIÓN PAGO MOVIL",
                            "COMISION X PAGO DE NOMINA",
                            "COMISION X PAGO DE NOMINAS",
                            "ITF",
                            "IMPUESTO A LAS TRANSACCIONES FINANCIERAS",
                            "CARGO BANCARIO",
                            "MANTENIMIENTO DE CUENTA",
                            "COMISION BANCARIA",
                            "COMISIÓN BANCARIA",
                            "CARGO POR SERVICIO",
                            "CARGO POR TRANSACCION",
                            "COMISION PAGO MOVIL COMERCIAL",
                            "COMISION X PAGO DE NOMINAS MB"
                        ]
                        
                        # Verificar si la descripción coincide con alguna comisión
                        for patron in palabras_comision_bdv:
                            if patron in descripcion:
                                es_comision_venezuela = True
                                break
                        
                        # 🔥 REGLA 2: Detectar por referencia (comisiones de BDV tienen referencias específicas)
                        if not es_comision_venezuela:
                            # Las comisiones de BDV suelen tener referencias que comienzan con 970 o 972 o 067
                            if referencia.startswith(("970", "972", "067")):
                                # Verificar si la descripción contiene palabras clave de comisión
                                if any(palabra in descripcion for palabra in ["COM", "PAGO OTRAS", "PAGO OTR", "COMISION"]):
                                    es_comision_venezuela = True
                        
                        # 🔥 REGLA 3: Si el tipo es ND y la descripción contiene "COM" es una comisión
                        if not es_comision_venezuela and tipo == "ND":
                            if "COM" in descripcion or "PAGO OTR" in descripcion:
                                es_comision_venezuela = True
                        
                        # 🔥 REGLA 4: Comisiones específicas de BDV por el monto exacto (189.50, 36.44, etc)
                        if not es_comision_venezuela and tipo == "ND":
                            # Montos típicos de comisiones de BDV (montos pequeños)
                            if monto_bs < 1000 and ("COM" in descripcion or "PAGO OTR" in descripcion):
                                es_comision_venezuela = True
                        
                        # Si es comisión, agregar a la lista de comisiones
                        if es_comision_venezuela:
                            comisiones.append(registro)
                        elif tipo in ["NC", "C", "CREDITO", "ABONO"]:
                            ingresos.append(registro)
                        else:
                            egresos.append(registro)
                else:
                    # 🔥 PASAR EL NOMBRE DEL BANCO A LA FUNCIÓN
                    ingresos, egresos, comisiones = procesar_archivo(df_original, usar_api, banco=banco)

            df_ingresos = pd.DataFrame(ingresos)
            df_egresos = pd.DataFrame(egresos)
            df_comisiones = pd.DataFrame(comisiones)

            # =========================================================
            # 🔥 CRUCE CON IPAGO - VERSIÓN MEJORADA (CRUCE FLEXIBLE)
            # =========================================================
            if archivo_ipago and not df_egresos.empty:
                # Enriquecer egresos con datos de iPago
                df_egresos = enriquecer_egresos_con_ipago(df_egresos, df_ipago)
                
                # 🔥 Separar comisiones de iPago (si las hay)
                if "ES_COMISION" in df_egresos.columns:
                    mascara_comisiones_ipago = df_egresos["ES_COMISION"] == True
                    
                    if mascara_comisiones_ipago.any():
                        df_comisiones_extra = df_egresos[mascara_comisiones_ipago].copy()
                        
                        # Remover columnas internas
                        df_comisiones_extra = df_comisiones_extra.drop(
                            columns=["ES_COMISION", "REFERENCIA_IPAGO"], 
                            errors="ignore"
                        )
                        
                        # Agregar a comisiones existentes
                        if not df_comisiones.empty:
                            df_comisiones = pd.concat([df_comisiones, df_comisiones_extra], ignore_index=True)
                        else:
                            df_comisiones = df_comisiones_extra
                        
                        # Remover comisiones de egresos
                        df_egresos = df_egresos[~mascara_comisiones_ipago].copy()
                        
                        st.success(f"💳 Se movieron {len(df_comisiones_extra)} comisiones desde iPago a la sección de COMISIONES")
                
                # Limpiar columnas auxiliares
                df_egresos = df_egresos.drop(
                    columns=["ES_COMISION", "REFERENCIA_IPAGO"], 
                    errors="ignore"
                )
                
                st.success(f"🎯 Egresos enriquecidos con iPago: {len(df_egresos)} registros")

            # Completar columnas vacías obligatorias para el reporte en openpyxl
            for df_t in [df_ingresos, df_egresos, df_comisiones]:
                if not df_t.empty:
                    if "STATUS" not in df_t.columns: df_t["STATUS"] = ""
                    if "OBSERVACIÓN" not in df_t.columns: df_t["OBSERVACIÓN"] = ""
                    if "TIPO_PAGO" not in df_t.columns: df_t["TIPO_PAGO"] = ""
                    if "PROVEEDOR_IPAGO" not in df_t.columns: df_t["PROVEEDOR_IPAGO"] = ""
                    if "DESCRIPCION_ORIGINAL" not in df_t.columns: df_t["DESCRIPCION_ORIGINAL"] = ""

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

            with tab1: st.dataframe(df_ingresos, use_container_width=True)
            with tab2: st.dataframe(df_egresos, use_container_width=True)
            with tab3: st.dataframe(df_comisiones, use_container_width=True)

            # =========================================================
            # MOTOR DE REPORTES EXCEL OPENPYXL COMPLETO
            # =========================================================
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

                    # 🔥 HEADERS MEJORADOS CON DATOS DE IPAGO
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
        st.error("Detalles del error para depuración:")
        st.code(str(e))

else:
    st.markdown("""
    ### 👋 Clasificador Bancario Inteligente Multi-Banco

    ## FUNCIONES
    ✅ **Bancos soportados:** Mercantil, Banco de Venezuela, Banesco, Provincial, BNC, Tesoro, Bancamiga.
    ✅ Clasifica automáticamente: Ingresos (NC, C, CREDITO, ABONO), Egresos (ND, D, DEBITO, DEBIT), Comisiones.
    ✅ **NUEVO:** Cruce inteligente con iPago usando REFERENCIA + DESCRIPCIÓN.
    ✅ **NUEVO:** Exporta con datos completos de iPago: Proveedor, Tipo de Egreso, Tipo de Pago.
    ✅ **NUEVO:** Conserva la descripción original del banco y la reemplaza con la de iPago cuando hay coincidencia.
    ✅ Calcula USD con tasa BCV real por fecha.
    ✅ Exporta reporte profesional con todas las columnas.
    """)
