import streamlit as st
import pandas as pd
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
# 🔥 DETECTAR COMISIONES - VERSIÓN DEFINITIVA CON REGLAS CLARAS
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
        "comisión por"
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
# DETECTOR DE BANCO CORREGIDO
# =========================================================

def detectar_banco(nombre_archivo):
    nombre = nombre_archivo.upper()

    if "TESORO" in nombre or "TESORERIA" in nombre or "TES" in nombre:
        return "tesoro"
    elif "BANESCO" in nombre or re.match(r"^J\d+", nombre_archivo):
        return "banesco"
    elif "BANCAMIGA" in nombre or "MOVIMIENTOS_" in nombre:
        return "bancamiga"
    elif "MOVIMIENTOS EN MONEDA NACIONAL" in nombre or "VENEZUELA" in nombre or "BANCO DE VENEZUELA" in nombre or "BDV" in nombre:
        return "venezuela"
    elif "PROVINCIAL" in nombre:
        return "provincial"
    elif "BNC" in nombre:
        return "bnc"
    elif "MERCANTIL" in nombre:
        return "mercantil"
    return "mercantil"

# =========================================================
# 🔥 PROCESAR VENEZUELA - VERSIÓN MEJORADA (USA TIPO DE MOVIMIENTO DIRECTAMENTE)
# =========================================================

def procesar_venezuela(df):
    """Procesa el archivo del BDV usando directamente la columna Tipo de Movimiento (NC/ND)"""
    st.info("Procesando Banco de Venezuela (Modo Protegido)...")
    
    try:
        # Barrido radical de espacios duros de Excel (\xa0) en nombres de columnas
        df.columns = [
            unicodedata.normalize("NFKD", str(c))
            .encode("ascii", "ignore")
            .decode("utf-8")
            .strip()
            .replace("\xa0", "")
            for c in df.columns
        ]
        
        # Mapeo posicional dinámico e independiente
        col_fecha = next((c for c in df.columns if "fecha" in c.lower()), None)
        col_ref = next((c for c in df.columns if "referencia" in c.lower() or "ref" in c.lower()), None)
        col_desc = next((c for c in df.columns if "descrip" in c.lower() or "concepto" in c.lower()), None)
        col_tipo = next((c for c in df.columns if "tipo" in c.lower() and "mov" in c.lower()), None)
        col_credito = next((c for c in df.columns if "credito" in c.lower() or "haber" in c.lower()), None)
        col_debito = next((c for c in df.columns if "debito" in c.lower() or "debe" in c.lower()), None)
        
        movimientos = []
        
        for idx, fila in df.iterrows():
            try:
                if not col_fecha or pd.isna(fila[col_fecha]):
                    continue
                
                # Saneamiento de cadenas de fecha (corte de horas y unificación de guiones)
                fecha_raw = str(fila[col_fecha]).strip().replace("\xa0", "")
                if " " in fecha_raw:
                    fecha_raw = fecha_raw.split(" ")[0]
                fecha_raw = fecha_raw.replace("-", "/")
                
                fecha_val = pd.to_datetime(fecha_raw, dayfirst=True, errors="coerce")
                if pd.isna(fecha_val):
                    continue
                
                # Extracción libre de cadenas ocultas \xa0 en descriptivos
                referencia = str(fila[col_ref]).strip().replace("\xa0", "") if col_ref and pd.notna(fila[col_ref]) else ""
                descripcion = str(fila[col_desc]).strip().replace("\xa0", "") if col_desc and pd.notna(fila[col_desc]) else ""
                
                # 🔥 Obtener tipo de movimiento directamente de la columna
                tipo_mov = ""
                if col_tipo and pd.notna(fila[col_tipo]):
                    tipo_mov = str(fila[col_tipo]).strip().upper()
                
                # Omitir metadatos de balances impresos en el cuerpo
                if any(p in descripcion.upper() for p in ["SALDO INICIAL", "SALDO FINAL", "TOTALES"]):
                    continue
                
                # Extracción limpia de montos eliminando basura tipográfica
                val_credito = 0
                val_debito = 0
                
                if col_credito and pd.notna(fila[col_credito]):
                    clean_cred = str(fila[col_credito]).replace("\xa0", "").strip()
                    val_credito = convertir_monto(clean_cred) or 0
                    
                if col_debito and pd.notna(fila[col_debito]):
                    clean_deb = str(fila[col_debito]).replace("\xa0", "").strip()
                    val_debito = convertir_monto(clean_deb) or 0
                
                monto = 0
                tipo = ""
                
                # 🔥 USAR DIRECTAMENTE EL TIPO DE MOVIMIENTO DE LA COLUMNA
                if tipo_mov == "NC":
                    monto = abs(val_credito)
                    tipo = "NC"
                elif tipo_mov == "ND":
                    monto = abs(val_debito)
                    tipo = "ND"
                else:
                    # Fallback: si no hay tipo, inferir por el monto
                    if abs(val_credito) > 0:
                        monto = abs(val_credito)
                        tipo = "NC"
                    elif abs(val_debito) > 0:
                        monto = abs(val_debito)
                        tipo = "ND"
                    else:
                        continue
                
                # 🔥 IMPORTANTE: NO FILTRAR POR TIPO AQUÍ
                # Permitir que TODOS los movimientos pasen (incluyendo ND de comisiones)
                if monto <= 0:
                    continue
                
                movimientos.append({
                    "FECHA": fecha_val.strftime("%d/%m/%Y"),
                    "REFERENCIA": referencia,
                    "DESCRIPCION": descripcion,
                    "TIPO": tipo,
                    "MONTO": monto
                })
                
            except Exception as e:
                continue
                
        df_resultado = pd.DataFrame(movimientos)
        
        if df_resultado.empty:
            st.error("❌ No se encontraron movimientos válidos en el archivo de Venezuela.")
            return pd.DataFrame()
            
        st.success(f"✅ Venezuela OK: {len(df_resultado)} movimientos detectados")
        st.dataframe(df_resultado.head(10))
        return df_resultado
        
    except Exception as e:
        st.error(f"Error procesando Venezuela: {str(e)}")
        return pd.DataFrame()

# =========================================================
# 🔥 FUNCIÓN MEJORADA: ENRIQUECER EGRESOS CON IPAGO (USA REFERENCIA + DESCRIPCIÓN)
# =========================================================

def enriquecer_egresos_con_ipago(df_egresos, df_ipago):
    """
    Enriquece los egresos del banco con datos de iPago.
    Usa REFERENCIA + DESCRIPCIÓN para hacer el cruce.
    Mantiene TODOS los egresos y agrega información cuando hay coincidencia.
    
    Args:
        df_egresos: DataFrame con los egresos del banco
        df_ipago: DataFrame con los movimientos de iPago
    
    Returns:
        DataFrame: egresos enriquecidos con datos de iPago
    """
    if df_ipago is None or df_ipago.empty:
        return df_egresos
    
    # Hacer una copia para no modificar el original
    df_resultado = df_egresos.copy()
    
    # 🔥 NORMALIZAR REFERENCIAS (eliminar puntos y espacios)
    df_resultado["REFERENCIA_NORM"] = (
        df_resultado["REFERENCIA"]
        .astype(str)
        .str.replace(".0", "", regex=False)
        .str.replace(" ", "", regex=False)
        .str.strip()
    )
    
    df_ipago["Referencia_Norm"] = (
        df_ipago["Referencia"]
        .astype(str)
        .str.replace(".0", "", regex=False)
        .str.replace(" ", "", regex=False)
        .str.strip()
    )
    
    # 🔥 NORMALIZAR DESCRIPCIONES (eliminar acentos, mayúsculas, espacios)
    df_resultado["DESCRIPCION_NORM"] = (
        df_resultado["DESCRIPCIÓN"]
        .astype(str)
        .apply(normalizar_texto)
        .str.replace(" ", "")
        .str.strip()
    )
    
    df_ipago["Descripción_Norm"] = (
        df_ipago["Descripción"]
        .astype(str)
        .apply(normalizar_texto)
        .str.replace(" ", "")
        .str.strip()
    )
    
    # 🔥 CREAR CLAVE DE CRUCE: REFERENCIA + PRIMERAS PALABRAS DE DESCRIPCIÓN
    df_resultado["CLAVE_CRUCE"] = (
        df_resultado["REFERENCIA_NORM"] + "|" + 
        df_resultado["DESCRIPCION_NORM"].str[:20]  # Primeros 20 caracteres de la descripción normalizada
    )
    
    df_ipago["CLAVE_CRUCE"] = (
        df_ipago["Referencia_Norm"] + "|" + 
        df_ipago["Descripción_Norm"].str[:20]  # Primeros 20 caracteres de la descripción normalizada
    )
    
    # 🔥 CREAR DICCIONARIO DE IPAGO CON TODOS LOS DATOS
    ipago_dict = {}
    for _, row in df_ipago.iterrows():
        clave = row.get("CLAVE_CRUCE", "")
        if clave and clave not in ipago_dict:
            ipago_dict[clave] = {
                "PROVEEDOR": row.get("Proveedor", ""),
                "TIPO_EGRESO": row.get("Tipo de Egreso", ""),
                "TIPO_PAGO": row.get("Tipo de Pago", ""),
                "DESCRIPCION_IPAGO": row.get("Descripción", ""),
                "FECHA_PAGO": row.get("Fecha Pago", ""),
                "EMPRESA": row.get("Empresa", ""),
                "MONTO": row.get("Monto", 0),
                "MONTO_USD": row.get("Monto USD", 0)
            }
    
    # 🔥 ENRIQUECER CADA EGRESO
    for idx, row in df_resultado.iterrows():
        clave = row.get("CLAVE_CRUCE", "")
        
        if clave in ipago_dict:
            # Si hay coincidencia, enriquecer con datos de iPago
            datos = ipago_dict[clave]
            df_resultado.at[idx, "STATUS"] = datos["PROVEEDOR"]
            df_resultado.at[idx, "OBSERVACIÓN"] = datos["TIPO_EGRESO"]
            df_resultado.at[idx, "TIPO_PAGO"] = datos["TIPO_PAGO"]
            df_resultado.at[idx, "PROVEEDOR_IPAGO"] = datos["PROVEEDOR"]
            
            # 🔥 Reemplazar descripción con la de iPago
            descripcion_ipago = datos["DESCRIPCION_IPAGO"]
            if descripcion_ipago:
                df_resultado.at[idx, "DESCRIPCIÓN"] = descripcion_ipago
                df_resultado.at[idx, "DESCRIPCION_ORIGINAL"] = row.get("DESCRIPCIÓN", "")
            
            # 🔥 Si es comisión, marcarlo como tal
            tipo = str(datos["TIPO_EGRESO"]).upper()
            desc = str(datos["DESCRIPCION_IPAGO"]).upper()
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
# PROCESAR PROVINCIAL
# =========================================================

def procesar_provincial(df):
    st.info("Procesando archivo de Provincial...")
    encabezado = None
    for i in range(min(20, len(df))):
        fila = df.iloc[i].astype(str)
        if fila.str.contains("fecha", case=False).any():
            encabezado = i
            break

    if encabezado is not None:
        df.columns = df.iloc[encabezado]
        df = df.iloc[encabezado+1:].reset_index(drop=True)

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
        df["FECHA"] = pd.to_datetime(df["FECHA"], dayfirst=True, errors="coerce")
        df = df[df["FECHA"].notna()]

    if "MONTO" in df.columns:
        monto = df["MONTO"]
        if isinstance(monto, pd.DataFrame):
            monto = monto.iloc[:, 0]
        df["MONTO"] = pd.to_numeric(monto, errors="coerce")
        df = df[df["MONTO"].notna()]

    if "TIPO" not in df.columns:
        df["TIPO"] = "ND"

    return df

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
# PROCESAR BANCAMIGA
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
    }
    fecha_str = fecha_obj.strftime("%d/%m/%Y")
    return tasas_bcv_local.get(fecha_str, None)

def obtener_tasa_por_fecha(fecha_obj, usar_api=False):
    return obtener_tasa_bcv_fecha(fecha_obj)

# =========================================================
# CONVERTIR A FORMATO MERCANTIL
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
                "",           # col9
            ]
            datos_convertidos.append(fila_convertida)
            
        except Exception as e:
            continue
    
    df_convertido = pd.DataFrame(datos_convertidos)
    return df_convertido if len(df_convertido) > 0 else pd.DataFrame()

# =========================================================
# 🔥 PROCESAMIENTO MERCANTIL - CON REGLAS ESPECÍFICAS PARA COMISIONES
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
                df_original = convertir_a_formato_mercantil(df_normalizado, banco)
            except Exception as e:
                st.error(f"Error leyendo Provincial: {str(e)}")
                st.stop()
            
        elif banco == "venezuela":
            df_raw = leer_excel_con_encabezados(archivo)
            df_normalizado = procesar_venezuela(df_raw)
            if df_normalizado.empty:
                st.stop()
            df_original = convertir_a_formato_mercantil(df_normalizado, banco)
            
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
            else:
                fechas_convertidas = pd.to_datetime(
                    df_original.iloc[:, 3],
                    errors="coerce",
                    dayfirst=True
                )

            fecha_inicio_dt = pd.to_datetime(fecha_inicio)
            fecha_fin_dt = pd.to_datetime(fecha_fin)

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
                        descripcion = str(row["DESCRIPCION"]).strip()

                        if es_comision(descripcion, None):
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
            # 🔥 CRUCE CON IPAGO - VERSIÓN MEJORADA (REFERENCIA + DESCRIPCIÓN)
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
                            columns=["REFERENCIA_NORM", "DESCRIPCION_NORM", "CLAVE_CRUCE", "ES_COMISION"], 
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
                    columns=["REFERENCIA_NORM", "DESCRIPCION_NORM", "CLAVE_CRUCE", "ES_COMISION"], 
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
