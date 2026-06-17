import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import date
import requests
from bs4 import BeautifulSoup
import json
import re
from datetime import datetime
import unicodedata  # 🔥 Para normalización estricta de textos y eliminar acentos

from openpyxl.styles import Font
from openpyxl.styles import PatternFill
from openpyxl.styles import Border
from openpyxl.styles import Side
from openpyxl.styles import Alignment
from openpyxl.drawing.image import Image

# =========================================================
# CONFIGURACIÓN GENERAL DE STREAMLIT
# =========================================================

st.set_page_config(
    page_title="Clasificador Bancario - Grupo Bodeguita Oriente",
    page_icon="🏦",
    layout="wide"
)

# =========================================================
# ESTILOS VISUALES CORPORATIVOS
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
# FUNCIÓN DE LIMPIEZA DE CARACTERES (ESCUDO ANTIBUGS)
# =========================================================

def normalizar_texto(texto):
    """Normaliza texto eliminando acentos, espacios duros y convirtiendo a minúsculas"""
    texto = str(texto)
    texto = (
        unicodedata.normalize("NFKD", texto)
        .encode("ascii", "ignore")
        .decode("utf-8")
    )
    return texto.lower()

# =========================================================
# COMPONENTE DE ENCABEZADO VISUAL
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
# PANEL LATERAL DE CONTROLES (SIDEBAR)
# =========================================================

with st.sidebar:
    st.image(
        "https://raw.githubusercontent.com/pelobravo/clasificador-excel/main/LOGO.jpeg",
        width=100
    )
    st.markdown("---")
    archivo = st.file_uploader(
        "📂 Cargar archivo Excel Bancario",
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
# ENGINES AUTOMÁTICOS DE RECONOCIMIENTO EXCEL
# =========================================================

def leer_excel_sin_encabezados(archivo):
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
# FORMATEADORES DE MONEDAS Y NÚMEROS
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
        valor = valor_original.replace(" ", "").replace("$", "").replace("Bs", "").replace("€", "")

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

# =========================================================
# REGLAS AUTOMÁTICAS DE COMISIONES BANCARIAS
# =========================================================

def es_comision(texto):
    texto = normalizar_texto(texto).strip()
    palabras = [
        "comision", "comisión", "cargo", "cargo bancario", "fee", "iva", "itf", "impuesto",
        "op.cred", "op cred", "credito directo", "transferencia de fondos",
        "comision por transferencia", "comision pago movil", "comisión pago movil",
        "servicio bancario", "gasto bancario", "mantenimiento de cuenta", "debito automatico bancario",
        "com ", "com.", "com pago", "com pago otr", "com pago otr bcos", "comision pago proveedores",
        "descuento tarjeta", "descuento de tarjeta", "comision tarjeta", "comisión tarjeta",
        "cargo tarjeta", "retencion tarjeta", "retención tarjeta", "comision punto de venta",
        "comisión punto de venta", "punto de venta", "comision pos", "comisión pos",
        "descuento pos", "cargo por servicio", "cargo por transaccion", "cargo por transacción",
    ]
    return any(p in texto for p in palabras)

# =========================================================
# MAQUINA DE RECONOCIMIENTO DE ENTIDADES BANCARIAS
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
# ESCUDO ULTRA-ROBUSTO: BANCO DE VENEZUELA (PROTEGIDO)
# =========================================================

def procesar_venezuela(df):
    """Procesa el archivo de BDV barriendo espacios invisibles, saltos de fecha y celdas vacías"""
    st.info("Procesando Banco de Venezuela (Modo Protegido)...")
    try:
        # Barrido de caracteres invisibles \xa0 y normalización estricta de cabeceras
        df.columns = [
            unicodedata.normalize("NFKD", str(c))
            .encode("ascii", "ignore")
            .decode("utf-8")
            .strip()
            .replace("\xa0", "")
            for c in df.columns
        ]
        
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
                
                # Gestión tolerante de fechas con horas inyectadas por el banco
                fecha_raw = str(fila[col_fecha]).strip().replace("\xa0", "")
                if " " in fecha_raw:
                    fecha_raw = fecha_raw.split(" ")[0]
                fecha_raw = fecha_raw.replace("-", "/")
                
                fecha_val = pd.to_datetime(fecha_raw, dayfirst=True, errors="coerce")
                if pd.isna(fecha_val):
                    continue
                
                referencia = str(fila[col_ref]).strip().replace("\xa0", "") if col_ref and pd.notna(fila[col_ref]) else ""
                descripcion = str(fila[col_desc]).strip().replace("\xa0", "") if col_desc and pd.notna(fila[col_desc]) else ""
                tipo = str(fila[col_tipo]).strip().upper() if col_tipo and pd.notna(fila[col_tipo]) else ""
                
                if any(p in descripcion.upper() for p in ["SALDO INICIAL", "SALDO FINAL", "TOTALES"]):
                    continue
                
                # Extracción limpia de montos eliminando basura invisible
                val_credito = 0
                val_debito = 0
                if col_credito and pd.notna(fila[col_credito]):
                    clean_cred = str(fila[col_credito]).replace("\xa0", "").strip()
                    val_credito = convertir_monto(clean_cred) or 0
                if col_debito and pd.notna(fila[col_debito]):
                    clean_deb = str(fila[col_debito]).replace("\xa0", "").strip()
                    val_debito = convertir_monto(clean_deb) or 0
                
                monto = 0
                if val_credito > 0:
                    monto = val_credito
                    tipo = "NC"
                elif val_debito > 0:
                    monto = val_debito
                    tipo = "ND"
                else:
                    continue
                
                movimientos.append({
                    "FECHA": fecha_val.strftime("%d/%m/%Y"),
                    "REFERENCIA": referencia,
                    "DESCRIPCION": descripcion,
                    "TIPO": tipo,
                    "MONTO": monto
                })
            except Exception:
                continue
                
        df_resultado = pd.DataFrame(movimientos)
        if df_resultado.empty:
            st.error("❌ No se encontraron movimientos válidos. Revise que sea el Excel original.")
            return pd.DataFrame()
            
        st.success(f"✅ BDV Protegido: {len(df_resultado)} registros clasificados sin obstrucción.")
        return df_resultado
    except Exception as e:
        st.error(f"Error crítico en escudo BDV: {str(e)}")
        return pd.DataFrame()

# =========================================================
# PARSERS INDEPENDIENTES DE ENTIDADES BANCARIAS SECUNDARIAS
# =========================================================

def procesar_banesco(df):
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
        df["FECHA"] = pd.to_datetime(df["FECHA"], dayfirst=True, errors="coerce")
        df = df[df["FECHA"].notna()]
        df["TIPO"] = df["MONTO_RAW"].astype(str).apply(lambda x: "NC" if "+" in x else "ND")
        df["MONTO"] = df["MONTO_RAW"].astype(str).str.replace("+", "", regex=False).str.replace("-", "", regex=False).str.replace(".", "", regex=False).str.replace(",", ".", regex=False).str.strip()
        df["MONTO"] = pd.to_numeric(df["MONTO"], errors="coerce")
        df = df[(df["MONTO"].notna()) & (df["MONTO"] > 0)]
        return df[["FECHA", "REFERENCIA", "DESCRIPCION", "TIPO", "MONTO"]]
    except: return pd.DataFrame()

def procesar_provincial(df):
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
        if "fecha" in col_str: rename_map[col] = "FECHA"
        elif "descrip" in col_str or "concepto" in col_str: rename_map[col] = "DESCRIPCION"
        elif "monto" in col_str: rename_map[col] = "MONTO"
    df = df.rename(columns=rename_map)
    if "FECHA" in df.columns:
        df["FECHA"] = pd.to_datetime(df["FECHA"], dayfirst=True, errors="coerce")
        df = df[df["FECHA"].notna()]
    if "MONTO" in df.columns:
        monto = df["MONTO"]
        if isinstance(monto, pd.DataFrame): monto = monto.iloc[:, 0]
        df["MONTO"] = pd.to_numeric(monto, errors="coerce")
        df = df[df["MONTO"].notna()]
    if "TIPO" not in df.columns: df["TIPO"] = "ND"
    return df

def procesar_bnc(df):
    encabezado = None
    for i in range(min(30, len(df))):
        fila = df.iloc[i].fillna("").astype(str)
        texto = " ".join(fila.tolist()).lower()
        if "fecha" in texto and ("descripcion" in texto or "descripción" in texto):
            encabezado = i
            break
    if encabezado is None: return pd.DataFrame()
    headers = [str(col).strip().replace("\n", " ") if str(col).strip() != "" else f"COLUMNA_{idx}" for idx, col in enumerate(df.iloc[encabezado])]
    headers_unicos = []
    contador = {}
    for h in headers:
        if h in contador:
            contador[h] += 1
            headers_unicos.append(f"{h}_{contador[h]}")
        else:
            contador[h] = 0
            headers_unicos.append(h)
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
    df["FECHA"] = pd.to_datetime(df["FECHA"], dayfirst=True, errors="coerce")
    df = df[df["FECHA"].notna()]
    df["CREDITO"] = pd.to_numeric(df["CREDITO"], errors="coerce").fillna(0)
    df["DEBITO"] = pd.to_numeric(df["DEBITO"], errors="coerce").fillna(0)
    df["MONTO"] = df["CREDITO"] - df["DEBITO"]
    df["TIPO"] = df["MONTO"].apply(lambda x: "NC" if x > 0 else "ND")
    df["MONTO"] = df["MONTO"].abs()
    return df[df["MONTO"] != 0]

def procesar_tesoro(df):
    try:
        encabezado = None
        for i in range(min(20, len(df))):
            fila = df.iloc[i].astype(str)
            texto = " ".join(map(str, fila.tolist())).lower()
            if "fecha" in texto and "referencia" in texto and "concepto" in texto:
                encabezado = i
                break
        if encabezado is None: return pd.DataFrame()
        df.columns = [str(c).strip() for c in df.iloc[encabezado]]
        df = df.iloc[encabezado + 1:].reset_index(drop=True)
        rename_map = {}
        for col in df.columns:
            c = str(col).strip().lower()
            if "fecha" in c: rename_map[col] = "FECHA"
            elif "referencia" in c: rename_map[col] = "REFERENCIA"
            elif "concepto" in c: rename_map[col] = "DESCRIPCION"
            elif "débito" in c or "debito" in c: rename_map[col] = "DEBITO"
            elif "crédito" in c or "credito" in c: rename_map[col] = "CREDITO"
        df = df.rename(columns=rename_map)
        df["FECHA"] = pd.to_datetime(df["FECHA"], dayfirst=True, errors="coerce")
        df = df[df["FECHA"].notna()]
        def limpiar_numero(v):
            v = str(v).replace(".", "").replace(",", ".")
            try: return float(v)
            except: return 0
        df["CREDITO"] = df["CREDITO"].apply(limpiar_numero)
        df["DEBITO"] = df["DEBITO"].apply(limpiar_numero)
        df["MONTO"] = df["CREDITO"] - df["DEBITO"]
        df["TIPO"] = df["MONTO"].apply(lambda x: "NC" if x > 0 else "ND")
        df["MONTO"] = df["MONTO"].abs()
        return df[df["MONTO"] > 0][["FECHA", "REFERENCIA", "DESCRIPCION", "TIPO", "MONTO"]]
    except: return pd.DataFrame()

def procesar_bancamiga(df):
    try:
        encabezado = None
        for i in range(min(15, len(df))):
            fila = df.iloc[i].fillna("").astype(str)
            texto = " ".join(fila.tolist()).lower()
            if "fecha" in texto and "referencia" in texto and "concepto" in texto:
                encabezado = i
                break
        if encabezado is not None:
            df.columns = [str(c).strip() for c in df.iloc[encabezado]]
            df = df.iloc[encabezado + 1:].reset_index(drop=True)
        if isinstance(df.columns, pd.MultiIndex):
            df.columns = [col[-1] for col in df.columns]
        df = df.rename(columns={"Fecha": "FECHA", "Referencia": "REFERENCIA", "Concepto": "DESCRIPCION", "Débito": "DEBITO", "Crédito": "CREDITO"})
        df["FECHA"] = pd.to_datetime(df["FECHA"], dayfirst=True, errors="coerce")
        df = df[df["FECHA"].notna()]
        df["CREDITO"] = df["CREDITO"].apply(convertir_monto).fillna(0)
        df["DEBITO"] = df["DEBITO"].apply(convertir_monto).fillna(0)
        df["MONTO"] = df["CREDITO"] - df["DEBITO"]
        df["TIPO"] = df["MONTO"].apply(lambda x: "NC" if x > 0 else "ND")
        df["MONTO"] = df["MONTO"].abs()
        return df[df["MONTO"] > 0][["FECHA", "REFERENCIA", "DESCRIPCION", "TIPO", "MONTO"]]
    except: return pd.DataFrame()

# =========================================================
# BANCO DE DATOS DE TASAS HISTÓRICAS LOCALES DEL BCV
# =========================================================

@st.cache_data(ttl=3600)
def obtener_tasa_bcv_fecha(fecha_obj):
    tasas_bcv_local = {
        "01/06/2026": 554.4258, "02/06/2026": 557.9741, "03/06/2026": 558.6436,
        "04/06/2026": 560.3753, "05/06/2026": 563.2892, "06/06/2026": 567.6828,
        "07/06/2026": 567.6828, "08/06/2026": 567.6828, "09/06/2026": 567.6828,
        "10/06/2026": 572.6784, "11/06/2026": 577.5461, "12/06/2026": 582.6862,
        "13/06/2026": 587.4059, "14/06/2026": 587.4059, "15/06/2026": 587.4059,
        "16/06/2026": 592.5163,
    }
    fecha_str = fecha_obj.strftime("%d/%m/%Y")
    return tasas_bcv_local.get(fecha_str, None)

def obtener_tasa_por_fecha(fecha_obj, usar_api=False):
    return obtener_tasa_bcv_fecha(fecha_obj)

# =========================================================
# CONTROLADOR BASE MERCANTIL ORIGINAL
# =========================================================

def procesar_archivo(df, usar_api=False):
    ingresos, egresos, comisiones = [], [], []
    registros_procesados = set()
    cache_tasas = {}

    for _, fila in df.iterrows():
        try:
            if len(fila) < 10: continue
            fecha_raw = str(fila[3]).strip().replace(".0", "")
            if fecha_raw.lower() == "nan": continue

            if len(fecha_raw) == 7: fecha = f"0{fecha_raw[0]}/{fecha_raw[1:3]}/{fecha_raw[3:]}"
            elif len(fecha_raw) == 8: fecha = f"{fecha_raw[0:2]}/{fecha_raw[2:4]}/{fecha_raw[3:]}"
            else: fecha = fecha_raw

            tipo = str(fila[5]).strip().upper()
            descripcion = str(fila[6]).strip()
            referencia = str(fila[4]).strip()
            monto_bs = convertir_monto(fila[7])

            if monto_bs is None or monto_bs == 0: continue
            fecha_obj = pd.to_datetime(fecha, dayfirst=True, errors="coerce")
            if pd.isna(fecha_obj): continue
            
            fecha_key = fecha_obj.strftime("%d/%m/%Y")
            tasa = cache_tasas.get(fecha_key, obtener_tasa_por_fecha(fecha_obj, usar_api))
            if tasa is None: tasa = 1.0
            else: cache_tasas[fecha_key] = tasa

            monto_usd = calcular_usd(monto_bs, tasa)
            if monto_usd is None: continue

            if descripcion.upper() in ["SALDO", "DESCRIPCION", "DESCRIPCIÓN", "REFERENCIA", "MOVIMIENTO", "FECHA", "SALDO INICIAL", "SALDO FINAL"]:
                continue

            registro = {"FECHA": fecha, "REFERENCIA": referencia, "DESCRIPCIÓN": descripcion, "MONTO BS": round(abs(monto_bs), 2), "TASA BCV": round(tasa, 4), "MONTO USD": monto_usd}
            clave = (fecha, referencia, descripcion, monto_usd, tipo)

            if clave in registros_procesados: continue
            registros_processed = registros_procesados.add(clave)

            if es_comision(descripcion): comisiones.append(registro)
            elif tipo in ["NC", "C", "CREDITO", "ABONO"]: ingresos.append(registro)
            else: egresos.append(registro)
        except:
            continue
    return ingresos, egresos, comisiones

def convertir_a_formato_mercantil(df, banco):
    datos_convertidos = []
    for idx, fila in df.iterrows():
        try:
            fecha = fila.get("FECHA", "")
            if pd.isna(fecha): continue
            fecha_str = fecha.strftime("%d/%m/%Y") if isinstance(fecha, (pd.Timestamp, datetime)) else str(fecha)
            
            fila_convertida = [
                "", "", "", fecha_str,
                fila.get("REFERENCIA", ""),
                fila.get("TIPO", ""),
                fila.get("DESCRIPCION", ""),
                fila.get("MONTO", 0),
                "", ""
            ]
            datos_convertidos.append(fila_convertida)
        except:
            continue
    return pd.DataFrame(datos_convertidos) if datos_convertidos else pd.DataFrame()

# =========================================================
# ORQUESTADOR CENTRAL DE EJECUCIÓN (STREAMLIT LOGIC)
# =========================================================

df_ipago = None

if archivo:
    st.info(f"📄 Archivo detectado en buffer: **{archivo.name}**")
    try:
        archivo.seek(0)
        primeros_bytes = archivo.read(100)
        archivo.seek(0)
        
        banco = "banesco" if b"<table" in primeros_bytes.lower() else detectar_banco(archivo.name)
        
        if banco == "mercantil":
            df_original = leer_excel_sin_encabezados(archivo)
        elif banco == "banesco":
            nombre = archivo.name.lower()
            df_raw = pd.read_excel(archivo, engine="openpyxl", header=None) if (nombre.endswith(".xlsx") or nombre.endswith(".xlsm")) else pd.read_html(archivo)[0]
            df_normalizado = procesar_banesco(df_raw)
            df_original = convertir_a_formato_mercantil(df_normalizado, banco)
        elif banco == "tesoro":
            df_raw = pd.read_excel(archivo, engine="openpyxl")
            df_normalizado = procesar_tesoro(df_raw)
            df_original = convertir_a_formato_mercantil(df_normalizado, banco)
        elif banco == "bancamiga":
            try: df_raw = pd.read_html(archivo)[0]
            except: 
                archivo.seek(0)
                df_raw = pd.read_excel(archivo, engine="xlrd" if archivo.name.lower().endswith(".xls") else "openpyxl", header=5 if archivo.name.lower().endswith(".xls") else 0)
            df_normalizado = procesar_bancamiga(df_raw)
            df_original = convertir_a_formato_mercantil(df_normalizado, banco)
        elif banco == "provincial":
            df_raw = leer_excel_sin_encabezados(archivo)
            df_normalizado = procesar_provincial(df_raw)
            df_original = convertir_a_formato_mercantil(df_normalizado, banco)
        elif banco == "venezuela":
            df_raw = leer_excel_con_encabezados(archivo)
            df_normalizado = procesar_venezuela(df_raw)
            if df_normalizado.empty: st.stop()
            df_original = convertir_a_formato_mercantil(df_normalizado, banco)
        elif banco == "bnc":
            df_raw = leer_excel_con_encabezados(archivo)
            df_normalizado = procesar_bnc(df_raw)
            df_original = convertir_a_formato_mercantil(df_normalizado, banco)
        else:
            df_raw = leer_excel_con_encabezados(archivo)
            df_original = convertir_a_formato_mercantil(df_raw, banco)

        # Filtros de Rangos Cronológicos
        if df_original.empty: st.stop()
        try:
            fechas_convertidas = pd.to_datetime(df_original[3].astype(str).str.zfill(8), format="%d%m%Y", errors="coerce") if banco == "mercantil" else pd.to_datetime(df_original.iloc[:, 3], errors="coerce", dayfirst=True)
            df_original = df_original[(fechas_convertidas >= pd.to_datetime(fecha_inicio)) & (fechas_convertidas <= pd.to_datetime(fecha_fin))]
        except Exception as e:
            st.warning(f"Filtro de fechas: {e}")

        if df_original.empty:
            st.error("❌ No existen transacciones registradas para este rango de fechas.")
            st.stop()

        with st.expander("👁️ Ver matriz de datos original"):
            st.dataframe(df_original.head(20), use_container_width=True)

        if archivo_ipago:
            try:
                df_ipago = pd.read_excel(archivo_ipago, engine="openpyxl")
                df_ipago["Referencia"] = pd.to_numeric(df_ipago["Referencia"], errors="coerce").fillna(0).astype("Int64").astype(str)
                df_ipago["REF_CRUCE"] = df_ipago["Referencia"].str[-6:]
            except Exception as e:
                st.error(f"Error procesando iPago: {e}")

        if procesar:
            with st.spinner("Clasificando movimientos y aplicando tasas comerciales..."):
                if banco == "venezuela":
                    ingresos, egresos, comisiones = [], [], []
                    for _, row in df_normalizado.iterrows():
                        fecha_obj = pd.to_datetime(row["FECHA"], dayfirst=True, errors="coerce")
                        tasa = obtener_tasa_por_fecha(fecha_obj, usar_api) or 1.0
                        monto_bs = float(row["MONTO"])
                        monto_usd = calcular_usd(monto_bs, tasa)

                        registro = {"FECHA": row["FECHA"], "REFERENCIA": row["REFERENCIA"], "DESCRIPCIÓN": row["DESCRIPCION"], "MONTO BS": monto_bs, "TASA BCV": tasa, "MONTO USD": monto_usd}
                        tipo = str(row["TIPO"]).strip().upper()
                        descripcion = str(row["DESCRIPCION"]).strip()

                        # SISTEMA UNIFICADO DE CLASIFICACIÓN (SANEADO)
                        if es_comision(descripcion): 
                            comisiones.append(registro)
                        elif tipo in ["NC", "C", "CREDITO", "ABONO"]: 
                            ingresos.append(registro)
                        else: 
                            egresos.append(registro)
                else:
                    ingresos, egresos, comisiones = procesar_archivo(df_original, usar_api)

            df_ingresos = pd.DataFrame(ingresos)
            df_egresos = pd.DataFrame(egresos)
            df_comisiones = pd.DataFrame(comisiones)

            # Sincronización Estricta de Cuentas de iPago
            if df_ipago is not None and not df_egresos.empty:
                df_egresos["REFERENCIA"] = df_egresos["REFERENCIA"].astype(str).str.replace(".0", "", regex=False).str.replace(" ", "", regex=False).str.strip()
                df_egresos["REF_CRUCE"] = df_egresos["REFERENCIA"].str[-6:]
                df_egresos = df_egresos.merge(df_ipago, on="REF_CRUCE", how="left")
                if "Descripción" in df_egresos.columns: df_egresos["DESCRIPCIÓN"] = df_egresos["Descripción"].fillna(df_egresos["DESCRIPCIÓN"])
                if "Proveedor" in df_egresos.columns: df_egresos["BENEFICIARIO"] = df_egresos["Proveedor"].fillna(df_egresos.get("BENEFICIARIO", ""))
                if "Tipo de Egreso" in df_egresos.columns: df_egresos["CONCEPTO"] = df_egresos["Tipo de Egreso"]

            # Tarjetas de Indicadores KPI en Pantalla
            total_ingresos = df_ingresos["MONTO USD"].sum() if not df_ingresos.empty else 0
            total_egresos = df_egresos["MONTO USD"].sum() if not df_egresos.empty else 0
            total_comisiones = df_comisiones["MONTO USD"].sum() if not df_comisiones.empty else 0

            col1, col2, col3 = st.columns(3)
            col1.metric("💰 TOTAL INGRESOS", len(df_ingresos), f"${total_ingresos:,.2f}")
            col2.metric("💸 TOTAL EGRESOS", len(df_egresos), f"${total_egresos:,.2f}")
            col3.metric("💳 TOTAL COMISIONES", len(df_comisiones), f"${total_comisiones:,.2f}")

            tab1, tab2, tab3 = st.tabs(["📈 INGRESOS", "📉 EGRESOS", "💳 COMISIONES"])
            tab1.dataframe(df_ingresos, use_container_width=True)
            tab2.dataframe(df_egresos, use_container_width=True)
            tab3.dataframe(df_comisiones, use_container_width=True)

            # =========================================================
            # ENGINE DE EXPORTACIÓN A EXCEL - OPENPYXL RESTAURADO
            # =========================================================
            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                workbook = writer.book
                hoja = workbook.create_sheet(title="REPORTE")
                if "Sheet" in workbook.sheetnames: workbook.remove(workbook["Sheet"])

                rojo = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
                verde = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")
                amarillo = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
                blanco = Font(color="FFFFFF", bold=True)
                borde = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
                centro = Alignment(horizontal="center", vertical="center")

                try:
                    logo = Image("LOGO.jpeg")
                    logo.width, logo.height = 130, 130
                    hoja.add_image(logo, "A1")
                except: pass

                hoja.merge_cells("C7:H7")
                hoja["C7"] = f"{banco.upper()} - INFORME CON AJUSTE BCV"
                hoja["C7"].font = Font(bold=True, size=14)
                hoja["C7"].alignment = centro

                def crear_tabla(titulo, dataframe, fila_inicio, color_total):
                    hoja.merge_cells(start_row=fila_inicio, start_column=1, end_row=fila_inicio, end_column=8)
                    t_cell = hoja.cell(row=fila_inicio, column=1, value=titulo)
                    t_cell.fill, t_cell.font, t_cell.alignment = rojo, blanco, centro

                    headers = ["FECHA", "REFERENCIA", "DESCRIPCIÓN", "MONTO BS", "TASA BCV", "MONTO USD", "STATUS", "OBSERVACIÓN"]
                    for col_num, header in enumerate(headers, 1):
                        c = hoja.cell(row=fila_inicio+1, column=col_num, value=header)
                        c.fill, c.font, c.border, c.alignment = rojo, blanco, borde, centro

                    fila_data = fila_inicio + 2
                    for _, row in dataframe.iterrows():
                        hoja.cell(row=fila_data, column=1, value=row["FECHA"])
                        hoja.cell(row=fila_data, column=2, value=row["REFERENCIA"])
                        hoja.cell(row=fila_data, column=3, value=row["DESCRIPCIÓN"])
                        hoja.cell(row=fila_data, column=4, value=row["MONTO BS"]).number_format = '#,##0.00'
                        hoja.cell(row=fila_data, column=5, value=row["TASA BCV"]).number_format = '#,##0.0000'
                        hoja.cell(row=fila_data, column=6, value=row["MONTO USD"]).number_format = '$#,##0.00'
                        for col in range(1, 9): hoja.cell(row=fila_data, column=col).border = borde
                        fila_data += 1

                    hoja.cell(row=fila_data, column=3, value=f"TOTAL {titulo}").font = Font(bold=True)
                    tot_bs = hoja.cell(row=fila_data, column=4, value=dataframe["MONTO BS"].sum())
                    tot_bs.number_format, tot_bs.fill = '#,##0.00', color_total
                    tot_usd = hoja.cell(row=fila_data, column=6, value=dataframe["MONTO USD"].sum())
                    tot_usd.number_format, tot_usd.fill = '$#,##0.00', color_total
                    return fila_data + 4

                fila_actual = 10
                if not df_ingresos.empty: fila_actual = crear_tabla("INGRESOS", df_ingresos, fila_actual, verde)
                if not df_egresos.empty: fila_actual = crear_tabla("EGRESOS", df_egresos, fila_actual, amarillo)
                if not df_comisiones.empty: fila_actual = crear_tabla("COMISIONES", df_comisiones, fila_actual, amarillo)

                for columna in hoja.columns:
                    try:
                        columna_letra = columna[0].column_letter
                        max_length = max(len(str(cell.value or '')) for cell in columna)
                        hoja.column_dimensions[columna_letra].width = min(max_length + 5, 50)
                    except: continue

            output.seek(0)
            st.download_button(
                label="📥 Descargar Excel Clasificado (con Tasas BCV)",
                data=output.getvalue(),
                file_name=f"balance_{banco}_{fecha_inicio}_{fecha_fin}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
    except Exception as e:
        st.error(f"❌ Fallo crítico en el procesamiento central: {str(e)}")
else:
    st.markdown("""
    ### 👋 Clasificador Bancario Inteligente Multi-Banco
    ## OPERACIONES DISPONIBLES
    ✅ **Bancos soportados:** Mercantil, Banco de Venezuela, Banesco, Provincial, BNC, Tesoro, Bancamiga.
    """)
