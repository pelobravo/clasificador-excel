import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import date
import requests
from bs4 import BeautifulSoup
import json
import re
from datetime import datetime
import unicodedata  # 🔥 NUEVO - Para normalizar texto

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
# 🔥 NUEVA FUNCIÓN PARA NORMALIZAR TEXTO (eliminar acentos)
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

        # =========================================
        # SI YA ES NUMÉRICO
        # =========================================

        if isinstance(valor, (int, float)):

            numero = float(valor)

            # =====================================
            # CORRECCIÓN BANCAMIGA
            # ENTEROS GRANDES SIN DECIMALES
            # =====================================

            if (
                isinstance(valor, int)
                and numero >= 100000
            ):

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

        # =========================================
        # FORMATO EUROPEO
        # 1.234,56
        # =========================================

        if "." in valor and "," in valor:

            valor = valor.replace(".", "")
            valor = valor.replace(",", ".")

        # =========================================
        # FORMATO CON COMA DECIMAL
        # 1234,56
        # =========================================

        elif "," in valor:

            valor = valor.replace(",", ".")

        numero = float(valor)

        # =====================================
        # ENTEROS INFLADOS
        # =====================================

        if (
            "." not in valor_original
            and "," not in valor_original
            and numero >= 100000
        ):

            numero = numero / 100

        return numero

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
# 🔥 DETECTAR COMISIONES - VERSIÓN MEJORADA CON NORMALIZACIÓN
# =========================================================

def es_comision(texto):

    # 🔥 NORMALIZAR TEXTO (eliminar acentos)
    texto = normalizar_texto(texto)

    texto = texto.strip()

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
        "debito automatico bancario",

        "com ",
        "com.",
        "com pago",
        "com pago otr",
        "com pago otr bcos",
        "comision pago proveedores",

        "descuento tarjeta",
        "descuento de tarjeta",
        "comision tarjeta",
        "comisión tarjeta",
        "cargo tarjeta",
        "retencion tarjeta",
        "retención tarjeta",
        "comision punto de venta",
        "comisión punto de venta",
        "punto de venta",
        "comision pos",
        "comisión pos",
        "descuento pos",
        "cargo por servicio",
        "cargo por transaccion",
        "cargo por transacción",
    ]

    return any(
        p in texto
        for p in palabras
    )

# =========================================================
# DETECTOR DE BANCO CORREGIDO - VERSIÓN MEJORADA
# =========================================================

def detectar_banco(nombre_archivo):

    nombre = nombre_archivo.upper()

    # =====================================
    # TESORO
    # =====================================

    if (
        "TESORO" in nombre
        or "TESORERIA" in nombre
        or "TES" in nombre
    ):

        return "tesoro"

    # =====================================
    # BANESCO
    # =====================================

    elif (
        "BANESCO" in nombre
        or re.match(r"^J\d+", nombre_archivo)
    ):

        return "banesco"

    # =====================================
    # BANCAMIGA
    # =====================================

    elif (
        "BANCAMIGA" in nombre
        or "MOVIMIENTOS_" in nombre
    ):

        return "bancamiga"

    # =====================================
    # VENEZUELA
    # =====================================

    elif (
        "MOVIMIENTOS EN MONEDA NACIONAL" in nombre
        or "VENEZUELA" in nombre
        or "BANCO DE VENEZUELA" in nombre
        or "BDV" in nombre
    ):

        return "venezuela"

    # =====================================
    # PROVINCIAL
    # =====================================

    elif "PROVINCIAL" in nombre:

        return "provincial"

    # =====================================
    # BNC
    # =====================================

    elif "BNC" in nombre:

        return "bnc"

    # =====================================
    # MERCANTIL
    # =====================================

    elif "MERCANTIL" in nombre:

        return "mercantil"

    # =====================================
    # DEFAULT
    # =====================================

    return "mercantil"

# =========================================================
# PROCESAR VENEZUELA - VERSIÓN MODIFICADA (CON ENCABEZADOS DIRECTOS)
# =========================================================

def procesar_venezuela(df):
    """Procesa archivo del Banco de Venezuela - VERSIÓN MODIFICADA"""
    
    st.info("Procesando Banco de Venezuela...")
    
    try:
        # ============================================
        # EL ARCHIVO YA VIENE CON ENCABEZADOS
        # ============================================
        
        df.columns = [
            str(c).strip()
            for c in df.columns
        ]
        
        st.write("Columnas originales:")
        st.write(df.columns.tolist())
        
        # ============================================
        # DETECTAR COLUMNAS POR NOMBRE O CONTENIDO
        # ============================================
        
        col_fecha = None
        col_ref = None
        col_desc = None
        col_credito = None
        col_debito = None
        col_tipo = None
        
        for col in df.columns:
            col_lower = str(col).lower()
            
            # Fecha
            if col_fecha is None:
                if "fecha" in col_lower or "fec" in col_lower or "fech" in col_lower:
                    col_fecha = col
                else:
                    # Buscar por contenido
                    sample = df[col].dropna().head(3).astype(str)
                    if sample.str.contains(r'\d{2}[/\-]\d{2}[/\-]\d{4}', regex=True).any():
                        col_fecha = col
            
            # Referencia
            if col_ref is None:
                if "referencia" in col_lower or "ref" in col_lower or "nro" in col_lower:
                    col_ref = col
            
            # Descripción
            if col_desc is None:
                if "descrip" in col_lower or "concepto" in col_lower or "detalle" in col_lower:
                    col_desc = col
            
            # Crédito / Haber (con y sin tilde)
            if col_credito is None:
                if (
                    "credito" in col_lower
                    or "crédito" in col_lower
                    or "haber" in col_lower
                ):
                    col_credito = col
            
            # Débito / Debe (con y sin tilde)
            if col_debito is None:
                if (
                    "debito" in col_lower
                    or "débito" in col_lower
                    or "cargo" in col_lower
                ):
                    col_debito = col
            
            # Tipo de movimiento
            if col_tipo is None:
                if "tipo" in col_lower and ("mov" in col_lower or "oper" in col_lower):
                    col_tipo = col
        
        # ============================================
        # CONSTRUIR DATAFRAME NORMALIZADO
        # ============================================
        
        movimientos = []
        
        for idx, fila in df.iterrows():
            try:
                # 🔥 FECHA - CORREGIDO CON dayfirst=True (en lugar de format fijo)
                fecha_val = None
                if col_fecha:
                    fecha_raw = fila[col_fecha]
                    if pd.notna(fecha_raw):
                        fecha_val = pd.to_datetime(fecha_raw, dayfirst=True, errors="coerce")
                
                if fecha_val is None or pd.isna(fecha_val):
                    continue
                
                # REFERENCIA
                referencia = ""
                if col_ref and pd.notna(fila[col_ref]):
                    referencia = str(fila[col_ref]).strip()
                
                # DESCRIPCIÓN
                descripcion = ""
                if col_desc and pd.notna(fila[col_desc]):
                    descripcion = str(fila[col_desc]).strip()
                
                # MONTO (determinar si es crédito o débito)
                monto = 0
                tipo = ""
                
                # Intentar obtener monto de columna crédito
                if col_credito and pd.notna(fila[col_credito]):
                    try:
                        val = fila[col_credito]
                        if isinstance(val, (int, float)):
                            monto = abs(float(val))
                        else:
                            monto = convertir_monto(val) or 0
                        if monto > 0:
                            tipo = "NC"
                    except:
                        pass
                
                # Si no hay crédito, intentar débito
                if monto == 0 and col_debito and pd.notna(fila[col_debito]):
                    try:
                        val = fila[col_debito]
                        if isinstance(val, (int, float)):
                            monto = abs(float(val))
                        else:
                            monto = convertir_monto(val) or 0
                        if monto > 0:
                            tipo = "ND"
                    except:
                        pass
                
                # Si aún no hay monto, intentar columna genérica de monto
                if monto == 0:
                    for col in df.columns:
                        col_lower = str(col).lower()
                        if "monto" in col_lower or "importe" in col_lower:
                            val = fila[col]
                            if pd.notna(val):
                                try:
                                    monto = convertir_monto(val) or 0
                                    break
                                except:
                                    pass
                
                # Si se encontró tipo de movimiento explícito
                if col_tipo and pd.notna(fila[col_tipo]):
                    tipo_raw = str(fila[col_tipo]).strip().upper()
                    if "CREDITO" in tipo_raw or "ABONO" in tipo_raw or "NC" in tipo_raw:
                        tipo = "NC"
                    elif "DEBITO" in tipo_raw or "DEBE" in tipo_raw or "ND" in tipo_raw:
                        tipo = "ND"
                
                # Si no hay tipo, inferir del contexto
                if tipo == "":
                    desc_upper = descripcion.upper()
                    if any(p in desc_upper for p in ["PAGO", "TRANSFERENCIA ENVIADA", "DEBITO", "COMPRA"]):
                        tipo = "ND"
                    elif any(p in desc_upper for p in ["DEPOSITO", "ABONO", "TRANSFERENCIA RECIBIDA", "CREDITO"]):
                        tipo = "NC"
                    else:
                        tipo = "ND" if monto > 0 else "NC"
                
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
        
        # ============================================
        # DEBUG: Mostrar información antes de retornar
        # ============================================
        st.write("🔍 DEBUG VENEZUELA:")
        st.write("Columnas detectadas:", list(df_resultado.columns))
        st.write("Cantidad de registros:", len(df_resultado))
        
        if len(df_resultado) > 0:
            st.write("Vista previa de los primeros registros:")
            st.dataframe(df_resultado.head(10))
        else:
            st.warning("⚠️ No se encontraron movimientos válidos")
        
        if df_resultado.empty:
            st.error("❌ No se encontraron movimientos válidos en el archivo.")
            st.info("""
            **Posibles causas:**
            - El archivo no es el original descargado del banco
            - El archivo fue modificado o re-exportado
            - El formato del archivo es diferente al esperado
            
            **Por favor:**
            1. Descargue el archivo NUEVAMENTE del Banco de Venezuela
            2. No lo abra ni modifique antes de cargarlo
            3. Cargue el archivo original (sin cambios)
            """)
            return pd.DataFrame()
        
        st.success(f"✅ Venezuela OK: {len(df_resultado)} movimientos detectados")
        
        return df_resultado
        
    except Exception as e:
        st.error(f"Error procesando Venezuela: {str(e)}")
        return pd.DataFrame()

# =========================================================
# PROCESAR BANESCO - NUEVA VERSIÓN
# =========================================================

def procesar_banesco(df):

    st.info("Procesando Banesco...")

    try:

        # ============================================
        # DEFINIR COLUMNAS DIRECTAMENTE (sin usar primera fila como header)
        # ============================================

        df.columns = [
            "FECHA",
            "REFERENCIA",
            "DESCRIPCION",
            "MONTO_RAW",
            "BALANCE"
        ]

        # ============================================
        # LIMPIAR COLUMNAS
        # ============================================

        df.columns = [
            str(c).strip().upper()
            for c in df.columns
        ]

        # ============================================
        # RENOMBRAR
        # ============================================

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

        # ============================================
        # VALIDAR COLUMNAS
        # ============================================

        for col in [
            "FECHA",
            "REFERENCIA",
            "DESCRIPCION",
            "MONTO_RAW"
        ]:

            if col not in df.columns:

                st.error(f"No existe columna: {col}")

                return pd.DataFrame()

        # ============================================
        # FECHA
        # ============================================

        df["FECHA"] = pd.to_datetime(
            df["FECHA"],
            dayfirst=True,
            errors="coerce"
        )

        df = df[
            df["FECHA"].notna()
        ]

        # ============================================
        # TIPO
        # ============================================

        df["TIPO"] = df["MONTO_RAW"].astype(str).apply(

            lambda x: "NC"
            if "+" in x
            else "ND"
        )

        # ============================================
        # LIMPIAR MONTO
        # ============================================

        df["MONTO"] = (

            df["MONTO_RAW"]

            .astype(str)

            .str.replace("+", "", regex=False)

            .str.replace("-", "", regex=False)

            .str.replace(".", "", regex=False)

            .str.replace(",", ".", regex=False)

            .str.strip()
        )

        df["MONTO"] = pd.to_numeric(
            df["MONTO"],
            errors="coerce"
        )

        # ============================================
        # LIMPIAR
        # ============================================

        df = df[
            df["MONTO"].notna()
        ]

        df = df[
            df["MONTO"] > 0
        ]

        # ============================================
        # COLUMNAS FINALES
        # ============================================

        df = df[[
            "FECHA",
            "REFERENCIA",
            "DESCRIPCION",
            "TIPO",
            "MONTO"
        ]]

        st.success(
            f"Banesco OK: {len(df)} movimientos"
        )

        st.dataframe(df.head())

        return df

    except Exception as e:

        st.error(f"Error Banesco: {str(e)}")

        return pd.DataFrame()

# =========================================================
# PROCESAR PROVINCIAL - CON DETECCIÓN DE TIPO REAL
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
        df["FECHA"] = pd.to_datetime(df["FECHA"], dayfirst=True, errors="coerce")
        df = df[df["FECHA"].notna()]
    
    if "MONTO" in df.columns:
        monto = df["MONTO"]
        if isinstance(monto, pd.DataFrame):
            monto = monto.iloc[:, 0]
        df["MONTO"] = pd.to_numeric(monto, errors="coerce")
        df = df[df["MONTO"].notna()]
    
    # =====================================================
    # TIPO TEMPORAL PARA BANCAMIGA
    # =====================================================
    
    if "TIPO" not in df.columns:
        df["TIPO"] = "ND"
    
    return df

# =========================================================
# PROCESAR BNC - SOLUCIÓN DEFINITIVA CON HEADERS ÚNICOS
# =========================================================

def procesar_bnc(df):

    st.info("Procesando archivo BNC...")

    # ============================================
    # BUSCAR FILA ENCABEZADO
    # ============================================

    encabezado = None

    for i in range(min(30, len(df))):

        fila = df.iloc[i].fillna("").astype(str)

        texto = " ".join(
            fila.tolist()
        ).lower()

        if (
            "fecha" in texto
            and (
                "descripcion" in texto
                or "descripción" in texto
            )
        ):

            encabezado = i
            break

    if encabezado is None:

        st.error(
            "No se encontró encabezado válido en BNC"
        )

        return pd.DataFrame()

    # ============================================
    # LIMPIAR Y CREAR HEADERS ÚNICOS
    # ============================================

    headers = []

    for idx, col in enumerate(df.iloc[encabezado]):

        col = str(col).strip()

        col = col.replace("\n", " ")

        if col == "" or col.lower() == "nan":
            col = f"COLUMNA_{idx}"

        headers.append(col)

    # ============================================
    # HACER COLUMNAS ÚNICAS
    # ============================================

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

    # ============================================
    # ASIGNAR COLUMNAS
    # ============================================

    df.columns = headers_unicos

    df = df.iloc[
        encabezado + 1:
    ].reset_index(drop=True)

    # ============================================
    # RENOMBRAR
    # ============================================

    rename_map = {}

    for col in df.columns:

        col_str = str(col).strip().lower()

        if "fecha" in col_str:

            rename_map[col] = "FECHA"

        elif (
            "descripcion" in col_str
            or "descripción" in col_str
            or "concepto" in col_str
        ):

            rename_map[col] = "DESCRIPCION"

        elif "referencia" in col_str:

            rename_map[col] = "REFERENCIA"

        elif "credito" in col_str or "haber" in col_str:

            rename_map[col] = "CREDITO"

        elif "debito" in col_str or "debe" in col_str:

            rename_map[col] = "DEBITO"

    df = df.rename(columns=rename_map)

    # ============================================
    # FECHA
    # ============================================

    if "FECHA" in df.columns:

        df["FECHA"] = pd.to_datetime(
            df["FECHA"],
            dayfirst=True,
            errors="coerce"
        )

        df = df[df["FECHA"].notna()]

    # ============================================
    # NUMÉRICOS
    # ============================================

    if "CREDITO" in df.columns:

        df["CREDITO"] = pd.to_numeric(
            df["CREDITO"],
            errors="coerce"
        ).fillna(0)

    else:

        df["CREDITO"] = 0

    if "DEBITO" in df.columns:

        df["DEBITO"] = pd.to_numeric(
            df["DEBITO"],
            errors="coerce"
        ).fillna(0)

    else:

        df["DEBITO"] = 0

    # ============================================
    # MONTO
    # ============================================

    df["MONTO"] = (
        df["CREDITO"]
        - df["DEBITO"]
    )

    df["TIPO"] = df["MONTO"].apply(

        lambda x: "NC" if x > 0 else "ND"
    )

    df["MONTO"] = df["MONTO"].abs()

    df = df[df["MONTO"] != 0]

    # ============================================
    # DEBUG
    # ============================================

    st.success(f"Registros BNC: {len(df)}")

    st.dataframe(df.head())

    return df

# =========================================================
# PROCESAR TESORO - VERSIÓN CORREGIDA CON MAP
# =========================================================

def procesar_tesoro(df):

    st.info("Procesando Banco del Tesoro...")

    try:

        # ============================================
        # BUSCAR FILA ENCABEZADO REAL
        # ============================================

        encabezado = None

        for i in range(min(20, len(df))):

            fila = df.iloc[i].astype(str)

            texto = " ".join(
                map(str, fila.tolist())
            ).lower()

            if (
                "fecha" in texto
                and "referencia" in texto
                and "concepto" in texto
            ):

                encabezado = i
                break

        if encabezado is None:

            st.error(
                "No se encontró encabezado válido en Tesoro"
            )

            return pd.DataFrame()

        # ============================================
        # USAR ESA FILA COMO HEADER
        # ============================================

        df.columns = df.iloc[encabezado]

        df = df.iloc[
            encabezado + 1:
        ].reset_index(drop=True)

        # ============================================
        # LIMPIAR COLUMNAS
        # ============================================

        df.columns = [
            str(c).strip()
            for c in df.columns
        ]

        # ============================================
        # RENOMBRAR
        # ============================================

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

        # ============================================
        # VALIDAR FECHA
        # ============================================

        if "FECHA" not in df.columns:

            st.error("No existe columna FECHA")

            return pd.DataFrame()

        # ============================================
        # FECHA
        # ============================================

        df["FECHA"] = pd.to_datetime(
            df["FECHA"],
            dayfirst=True,
            errors="coerce"
        )

        df = df[
            df["FECHA"].notna()
        ]

        # ============================================
        # NUMÉRICOS
        # ============================================

        def limpiar_numero(valor):

            valor = str(valor)

            valor = valor.replace(".", "")
            valor = valor.replace(",", ".")

            try:
                return float(valor)
            except:
                return 0

        if "CREDITO" in df.columns:

            df["CREDITO"] = df[
                "CREDITO"
            ].apply(limpiar_numero)

        else:

            df["CREDITO"] = 0

        if "DEBITO" in df.columns:

            df["DEBITO"] = df[
                "DEBITO"
            ].apply(limpiar_numero)

        else:

            df["DEBITO"] = 0

        # ============================================
        # MONTO
        # ============================================

        df["MONTO"] = (
            df["CREDITO"]
            - df["DEBITO"]
        )

        # ============================================
        # TIPO
        # ============================================

        df["TIPO"] = df["MONTO"].apply(

            lambda x: "NC"
            if x > 0
            else "ND"
        )

        # ============================================
        # ABS
        # ============================================

        df["MONTO"] = df["MONTO"].abs()

        # ============================================
        # LIMPIAR
        # ============================================

        df = df[
            df["MONTO"] > 0
        ]

        # ============================================
        # COLUMNAS FINALES
        # ============================================

        df = df[[
            "FECHA",
            "REFERENCIA",
            "DESCRIPCION",
            "TIPO",
            "MONTO"
        ]]

        st.success(
            f"Tesoro OK: {len(df)} registros"
        )

        st.dataframe(df.head())

        return df

    except Exception as e:

        st.error(
            f"Error Tesoro: {str(e)}"
        )

        return pd.DataFrame()

# =========================================================
# 🔥 PROCESAR BANCAMIGA - CON DETECCIÓN DE ENCABEZADO REAL Y FECHA CORREGIDA
# =========================================================

def procesar_bancamiga(df):

    st.info("Procesando Bancamiga...")

    try:

        # ============================================
        # BUSCAR ENCABEZADO REAL
        # ============================================

        encabezado = None

        for i in range(min(15, len(df))):

            fila = df.iloc[i].fillna("").astype(str)

            texto = " ".join(fila.tolist()).lower()

            if (
                "fecha" in texto
                and "referencia" in texto
                and "concepto" in texto
            ):

                encabezado = i
                break

        if encabezado is not None:

            df.columns = df.iloc[encabezado]

            df = df.iloc[
                encabezado + 1:
            ].reset_index(drop=True)

            df.columns = [
                str(c).strip()
                for c in df.columns
            ]

        # Aplanar columnas multinivel
        if isinstance(df.columns, pd.MultiIndex):

            df.columns = [
                col[-1]
                for col in df.columns
            ]

        # Renombrar columnas
        df = df.rename(columns={

            "Fecha": "FECHA",
            "Referencia": "REFERENCIA",
            "Concepto": "DESCRIPCION",
            "Débito": "DEBITO",
            "Crédito": "CREDITO"

        })

        # Validar fecha
        if "FECHA" not in df.columns:

            st.error("No se encontró columna FECHA")

            return pd.DataFrame()

        # 🔥 Convertir fecha - CORREGIDO con dayfirst=True en lugar de format fijo
        df["FECHA"] = pd.to_datetime(
            df["FECHA"],
            dayfirst=True,
            errors="coerce"
        )

        df = df[
            df["FECHA"].notna()
        ]

        # =========================================================
        # LIMPIAR NUMEROS BANCAMIGA - USANDO LA MISMA LÓGICA DE CONVERTIR_MONTO
        # =========================================================

        def limpiar_numero_bancamiga(valor):

            if pd.isna(valor):
                return 0

            try:

                # =========================================
                # SI YA ES NUMÉRICO
                # =========================================

                if isinstance(valor, (int, float)):

                    numero = float(valor)

                    # =====================================
                    # CORRECCIÓN BANCAMIGA
                    # ENTEROS GRANDES SIN DECIMALES
                    # =====================================

                    if (
                        isinstance(valor, int)
                        and numero >= 100000
                    ):

                        numero = numero / 100

                    return numero

                valor_original = str(valor).strip()

                valor = valor_original

                valor = valor.replace(" ", "")
                valor = valor.replace("$", "")
                valor = valor.replace("Bs", "")
                valor = valor.replace("€", "")

                if valor == "":
                    return 0

                # =========================================
                # FORMATO EUROPEO
                # 21.844,76
                # =========================================

                if "." in valor and "," in valor:

                    valor = valor.replace(".", "")
                    valor = valor.replace(",", ".")

                # =========================================
                # FORMATO SOLO COMA
                # 21844,76
                # =========================================

                elif "," in valor:

                    valor = valor.replace(",", ".")

                numero = float(valor)

                # =====================================
                # ENTEROS INFLADOS
                # =====================================

                if (
                    "." not in valor_original
                    and "," not in valor_original
                    and numero >= 100000
                ):

                    numero = numero / 100

                return numero

            except:

                return 0

        df["CREDITO"] = df["CREDITO"].apply(
            limpiar_numero_bancamiga
        )

        df["DEBITO"] = df["DEBITO"].apply(
            limpiar_numero_bancamiga
        )

        # Calcular monto
        df["MONTO"] = (
            df["CREDITO"] - df["DEBITO"]
        )

        # Crear tipo
        df["TIPO"] = df["MONTO"].apply(
            lambda x: "NC" if x > 0 else "ND"
        )

        # Absoluto
        df["MONTO"] = df["MONTO"].abs()

        # Limpiar
        df = df[df["MONTO"] > 0]

        columnas_finales = [
            "FECHA",
            "REFERENCIA",
            "DESCRIPCION",
            "TIPO",
            "MONTO"
        ]

        df = df[columnas_finales]

        st.success(
            f"Bancamiga OK: {len(df)} registros"
        )

        st.dataframe(df.head())

        return df

    except Exception as e:

        st.error(
            f"Error Bancamiga: {str(e)}"
        )

        return pd.DataFrame()

# =========================================================
# OBTENER TASA BCV (API o SCRAPING) - CORREGIDO
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
    }
    
    fecha_str = fecha_obj.strftime("%d/%m/%Y")
    
    if fecha_str in tasas_bcv_local:
        return tasas_bcv_local[fecha_str]
    
    # NO usar última tasa automáticamente
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
# PROCESAMIENTO MERCANTIL ORIGINAL (COMPLETO, SIN REGLAS ESPECIALES)
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
            
            # Intentar múltiples formatos de fecha
            fecha_obj = pd.to_datetime(
                fecha,
                dayfirst=True,
                errors="coerce"
            )
            
            # Validar que la fecha sea válida antes de buscar tasa
            if pd.isna(fecha_obj):
                st.warning(f"Fecha inválida: {fecha}, se omite")
                continue
            
            fecha_key = fecha_obj.strftime("%d/%m/%Y")
            
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

            # =================================================
            # 🔥 COMISIONES - AHORA USA LA FUNCIÓN MEJORADA
            # =================================================

            if es_comision(descripcion):

                comisiones.append(registro)

            # =================================================
            # INGRESOS - SOLO POR TIPO (sin reglas especiales)
            # =================================================

            elif tipo in tipos_ingresos:

                ingresos.append(registro)

            # =================================================
            # EGRESOS
            # =================================================

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
    
    # Verificar columnas antes de convertir
    st.write(f"📋 COLUMNAS ANTES DE CONVERTIR {banco.upper()}:")
    st.write(df.columns.tolist())
    
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
# INTERFAZ PRINCIPAL - CON LECTURA MEJORADA DE PROVINCIAL
# =========================================================

df_ipago = None

if archivo:

    st.info(
        f"📄 Archivo: **{archivo.name}** "
        f"- {archivo.size/1024:.1f} KB"
    )

    try:
        # =========================================================
        # DETECTAR BANCO
        # =========================================================
        
        # SOLUCIÓN 2: Detectar por contenido HTML
        archivo.seek(0)
        primeros_bytes = archivo.read(100)
        archivo.seek(0)
        
        if b"<table" in primeros_bytes.lower():
            banco = "banesco"
        else:
            banco = detectar_banco(archivo.name)
        
        # st.info(f"🏦 Banco detectado: **{banco.upper()}**")
        
        # =========================================================
        # LEER Y PROCESAR SEGÚN BANCO
        # =========================================================
        
        if banco == "mercantil":
            # MERCANTIL: usar el procesamiento original con lectura sin encabezados
            df_original = leer_excel_sin_encabezados(archivo)
            usar_procesamiento_original = True
            
        elif banco == "banesco":

            try:

                nombre = archivo.name.lower()

                # ============================================
                # SI ES XLSX REAL
                # ============================================

                if nombre.endswith(".xlsx") or nombre.endswith(".xlsm"):

                    df_raw = pd.read_excel(
                        archivo,
                        engine="openpyxl",
                        header=None
                    )

                # ============================================
                # SI ES HTML DISFRAZADO
                # ============================================

                else:

                    tablas = pd.read_html(archivo)

                    if len(tablas) == 0:

                        st.error("No se encontraron tablas en Banesco")
                        st.stop()

                    df_raw = tablas[0]

                st.success(
                    f"✓ Banesco: {len(df_raw)} registros encontrados"
                )

                st.dataframe(df_raw.head())

                df_normalizado = procesar_banesco(df_raw)

                df_original = convertir_a_formato_mercantil(
                    df_normalizado,
                    banco
                )

            except Exception as e:

                st.error(
                    f"Error leyendo Banesco: {str(e)}"
                )

                st.stop()
            
        elif banco == "tesoro":
            # TESORO: leer Excel real
            try:
                df_raw = pd.read_excel(
                    archivo,
                    engine="openpyxl"
                )
                st.success(f"✓ Tesoro: {len(df_raw)} registros encontrados")
                st.dataframe(df_raw.head())
            except Exception as e:
                st.error(f"Error leyendo Tesoro: {str(e)}")
                st.stop()
            
            df_normalizado = procesar_tesoro(df_raw)
            df_original = convertir_a_formato_mercantil(df_normalizado, banco)
            
        elif banco == "bancamiga":

            try:

                nombre = archivo.name.lower()

                # ============================================
                # PRIMER INTENTO: HTML DISFRAZADO
                # ============================================

                try:

                    tablas = pd.read_html(archivo)

                    if len(tablas) > 0:

                        df_raw = tablas[0]

                        st.success(
                            f"✓ Bancamiga HTML: {len(df_raw)} registros"
                        )

                    else:

                        raise Exception(
                            "Sin tablas HTML"
                        )

                # ============================================
                # SEGUNDO INTENTO: EXCEL REAL
                # ============================================

                except:

                    archivo.seek(0)

                    if nombre.endswith(".xls"):

                        df_raw = pd.read_excel(
                            archivo,
                            engine="xlrd",
                            header=5
                        )

                    else:

                        df_raw = pd.read_excel(
                            archivo,
                            engine="openpyxl",
                            header=0
                        )

                    st.success(
                        f"✓ Bancamiga Excel: {len(df_raw)} registros"
                    )

                st.dataframe(df_raw.head())

                # ============================================
                # PROCESAR
                # ============================================

                df_normalizado = procesar_bancamiga(df_raw)

                df_original = convertir_a_formato_mercantil(
                    df_normalizado,
                    banco
                )

            except Exception as e:

                st.error(
                    f"Error leyendo Bancamiga: {str(e)}"
                )

                st.stop()
            
        elif banco == "provincial":

            try:

                movimientos = []

                # =====================================================
                # INTENTO 1 - LEER COMO TEXTO
                # =====================================================

                archivo.seek(0)

                try:

                    contenido = archivo.read().decode(
                        "latin-1",
                        errors="ignore"
                    )

                except:

                    archivo.seek(0)

                    contenido = archivo.read().decode(
                        "utf-8",
                        errors="ignore"
                    )

                lineas = contenido.splitlines()

                referencia_actual = ""
                leyendo_movimiento = False
                descripcion = ""

                for linea in lineas:

                    linea = linea.strip()

                    if not linea:
                        continue

                    # =========================================
                    # REFERENCIA
                    # =========================================

                    if re.search(r"'?\d{8,20}", linea):

                        match = re.search(
                            r"(\d{8,20})",
                            linea
                        )

                        if match:

                            referencia_actual = match.group(1)

                    # =========================================
                    # MOVIMIENTO
                    # =========================================

                    if (
                        "TRAV" in linea.upper()
                        or "CR./" in linea.upper()
                        or "COMIS" in linea.upper()
                        or "PAGO" in linea.upper()
                        or "TRANSFER" in linea.upper()
                    ):

                        descripcion = linea
                        leyendo_movimiento = True
                        continue

                    # =========================================
                    # FECHA + MONTO
                    # =========================================

                    if not leyendo_movimiento:
                        continue

                    # =========================================
                    # FILTRAR SALDOS
                    # =========================================

                    texto_upper = linea.upper()

                    if (
                        "SALDO INICIAL" in texto_upper
                        or "SALDO FINAL" in texto_upper
                    ):
                        continue

                    fecha_match = re.search(
                        r"(\d{2}-\d{2}-\d{4})",
                        linea
                    )

                    monto_match = re.search(
                        r"([\d\.]+,\d{2})",
                        linea
                    )

                    if fecha_match and monto_match:

                        fecha = fecha_match.group(1)

                        monto_txt = monto_match.group(1)

                        monto = convertir_monto(
                            monto_txt
                        )

                        if monto is None:
                            continue

                        descripcion_final = (
                            descripcion
                            if 'descripcion' in locals()
                            else ""
                        )

                        descripcion_upper = (
                            descripcion_final.upper()
                        )

                        # =====================================
                        # TIPO
                        # =====================================

                        if es_comision(
                            descripcion_upper
                        ):

                            tipo = "COMISION"

                        elif (
                            "TRAV" in descripcion_upper
                            or "CR./" in descripcion_upper
                        ):

                            tipo = "NC"

                        else:

                            tipo = "ND"

                        movimientos.append({

                            "FECHA": fecha,

                            "REFERENCIA": referencia_actual,

                            "DESCRIPCION": descripcion_final,

                            "TIPO": tipo,

                            "MONTO": abs(monto)

                        })

                        # Resetear flag para seguir buscando próximos movimientos
                        leyendo_movimiento = False
                        descripcion = ""

                df_normalizado = pd.DataFrame(
                    movimientos
                )

                st.success(
                    f"✓ Provincial: {len(df_normalizado)} movimientos detectados"
                )

                st.dataframe(
                    df_normalizado.head()
                )

                if df_normalizado.empty:

                    st.error(
                        "No se detectaron movimientos válidos en Provincial."
                    )

                    st.stop()

                df_original = convertir_a_formato_mercantil(
                    df_normalizado,
                    banco
                )

            except Exception as e:

                st.error(
                    f"Error leyendo Provincial: {str(e)}"
                )

                st.stop()
            
        elif banco == "venezuela":
            # Venezuela: usar el parser mejorado con encabezados
            df_raw = leer_excel_con_encabezados(archivo)
            
            # PRUEBA VENEZUELA - Mostrar información de depuración
            st.write("PRUEBA VENEZUELA")
            st.write(df_raw.columns.tolist())
            st.dataframe(df_raw.head())
            
            # Procesar con función mejorada
            df_normalizado = procesar_venezuela(df_raw)
            
            if df_normalizado.empty:
                st.stop()
            
            # Convertir al formato que espera procesar_archivo
            df_original = convertir_a_formato_mercantil(df_normalizado, banco)
            
        elif banco == "bnc":
            df_raw = leer_excel_con_encabezados(archivo)
            df_normalizado = procesar_bnc(df_raw)
            df_original = convertir_a_formato_mercantil(df_normalizado, banco)
            
        else:
            # OTROS BANCOS: aplicar parser específico
            df_raw = leer_excel_con_encabezados(archivo)
            df_normalizado = df_raw
            
            # Convertir al formato que espera procesar_archivo
            df_original = convertir_a_formato_mercantil(df_normalizado, banco)
            
        # ============================================
        # FILTRAR POR FECHAS - CORREGIDO
        # ============================================

        # Verificar que hay datos antes de filtrar
        if df_original.empty:
            st.error(
                "No se detectaron movimientos para procesar."
            )
            st.stop()

        try:
            if banco == "mercantil":
                # 🔥 CORREGIDO: Aplicar zfill(8) para fechas con 7 dígitos
                fechas_convertidas = pd.to_datetime(
                    df_original[3].astype(str).str.zfill(8),
                    format="%d%m%Y",
                    errors="coerce"
                )
            else:
                # Otros bancos usan formato con separadores (ej: 14/05/2026)
                fechas_convertidas = pd.to_datetime(
                    df_original.iloc[:, 3],
                    errors="coerce",
                    dayfirst=True
                )

            # Convertir fechas del filtro a datetime
            fecha_inicio_dt = pd.to_datetime(fecha_inicio)
            fecha_fin_dt = pd.to_datetime(fecha_fin)

            # DEBUG: Mostrar información de fechas
            st.write("DEBUG FECHAS:")
            st.write("Min fecha archivo:", fechas_convertidas.min())
            st.write("Max fecha archivo:", fechas_convertidas.max())
            st.write("Registros antes de filtrar:", len(df_original))

            # Filtrar correctamente
            df_original = df_original[
                (fechas_convertidas >= fecha_inicio_dt) & 
                (fechas_convertidas <= fecha_fin_dt)
            ]

            st.write("Registros después de filtrar:", len(df_original))

            st.success(
                f"Filtro de fechas aplicado: {fecha_inicio} a {fecha_fin}"
            )

        except Exception as e:
            st.warning(f"Error filtrando fechas: {e}")
            
        # Verificar que se pudieron convertir los datos
        if df_original.empty or len(df_original) == 0:
            st.error("""

❌ No se encontraron movimientos válidos.

Posibles causas:

• El archivo ya fue procesado anteriormente por el sistema.
• El archivo no es el original descargado del banco.
• El formato del banco no coincide con el esperado.

Por favor cargue el archivo ORIGINAL del banco.

""")
            st.stop()

        with st.expander("👁️ Vista previa archivo original"):

            st.dataframe(
                df_original.head(20),
                use_container_width=True
            )

        # =========================================================
        # LEER ARCHIVO IPAGO
        # =========================================================
        
        if archivo_ipago:
            try:
                df_ipago = pd.read_excel(
                    archivo_ipago,
                    engine="openpyxl"
                )
                st.success(f"Archivo iPago cargado: {len(df_ipago)} registros")
                
                # Limpiar referencia iPago - CONVERTIR A ENTERO REAL
                df_ipago["Referencia"] = (
                    pd.to_numeric(
                        df_ipago["Referencia"],
                        errors="coerce"
                    )
                    .fillna(0)
                    .astype("Int64")
                    .astype(str)
                )
                
            except Exception as e:
                st.error(f"Error leyendo archivo iPago: {e}")

        if procesar:

            with st.spinner("Procesando archivo con tasas BCV..."):

                if banco == "venezuela":
                    ingresos = []
                    egresos = []
                    comisiones = []

                    for _, row in df_normalizado.iterrows():
                        # ============================================
                        # OBTENER FECHA - CORREGIDO CON dayfirst=True
                        # ============================================
                        fecha_obj = pd.to_datetime(
                            row["FECHA"],
                            dayfirst=True,
                            errors="coerce"
                        )

                        # ============================================
                        # OBTENER TASA BCV
                        # ============================================
                        tasa = obtener_tasa_por_fecha(
                            fecha_obj,
                            usar_api
                        )

                        if tasa is None:
                            tasa = 1.0

                        # ============================================
                        # CALCULAR USD
                        # ============================================
                        monto_bs = float(row["MONTO"])
                        monto_usd = calcular_usd(monto_bs, tasa)

                        registro = {
                            "FECHA": row["FECHA"],
                            "REFERENCIA": row["REFERENCIA"],
                            "DESCRIPCIÓN": row["DESCRIPCION"],
                            "MONTO BS": monto_bs,
                            "TASA BCV": tasa,
                            "MONTO USD": monto_usd
                        }

                        tipo = str(row["TIPO"]).strip().upper()
                        descripcion = str(row["DESCRIPCION"]).strip()

                        # ============================================
                        # 🔥 COMISIONES - AHORA USA LA FUNCIÓN MEJORADA
                        # ============================================
                        if es_comision(descripcion):
                            comisiones.append(registro)
                            continue

                        # ============================================
                        # INGRESOS
                        # ============================================
                        if tipo in ["NC", "C", "CREDITO", "ABONO"]:
                            ingresos.append(registro)

                        # ============================================
                        # EGRESOS
                        # ============================================
                        elif tipo in ["ND", "D", "DEBITO", "DEBIT"]:
                            egresos.append(registro)

                        # ============================================
                        # DEFAULT
                        # ============================================
                        else:
                            egresos.append(registro)
                else:
                    ingresos, egresos, comisiones = procesar_archivo(df_original, usar_api)

            df_ingresos = pd.DataFrame(ingresos)
            df_egresos = pd.DataFrame(egresos)
            df_comisiones = pd.DataFrame(comisiones)

            # =========================================================
            # NORMALIZAR REFERENCIAS BANCO
            # =========================================================

            if not df_egresos.empty:

                df_egresos["REFERENCIA"] = (

                    df_egresos["REFERENCIA"]

                    .astype(str)

                    .str.replace(".0", "", regex=False)

                    .str.replace(" ", "", regex=False)

                    .str.strip()

                )

                # Tomar últimos 6 dígitos
                df_egresos["REF_CRUCE"] = (
                    df_egresos["REFERENCIA"]
                    .str[-6:]
                )

            # =========================================================
            # NORMALIZAR REFERENCIAS IPAGO
            # =========================================================

            if df_ipago is not None:

                df_ipago["Referencia"] = (

                    df_ipago["Referencia"]

                    .astype(str)

                    .str.replace(".0", "", regex=False)

                    .str.replace(" ", "", regex=False)

                    .str.strip()

                )

                # Tomar últimos 6 dígitos
                df_ipago["REF_CRUCE"] = (
                    df_ipago["Referencia"]
                    .str[-6:]
                )

            # =========================================================
            # DEBUG
            # =========================================================

            if df_ipago is not None and not df_egresos.empty:

                st.write("🔍 REFERENCIAS BANCO:")
                st.write(df_egresos[["REFERENCIA", "REF_CRUCE"]].head())

                st.write("🔍 REFERENCIAS IPAGO:")
                st.write(df_ipago[["Referencia", "REF_CRUCE"]].head())
            
            # =========================================================
            # HACER EL CRUCE CON IPAGO
            # =========================================================
            
            if df_ipago is not None and not df_egresos.empty:
                df_egresos = df_egresos.merge(
                    df_ipago,
                    on="REF_CRUCE",
                    how="left"
                )
                st.success("Cruce con iPago realizado correctamente")
                
                # =========================================================
                # REEMPLAZAR DESCRIPCIÓN CON DATOS DE IPAGO
                # =========================================================
                
                # Reemplazar descripción del banco con la descripción de iPago (si existe)
                if "Descripción" in df_egresos.columns:
                    df_egresos["DESCRIPCIÓN"] = (
                        df_egresos["Descripción"]
                        .fillna(df_egresos["DESCRIPCIÓN"])
                    )
                
                # Reemplazar beneficiario con Proveedor de iPago (si existe)
                if "Proveedor" in df_egresos.columns:
                    df_egresos["BENEFICIARIO"] = (
                        df_egresos["Proveedor"]
                        .fillna(df_egresos.get("BENEFICIARIO", ""))
                    )
                
                # Crear columna CONCEPTO con el Tipo de Egreso de iPago
                if "Tipo de Egreso" in df_egresos.columns:
                    df_egresos["CONCEPTO"] = df_egresos["Tipo de Egreso"]

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

                        # FORMATO ORIGINAL RESTAURADO
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

                    # TOTAL BS - FORMATO ORIGINAL RESTAURADO
                    total_bs_cell = hoja.cell(
                        row=fila_data,
                        column=4
                    )

                    total_bs_cell.value = dataframe[
                        "MONTO BS"
                    ].sum()

                    total_bs_cell.number_format = '#,##0.00'
                    total_bs_cell.fill = color_total

                    # TOTAL USD
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
    - Bancamiga

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
