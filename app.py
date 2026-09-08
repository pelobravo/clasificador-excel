import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import date
import requests
from bs4 import BeautifulSoup
import json
import re
import os
from datetime import datetime, timedelta
import unicodedata
from collections import Counter

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
if "saldo_banplus" not in st.session_state: st.session_state.saldo_banplus = 0.0
if "saldo_activo" not in st.session_state: st.session_state.saldo_activo = 0.0
if "saldo_efectivo" not in st.session_state: st.session_state.saldo_efectivo = 0.0
if "saldo_binance" not in st.session_state: st.session_state.saldo_binance = 0.0
if "total_ingresos_consolidado" not in st.session_state: st.session_state.total_ingresos_consolidado = 0.0
if "total_egresos_ipago_ves" not in st.session_state: st.session_state.total_egresos_ipago_ves = 0.0
if "info_fechas_por_banco" not in st.session_state: st.session_state.info_fechas_por_banco = {}
if "total_creditos_venezuela" not in st.session_state: st.session_state.total_creditos_venezuela = 0.0
if "creditos_por_banco" not in st.session_state: st.session_state.creditos_por_banco = {}
if "resumen_bancos" not in st.session_state: st.session_state.resumen_bancos = {}

_BANCOS_RESUMEN = ["Banesco", "BNC", "Mercantil", "Banco de Venezuela (BDV)", "Provincial", "Bancamiga", "BanPlus", "Banco Activo", "Banco del Tesoro"]

def _nombre_banco_resumen(banco):
    """Convierte la clave interna de banco (banesco, banplus, ...) al nombre mostrado en reportes"""
    return {
        "banesco": "Banesco", "bnc": "BNC", "mercantil": "Mercantil",
        "venezuela": "Banco de Venezuela (BDV)", "provincial": "Provincial",
        "bancamiga": "Bancamiga", "banplus": "BanPlus", "activo": "Banco Activo",
        "tesoro": "Banco del Tesoro"
    }.get(str(banco).lower(), str(banco).capitalize())

def _acumular_resumen_banco(nombre, resumen):
    """Acumula (por archivo/día) el resumen de saldos de un banco en session_state"""
    if not isinstance(resumen, dict):
        return
    st.session_state.resumen_bancos.setdefault(nombre, {
        "saldo_inicial": 0.0, "saldo_final": 0.0, "creditos_total": 0.0, "debitos_total": 0.0
    })
    for k in ["saldo_inicial", "saldo_final", "creditos_total", "debitos_total"]:
        v = resumen.get(k)
        if v:
            st.session_state.resumen_bancos[nombre][k] += v

def _acumular_info_fechas(nombre, fecha, registros, total_original, excluidos, detalle_fechas):
    """Acumula la info de fechas de varios archivos (días) del mismo banco en session_state"""
    info = st.session_state.info_fechas_por_banco.setdefault(nombre, {
        "fechas": [], "registros": 0, "total_original": 0, "excluidos": 0, "detalle_fechas": {}
    })
    if fecha:
        if fecha not in info["fechas"]:
            info["fechas"].append(fecha)
        info["fecha"] = fecha
    info["registros"] += registros or 0
    info["total_original"] += total_original or 0
    info["excluidos"] += excluidos or 0
    for f, c in (detalle_fechas or {}).items():
        info["detalle_fechas"][f] = info["detalle_fechas"].get(f, 0) + (c or 0)
def _fecha_archivo(nombre_archivo):
    """Extrae la fecha del nombre del archivo (patrón dd-mm o dd/mm, ej. 'Edo cta Provincial 10-08.xls').
    Permite ordenar los archivos de un banco y usar el saldo del día más reciente."""
    try:
        m = re.search(r"(\d{2})[-/](\d{2})(?:[-/](\d{4}))?", str(nombre_archivo))
        if not m:
            return None
        d, mo = int(m.group(1)), int(m.group(2))
        y = int(m.group(3)) if m.group(3) else date.today().year
        return datetime(y, mo, d).date()
    except Exception:
        return None

def _fijar_saldo_ultimo_dia(datos):
    """Regla 'saldo final = último día': para un banco con N archivos (días),
    el saldo final es el del archivo de fecha MÁS RECIENTE (si 2 archivos comparten
    la fecha máxima se suman). El saldo inicial del resumen es el del archivo más
    antiguo y los créditos/débitos totales se suman del período completo."""
    if not datos:
        return 0.0, {"saldo_inicial": None, "saldo_final": None, "creditos_total": None, "debitos_total": None}
    fecha_base = datetime(1900, 1, 1).date()
    fechas = [(d.get("fecha") or fecha_base) for d in datos]
    max_fecha = max(fechas)
    min_fecha = min(fechas)
    saldo_final = sum(d["saldo"] for d, f in zip(datos, fechas) if f == max_fecha)
    resumen = {"saldo_inicial": None, "saldo_final": None, "creditos_total": None, "debitos_total": None}
    for d, f in zip(datos, fechas):
        r = d.get("resumen") or {}
        if f == min_fecha and r.get("saldo_inicial"):
            resumen["saldo_inicial"] = (resumen["saldo_inicial"] or 0) + r["saldo_inicial"]
        if f == max_fecha and r.get("saldo_final"):
            resumen["saldo_final"] = (resumen["saldo_final"] or 0) + r["saldo_final"]
        for k in ("creditos_total", "debitos_total"):
            if r.get(k):
                resumen[k] = (resumen[k] or 0) + r[k]
    return saldo_final, resumen


def _sumar_creditos_convertidos(df_convertido):
    """Suma los ingresos (créditos) de un dataframe convertido al formato estándar.
    Usa la misma clasificación que procesar_archivo: TIPO en tipos de ingreso,
    excluye comisiones (col 9) y filas de encabezados/saldos."""
    total = 0.0
    if df_convertido is None or getattr(df_convertido, "empty", True):
        return total
    tipos_ing = {"NC", "C", "CREDITO", "ABONO", "DP", "DEP"}
    excluir_texto = {"SALDO", "DESCRIPCION", "DESCRIPCIÓN", "REFERENCIA", "MOVIMIENTO", "FECHA", "SALDO INICIAL", "SALDO FINAL"}
    for _, fila in df_convertido.iterrows():
        try:
            if len(fila) > 9 and bool(fila[9]):
                continue
            tipo = str(fila[5]).strip().upper()
            if tipo not in tipos_ing:
                continue
            desc = str(fila[6]).strip().upper()
            if desc in excluir_texto:
                continue
            try:
                m = float(str(fila[7]).replace(",", ""))
            except Exception:
                m = 0.0
            if m > 0:
                total += m
        except Exception:
            continue
    return total

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
# 🔥 VALIDACIÓN DE FECHAS - DETECCIÓN DE FECHA PREDOMINANTE
# =========================================================

_PALABRAS_ENCABEZADO = [
    "SALDO INICIAL", "SALDO FINAL", "TOTAL CRÉDITO", "TOTAL DEBITO",
    "TOTAL CREDITO", "TOTAL DÉBITO", "SALDO PROMEDIO", "PERÍODO", "PERIODO",
    "FECHA", "REFERENCIA", "DESCRIPCIÓN", "DESCRIPCION", "MOVIMIENTO",
    "NRO", "Nº", "TIPO DE MOVIMIENTO", "CRÉDITOS:", "CREDITOS:",
    "DÉBITOS:", "DEBITOS:", "TOTAL CREDITOS", "TOTAL DEBITOS",
    "TOTAL CRÉDITOS", "TOTAL DÉBITOS"
]

def _es_fila_encabezado(fila_completa):
    """Detecta filas de encabezado/totales comparando con LÍMITES DE PALABRA.

    El chequeo anterior usaba subcadenas ('NRO' in fila), lo que descartaba
    movimientos legítimos cuyo texto contenía esas letras dentro de una palabra
    (p.ej. beneficiario 'JIANRONG WU' contiene 'NRO'). Con límites de palabra,
    'NRO' solo coincide si aparece como término aislado del encabezado."""
    for palabra in _PALABRAS_ENCABEZADO:
        if re.search(r"(?<!\w)" + re.escape(palabra) + r"(?!\w)", fila_completa):
            return True
    return False

def detectar_fecha_predominante(df_raw, columna_fecha_idx=0):
    """
    Detecta la fecha más frecuente en un archivo de estado de cuenta.
    AHORA EXCLUYE FILAS QUE NO SON MOVIMIENTOS VÁLIDOS.
    
    Args:
        df_raw: DataFrame con los datos del banco
        columna_fecha_idx: Índice de la columna donde están las fechas (por defecto 0)
    
    Returns:
        tuple: (fecha_predominante, dict_conteo_fechas, porcentaje_predominante, total_movimientos_validos)
    """
    try:
        # Extraer todas las fechas de la columna especificada
        fechas = []
        filas_invalidas = 0
        
        for idx in range(len(df_raw)):
            try:
                valor = df_raw.iloc[idx, columna_fecha_idx]
                if pd.isna(valor):
                    filas_invalidas += 1
                    continue
                
                # Convertir a string y limpiar
                fecha_str = str(valor).strip()
                if not fecha_str or fecha_str.lower() == 'nan':
                    filas_invalidas += 1
                    continue
                
                # 🔥 EXCLUIR FILAS QUE NO SON MOVIMIENTOS
                # Verificar si la fila contiene palabras clave de encabezados/totales
                # (búsqueda con límites de palabra: 'NRO' dentro de 'JIANRONG' NO es encabezado)
                es_encabezado = False
                try:
                    # Revisar toda la fila para detectar palabras de totales/encabezados
                    fila_completa = " ".join([str(v) for v in df_raw.iloc[idx].tolist()]).upper()
                    es_encabezado = _es_fila_encabezado(fila_completa)
                except:
                    pass
                
                if es_encabezado:
                    filas_invalidas += 1
                    continue
                
                # Intentar parsear la fecha
                fecha_dt = None
                
                # Intentar diferentes formatos
                formatos = [
                    "%Y-%m-%d %H:%M:%S", "%Y-%m-%d",
                    "%d/%m/%Y", "%d/%m/%y", "%d-%m-%Y", "%d-%m-%y",
                    "%Y/%m/%d", "%m/%d/%Y", "%m/%d/%y"
                ]
                
                for fmt in formatos:
                    try:
                        fecha_dt = pd.to_datetime(fecha_str, format=fmt)
                        break
                    except:
                        continue
                
                if fecha_dt is None:
                    # Último intento con dayfirst=True
                    fecha_dt = pd.to_datetime(fecha_str, dayfirst=True, errors='coerce')
                
                if pd.notna(fecha_dt):
                    # Normalizar a solo fecha (sin hora)
                    fechas.append(fecha_dt.date())
                else:
                    filas_invalidas += 1
                    
            except Exception as e:
                filas_invalidas += 1
                continue
        
        if not fechas:
            return None, {}, 0.0, 0
        
        # Contar frecuencias de cada fecha
        conteo = Counter(fechas)
        
        # Ordenar por frecuencia (descendente)
        fechas_ordenadas = conteo.most_common()
        
        if not fechas_ordenadas:
            return None, {}, 0.0, 0
        
        # Fecha predominante
        fecha_predominante = fechas_ordenadas[0][0]
        total_filas_validas = len(fechas)
        cantidad_predominante = fechas_ordenadas[0][1]
        porcentaje = (cantidad_predominante / total_filas_validas) * 100 if total_filas_validas > 0 else 0.0
        
        # Crear diccionario con todas las fechas y sus conteos
        dict_conteo = {fecha.strftime("%d/%m/%Y"): count for fecha, count in conteo.items()}
        
        return fecha_predominante, dict_conteo, porcentaje, total_filas_validas
        
    except Exception as e:
        st.warning(f"⚠️ Error al detectar fechas predominantes: {str(e)}")
        return None, {}, 0.0, 0

def _es_fila_saldo(fila):
    """Detecta filas de resumen del estado de cuenta (Saldo Inicial/Final, Créditos/Débitos Total).
    Estas filas no son movimientos: se capturan para el resumen de saldos y el exportable."""
    try:
        fila_completa = " ".join([str(v) for v in fila]).upper()
        return any(p in fila_completa for p in [
            "SALDO INICIAL", "SALDO FINAL", "CREDITOS TOTAL", "CRÉDITOS TOTAL",
            "CREDITO TOTAL", "CRÉDITO TOTAL", "DEBITOS TOTAL", "DÉBITOS TOTAL",
            "DEBITO TOTAL", "DÉBITO TOTAL"
        ])
    except:
        return False

def filtrar_por_fecha_predominante(df_raw, columna_fecha_idx=0, nombre_banco="banco"):
    """
    Filtra el DataFrame para conservar solo los registros de la fecha predominante.
    AHORA EXCLUYE FILAS QUE NO SON MOVIMIENTOS VÁLIDOS.
    
    Args:
        df_raw: DataFrame original
        columna_fecha_idx: Índice de la columna de fechas
        nombre_banco: Nombre del banco para mensajes
    
    Returns:
        tuple: (df_filtrado, fecha_predominante, dict_conteo, total_filas_validas, registros_excluidos)
    """
    # Detectar fecha predominante (ahora excluye filas inválidas)
    fecha_pred, dict_conteo, porcentaje, total_filas_validas = detectar_fecha_predominante(
        df_raw, columna_fecha_idx
    )
    
    if fecha_pred is None:
        st.warning(f"⚠️ {nombre_banco}: No se pudieron detectar fechas válidas. Se procesará todo el archivo.")
        return df_raw, None, {}, len(df_raw), 0
    
    # Mostrar información al usuario
    fechas_texto = ", ".join([f"{fecha}: {count} registros" for fecha, count in sorted(dict_conteo.items())])
    
    st.info(f"""
    📅 **{nombre_banco} - Análisis de fechas:**
    - Total de movimientos válidos: {total_filas_validas}
    - Fechas encontradas: {fechas_texto}
    - **Fecha predominante: {fecha_pred.strftime('%d/%m/%Y')}** ({porcentaje:.1f}% de los movimientos)
    """)
    
    # 🔥 IDENTIFICAR FILAS QUE SERÁN EXCLUIDAS
    filas_excluidas = []
    filas_con_fecha_correcta = []
    
    for idx in range(len(df_raw)):
        try:
            valor = df_raw.iloc[idx, columna_fecha_idx]
            if pd.isna(valor):
                es_fila_saldo = _es_fila_saldo(df_raw.iloc[idx].tolist())
                filas_excluidas.append({
                    'indice': idx,
                    'razon': 'Saldo del estado de cuenta (capturado)' if es_fila_saldo else 'Fecha vacía o NaN',
                    'fila': df_raw.iloc[idx].tolist()
                })
                continue
            
            fecha_str = str(valor).strip()
            if not fecha_str or fecha_str.lower() == 'nan':
                es_fila_saldo = _es_fila_saldo(df_raw.iloc[idx].tolist())
                filas_excluidas.append({
                    'indice': idx,
                    'razon': 'Saldo del estado de cuenta (capturado)' if es_fila_saldo else 'Fecha vacía o NaN',
                    'fila': df_raw.iloc[idx].tolist()
                })
                continue
            
            # Verificar si es encabezado
            # (búsqueda con límites de palabra: 'NRO' dentro de 'JIANRONG' NO es encabezado)
            es_encabezado = False
            try:
                fila_completa = " ".join([str(v) for v in df_raw.iloc[idx].tolist()]).upper()
                es_encabezado = _es_fila_encabezado(fila_completa)
            except:
                pass
            
            if es_encabezado:
                filas_excluidas.append({
                    'indice': idx,
                    'razon': 'Es encabezado o total',
                    'fila': df_raw.iloc[idx].tolist()
                })
                continue
            
            # Intentar convertir a fecha (ISO explícito primero: pandas 3 intercambia mes/día con dayfirst en ISO)
            m_iso = re.match(r"^(\d{4})-(\d{2})-(\d{2})", fecha_str)
            if m_iso:
                fecha_dt = pd.to_datetime(f"{m_iso.group(1)}-{m_iso.group(2)}-{m_iso.group(3)}", errors="coerce")
            else:
                fecha_dt = pd.to_datetime(fecha_str, dayfirst=True, errors='coerce')
            if pd.isna(fecha_dt):
                filas_excluidas.append({
                    'indice': idx,
                    'razon': f'Fecha no reconocida: {fecha_str}',
                    'fila': df_raw.iloc[idx].tolist()
                })
                continue
            
            # Si la fecha no coincide con la predominante
            if fecha_dt.date() != fecha_pred:
                filas_excluidas.append({
                    'indice': idx,
                    'razon': f'Fecha diferente: {fecha_dt.strftime("%d/%m/%Y")} (predominante: {fecha_pred.strftime("%d/%m/%Y")})',
                    'fila': df_raw.iloc[idx].tolist()
                })
            else:
                filas_con_fecha_correcta.append(idx)
                
        except Exception as e:
            filas_excluidas.append({
                'indice': idx,
                'razon': f'Error al procesar: {str(e)}',
                'fila': df_raw.iloc[idx].tolist()
            })
    
    # 🔥 MOSTRAR DETALLE DE FILAS EXCLUIDAS
    if filas_excluidas:
        # Resumir razones para que quede claro que NO son ingresos perdidos
        razones = {}
        for item in filas_excluidas:
            raz = str(item.get('razon', '?')).split(' (')[0]
            razones[raz] = razones.get(raz, 0) + 1
        detalle_razones = " · ".join([f"{k}: {v}" for k, v in sorted(razones.items())])
        n_filas_saldo = sum(1 for item in filas_excluidas if "capturado" in str(item.get("razon", "")))
        st.warning(f"⚠️ **{nombre_banco}: Se encontraron {len(filas_excluidas)} filas que no son movimientos válidos** ({detalle_razones}). Estas filas son encabezados, totales o filas sin fecha del archivo — NO son ingresos perdidos.")
        if n_filas_saldo > 0:
            st.success(f"✅ {nombre_banco}: {n_filas_saldo} filas de SALDO del estado de cuenta (Saldo Inicial/Final, Créditos/Débitos Total) fueron CAPTURADAS automáticamente — se usan en el resumen de saldos y en el exportable. No son movimientos perdidos.")
        
        # Mostrar tabla de filas excluidas
        datos_excluidos = []
        for item in filas_excluidas:
            # Crear una representación legible de la fila
            fila_str = " | ".join([str(v) for v in item['fila'][:8]])  # Mostrar primeras 8 columnas
            datos_excluidos.append({
                'Fila': item['indice'],
                'Razón': item['razon'],
                'Contenido': fila_str[:100] + "..." if len(fila_str) > 100 else fila_str
            })
        
        df_excluidos = pd.DataFrame(datos_excluidos)
        st.dataframe(df_excluidos, use_container_width=True)
        
        st.info(f"""
        **Detalle de exclusión:**
        - Total de filas en el archivo: {len(df_raw)}
        - Filas con fecha válida: {len(filas_con_fecha_correcta)}
        - Filas excluidas: {len(filas_excluidas)}
        - **Registros a procesar: {len(filas_con_fecha_correcta)}**
        """)
    else:
        st.success(f"✅ {nombre_banco}: Todas las {len(df_raw)} filas son movimientos válidos con fecha {fecha_pred.strftime('%d/%m/%Y')}")
    
    # Filtrar solo las filas con fecha correcta
    df_filtrado = df_raw.iloc[filas_con_fecha_correcta].copy()
    
    return df_filtrado, fecha_pred, dict_conteo, len(filas_con_fecha_correcta), len(filas_excluidas)

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

def obtener_tasa_bcv_con_origen(fecha=None, usar_api=False):
    """Obtiene la tasa de la fecha seleccionada + una descripción del origen de dónde salió.
    Con usar_api=True consulta la API BCV automáticamente según la fecha."""
    return _tasa_con_origen(fecha, usar_api)

def obtener_tasa_bcv(fecha=None, usar_api=False):
    """Obtiene la tasa de la fecha especificada de forma segura (automática vía API si usar_api=True)."""
    if fecha is None:
        fecha = date.today()
    tasa, _ = _tasa_con_origen(fecha, usar_api)
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
        # CORREGIDO: recorrido de abajo hacia arriba (última coincidencia = día más reciente)
        for r_idx in range(df_raw.shape[0] - 1, -1, -1):
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
        for r_idx in range(df_raw.shape[0] - 1, -1, -1):
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
        if isinstance(df_raw.columns, pd.MultiIndex):
            df_raw.columns = df_raw.columns.get_level_values(-1)
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

def extraer_resumen_bancamiga(df_raw):
    """Extrae datos resumen de filas excluidas (Saldo Inicial, Creditos Total, Debito Total, Saldo Final).
    CORREGIDO: soporta "Debito Total:" en singular (formato real del banco) y montos en celda siguiente."""
    resumen = {"saldo_inicial": None, "creditos_total": None, "debitos_total": None, "saldo_final": None}
    try:
        for r_idx in range(df_raw.shape[0]):
            for c_idx in range(df_raw.shape[1]):
                val_raw = str(df_raw.iloc[r_idx, c_idx]).strip()
                val_lower = val_raw.lower()
                if not val_lower or val_lower == "nan":
                    continue
                etiqueta = val_lower
                monto = None
                if ":" in val_raw:
                    partes = val_raw.split(":")
                    etiqueta = partes[0].lower()
                    monto = convertir_monto(partes[-1])
                es_inicial = "saldo inicial" in etiqueta
                es_creditos = any(t in etiqueta for t in [
                    "creditos total", "créditos total", "credito total", "crédito total"
                ])
                es_debitos = any(t in etiqueta for t in [
                    "debitos total", "débitos total", "debito total", "débito total"
                ])
                es_saldo = "saldo final" in etiqueta
                if not (es_inicial or es_creditos or es_debitos or es_saldo):
                    continue
                if monto is None:
                    for rc in range(c_idx + 1, df_raw.shape[1]):
                        monto = convertir_monto(df_raw.iloc[r_idx, rc])
                        if monto is not None:
                            break
                if monto is not None:
                    if es_inicial:
                        resumen["saldo_inicial"] = monto
                    elif es_creditos:
                        resumen["creditos_total"] = monto
                    elif es_debitos:
                        resumen["debitos_total"] = monto
                    else:
                        resumen["saldo_final"] = monto
    except:
        pass
    return resumen

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

def obtener_saldo_final_banplus(df_raw):
    """Extrae el saldo de Banplus buscando 'Saldo Total' al inicio del archivo.
    CORREGIDO: recorre de abajo hacia arriba para tomar el ÚLTIMO 'Saldo Total'
    (con varios días el saldo final es el del día más reciente)."""
    try:
        df_temp = df_raw.copy()
        for idx in range(min(15, len(df_temp)) - 1, -1, -1):
            for col_idx in range(df_temp.shape[1]):
                val_str = str(df_temp.iloc[idx, col_idx]).strip().lower()
                if "saldo total" in val_str:
                    for r_col in range(col_idx + 1, df_temp.shape[1]):
                        val_saldo = df_temp.iloc[idx, r_col]
                        val_clean = convertir_monto(val_saldo)
                        if val_clean is not None and val_clean > 0:
                            return val_clean
                    if df_temp.shape[1] > 5:
                        val_saldo = df_temp.iloc[idx, 5]
                        val_clean = convertir_monto(val_saldo)
                        if val_clean is not None:
                            return val_clean
    except Exception as e:
        st.warning(f"No se pudo extraer el saldo de Banplus: {e}")
    return 0.0

def parsear_fecha_flexible(val):
    """Parsea fechas en múltiples formatos: ISO con/sin hora, dd/mm/yyyy, etc.
    Evita el bug de BanPlus donde formato estricto %d/%m/%Y destruía todas las filas."""
    s = str(val).strip()
    if not s or s.lower() in ("nan", "nat"):
        return pd.NaT
    # Formato ISO: 2026-08-07 o 2026-08-07 00:00:00
    m = re.match(r'^(\d{4})-(\d{2})-(\d{2})', s)
    if m:
        try:
            return pd.to_datetime(f"{m.group(1)}-{m.group(2)}-{m.group(3)}")
        except Exception:
            return pd.NaT
    # Intentar formatos comunes (dayfirst para dd/mm/yyyy)
    for fmt in ["%d/%m/%Y", "%d/%m/%y", "%d-%m-%Y", "%d-%m-%y"]:
        try:
            return pd.to_datetime(s, format=fmt)
        except Exception:
            continue
    return pd.to_datetime(s, dayfirst=True, errors="coerce")

def extraer_resumen_banplus(df_raw):
    """Extrae datos resumen del archivo BanPlus (Saldo Total, Saldo Inicial)"""
    resumen = {"saldo_total": None, "saldo_inicial": None}
    try:
        for r_idx in range(df_raw.shape[0]):
            for c_idx in range(df_raw.shape[1]):
                val_raw = str(df_raw.iloc[r_idx, c_idx]).strip()
                val_lower = val_raw.lower()
                if not val_lower or val_lower == "nan":
                    continue
                # Buscar en la misma celda (ej: "Saldo Total: 198.883,66")
                if ":" in val_raw:
                    partes = val_raw.split(":")
                    monto = convertir_monto(partes[-1])
                else:
                    monto = None
                if "saldo total" in val_lower:
                    if monto is not None:
                        resumen["saldo_total"] = monto
                    else:
                        # Buscar en celdas a la derecha
                        for rc in range(c_idx + 1, df_raw.shape[1]):
                            monto = convertir_monto(df_raw.iloc[r_idx, rc])
                            if monto is not None:
                                resumen["saldo_total"] = monto
                                break
                elif "saldo inicial" in val_lower:
                    if monto is not None:
                        resumen["saldo_inicial"] = monto
                    else:
                        for rc in range(c_idx + 1, df_raw.shape[1]):
                            monto = convertir_monto(df_raw.iloc[r_idx, rc])
                            if monto is not None:
                                resumen["saldo_inicial"] = monto
                                break
    except:
        pass
    return resumen

def extraer_resumen_provincial(df_raw):
    """Extrae Saldo Inicial y Saldo Final por día del archivo de Provincial (texto TSV).
    Filas tipo: ['', '', '', '', 'Saldo Inicial: 09-08-2026', '6.208,43', ''].
    La fecha va dentro del texto de la etiqueta y el monto en la misma celda o la siguiente."""
    resumen = {"saldo_inicial": None, "saldo_final": None, "fechas": []}
    try:
        for r_idx in range(df_raw.shape[0]):
            for c_idx in range(df_raw.shape[1]):
                val_raw = str(df_raw.iloc[r_idx, c_idx]).strip()
                val_lower = val_raw.lower()
                if not val_lower or val_lower == "nan":
                    continue
                es_inicial = "saldo inicial" in val_lower
                es_final = "saldo final" in val_lower
                if not (es_inicial or es_final):
                    continue
                fecha = None
                m_fecha = re.search(r"\b(\d{2}[/\-]\d{2}[/\-]\d{4})\b", val_raw)
                if m_fecha:
                    fecha = m_fecha.group(1)
                monto = None
                if ":" in val_raw:
                    partes = val_raw.split(":")
                    monto = convertir_monto(partes[-1])
                    if monto is None:
                        texto_sin_fecha = re.sub(r"\b\d{2}[/\-]\d{2}[/\-]\d{4}\b", "", partes[-1])
                        numeros = re.findall(r"[\d\.\,]+", texto_sin_fecha)
                        if numeros:
                            monto = convertir_monto(numeros[-1])
                if monto is None:
                    for rc in range(c_idx + 1, df_raw.shape[1]):
                        monto = convertir_monto(df_raw.iloc[r_idx, rc])
                        if monto is not None:
                            break
                if monto is None:
                    continue
                if es_inicial:
                    resumen["saldo_inicial"] = monto
                else:
                    resumen["saldo_final"] = monto
                if fecha:
                    resumen["fechas"].append({
                        "tipo": "Inicial" if es_inicial else "Final",
                        "fecha": fecha,
                        "monto": monto
                    })
    except:
        pass
    return resumen

def extraer_resumen_venezuela(df_raw):
    """Extractor dedicado del Banco de Venezuela (BDV).

    Devuelve {"saldo_inicial", "saldo_final", "creditos_total", "debitos_total"}.

    Soporta dos formatos:
    A) Formato actual del estado de cuenta BDV (exportado con encabezado:
       Día|Referencia|Descripción|Fecha|Tipo de Movimiento|Crédito|Débito|Saldo|
       Saldo Inicial|Saldo Final|Total Crédito|Todal Débito), donde las columnas
       I..L repiten el resumen del período en cada fila de movimientos.
    B) Formato clásico: filas de texto tipo "SALDO INICIAL", "SALDO FINAL",
       "TOTAL CRÉDITO"/"TOTAL DÉBITO" con el monto en la celda de la derecha o abajo.
    """
    resumen = {"saldo_inicial": None, "saldo_final": None,
               "creditos_total": None, "debitos_total": None}

    # ---------- A) Formato con columnas de resumen por fila ----------
    try:
        enc = None
        for r_idx in range(min(25, df_raw.shape[0])):
            celdas = [str(v).upper().strip() for v in df_raw.iloc[r_idx].tolist() if pd.notna(v)]
            txt = " | ".join(celdas)
            if (re.search(r"(?<!\w)SALDO INICIAL(?!\w)", txt)
                    and re.search(r"(?<!\w)SALDO FINAL(?!\w)", txt)
                    and re.search(r"(?<!\w)(TODAL|TOTAL) (DÉBITO|DEBITO)(?!\w)", txt)):
                enc = r_idx
                break

        if enc is not None:
            h = df_raw.iloc[enc].tolist()

            def col_con_(*nombres):
                for i, c in enumerate(h):
                    ct = str(c).upper().strip()
                    for nb in nombres:
                        if re.search(r"(?<!\w)" + re.escape(nb) + r"(?!\w)", ct):
                            return i
                return None

            c_ini = col_con_("SALDO INICIAL")
            c_fin = col_con_("SALDO FINAL")
            c_cred = col_con_("TOTAL CRÉDITO", "TOTAL CREDITO", "CRÉDITO TOTAL", "CREDITO TOTAL")
            c_deb = col_con_("TODAL DÉBITO", "TODAL DEBITO", "TOTAL DÉBITO", "TOTAL DEBITO",
                             "DÉBITO TOTAL", "DEBITO TOTAL")

            primera_ini = None
            ultima = None
            for r_idx in range(enc + 1, df_raw.shape[0]):
                fila = df_raw.iloc[r_idx]

                def num_celda(c):
                    if c is None or c >= len(fila):
                        return None
                    return convertir_monto(fila[c])

                vi, vf = num_celda(c_ini), num_celda(c_fin)
                # En el formato BDV cada fila de movimiento repite el resumen del período
                if vi is None and vf is None:
                    continue
                if primera_ini is None and vi is not None and abs(vi) > 0:
                    primera_ini = vi
                if vf is not None and abs(vf) > 0:
                    ultima = (vi, vf,
                              num_celda(c_cred), num_celda(c_deb))

            if primera_ini is not None:
                resumen["saldo_inicial"] = primera_ini
            if ultima is not None:
                _, vf, vk, vd = ultima
                resumen["saldo_final"] = vf
                if vk is not None:
                    resumen["creditos_total"] = abs(vk)
                if vd is not None:
                    resumen["debitos_total"] = abs(vd)

            # Confirmación cruzada: el Saldo Final también debe aparecer como el
            # último saldo corrido (columna "Saldo") si esa columna existe.
            if resumen["saldo_final"] is None:
                saldos_corridos = []
                for r_idx in range(enc + 1, df_raw.shape[0]):
                    v = convertir_monto(df_raw.iloc[r_idx, 7]) if df_raw.shape[1] > 7 else None
                    if v is not None and abs(v) > 0:
                        saldos_corridos.append(v)
                if saldos_corridos:
                    resumen["saldo_final"] = saldos_corridos[-1]
    except Exception:
        pass

    # ---------- B) Formato clásico: filas de texto ----------
    if (resumen["saldo_inicial"] is None or resumen["saldo_final"] is None
            or resumen["creditos_total"] is None or resumen["debitos_total"] is None):
        try:
            def escanear_texto(terminos, permitir_saldo=False):
                for r_idx in range(df_raw.shape[0]):
                    for c_idx in range(df_raw.shape[1]):
                        val = str(df_raw.iloc[r_idx, c_idx]).upper()
                        if not any(t in val for t in terminos):
                            continue
                        # a) Número embebido en la misma celda ("SALDO FINAL Bs 1.234,56")
                        texto_sin_fechas = re.sub(r"\b\d{2}[/\-]\d{2}[/\-]\d{4}\b", "", val)
                        partes = re.findall(r"-?[\d\.\,]+", texto_sin_fechas)
                        for p in reversed(partes):
                            n = convertir_monto(p)
                            if n is not None and (permitir_saldo or abs(n) > 0):
                                return n
                        # b) Cualquier número en la MISMA FILA hacia la derecha del texto
                        # (el formato BDV suele poner el monto varias celdas más a la derecha)
                        for j_idx in range(c_idx + 1, df_raw.shape[1]):
                            n = convertir_monto(df_raw.iloc[r_idx, j_idx])
                            if n is not None and (permitir_saldo or abs(n) > 0):
                                return n
                        # c) Celda de abajo (misma columna)
                        if r_idx + 1 < df_raw.shape[0]:
                            n = convertir_monto(df_raw.iloc[r_idx + 1, c_idx])
                            if n is not None and (permitir_saldo or abs(n) > 0):
                                return n
                return None

            if resumen["saldo_inicial"] is None:
                resumen["saldo_inicial"] = escanear_texto(
                    ["SALDO INICIAL", "SALDO ANTERIOR"], True)
            if resumen["saldo_final"] is None:
                resumen["saldo_final"] = escanear_texto(
                    ["SALDO FINAL", "SALDO DISPONIBLE", "SALDO ACTUAL", "SALDO A LA FECHA"], True)
            if resumen["creditos_total"] is None:
                resumen["creditos_total"] = escanear_texto(
                    ["TOTAL CREDITO", "TOTAL CRÉDITO", "CREDITO TOTAL", "CRÉDITO TOTAL",
                     "CREDITOS:", "CRÉDITOS:", "TOTAL DE CREDITOS", "TOTAL DE CRÉDITOS"])
            if resumen["debitos_total"] is None:
                resumen["debitos_total"] = escanear_texto(
                    ["TODAL DEBITO", "TODAL DÉBITO", "TOTAL DEBITO", "TOTAL DÉBITO",
                     "DEBITO TOTAL", "DÉBITO TOTAL", "DEBITOS:", "DÉBITOS:",
                     "TOTAL DE DEBITOS", "TOTAL DE DÉBITOS"])
        except Exception:
            pass

    # Normalizar: el débito se expresa positivo y None se deja como 0.0 para la UI
    if resumen["creditos_total"] is not None:
        resumen["creditos_total"] = abs(float(resumen["creditos_total"]))
    if resumen["debitos_total"] is not None:
        resumen["debitos_total"] = abs(float(resumen["debitos_total"]))
    return resumen


def extraer_resumen_banco(df_raw, banco):
    """Despachador: extrae el resumen (saldo inicial/final, créditos/débitos total) de un archivo de banco.
    Los bancos sin extractor específico usan los extractores genéricos existentes."""
    try:
        if banco == "venezuela":
            return extraer_resumen_venezuela(df_raw)
        if banco == "bancamiga":
            r = extraer_resumen_bancamiga(df_raw)
            return {"saldo_inicial": r.get("saldo_inicial"), "saldo_final": r.get("saldo_final"),
                    "creditos_total": r.get("creditos_total"), "debitos_total": r.get("debitos_total")}
        if banco == "banplus":
            r = extraer_resumen_banplus(df_raw)
            return {"saldo_inicial": r.get("saldo_inicial"), "saldo_final": r.get("saldo_total"),
                    "creditos_total": None, "debitos_total": None}
        if banco == "provincial":
            r = extraer_resumen_provincial(df_raw)
            return {"saldo_inicial": r.get("saldo_inicial"), "saldo_final": r.get("saldo_final"),
                    "creditos_total": None, "debitos_total": None}
    except Exception:
        pass
    try:
        return {"saldo_inicial": buscar_saldo_inicial(df_raw),
                "saldo_final": obtener_saldo_banco(df_raw, banco),
                "creditos_total": None, "debitos_total": None}
    except Exception:
        return {"saldo_inicial": None, "saldo_final": None, "creditos_total": None, "debitos_total": None}

def obtener_saldo_final_banco_activo(df_raw):
    """
    Extrae el saldo final de Banco Activo.
    El saldo aparece en la columna SALDO (columna 7) en la última fila de movimientos.
    """
    try:
        # Saltar la primera fila (Número de Cuenta) y empezar desde la fila 1
        df_datos = df_raw.iloc[1:].copy().reset_index(drop=True)
        
        # Asignar nombres de columnas
        if df_datos.shape[1] >= 7:
            df_datos.columns = ["FECHA", "FECHA_VALOR", "DESCRIPCION", "NRO", "DEBITO", "CREDITO", "SALDO"]
            
            # Buscar el último valor no nulo en la columna SALDO
            saldos = df_datos["SALDO"].dropna()
            for val in reversed(saldos.values):
                val_clean = convertir_monto(val)
                if val_clean is not None and val_clean > 0:
                    return val_clean
        
        # Fallback: buscar en texto
        return buscar_saldo_en_texto(df_raw)
        
    except Exception as e:
        st.warning(f"No se pudo extraer el saldo de Banco Activo: {e}")
        return 0.0

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
    elif banco == "banplus":
        return obtener_saldo_final_banplus(df_raw) or buscar_saldo_en_texto(df_raw) or obtener_saldo_final_columna_derecha(df_raw)
    elif banco == "activo":
        return obtener_saldo_final_banco_activo(df_raw) or buscar_saldo_en_texto(df_raw)
    elif banco == "venezuela":
        # Extractor dedicado BDV (lee el Saldo Final de la columna del reporte o la
        # última fila de saldo corrido); respaldo: escáner genérico de texto
        r_vz = extraer_resumen_venezuela(df_raw)
        if r_vz.get("saldo_final"):
            return r_vz["saldo_final"]
        return buscar_saldo_en_texto(df_raw) or obtener_saldo_final_columna_derecha(df_raw)
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
                try:
                    df_html = leer_tabla_html(archivo)
                    if not df_html.empty:
                        return df_html
                except Exception:
                    pass
                archivo.seek(0)
                contenido = archivo.read()
                try:
                    contenido = contenido.decode("utf-8")
                except UnicodeDecodeError:
                    archivo.seek(0)
                    contenido = archivo.read().decode("latin-1")
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
            except Exception:
                archivo.seek(0)
                try:
                    df_html = leer_tabla_html(archivo)
                    if not df_html.empty:
                        return df_html
                except Exception:
                    pass
                archivo.seek(0)
                return pd.read_excel(archivo, sheet_name=0, header=0, engine='xlrd')
        else:
            return pd.read_excel(archivo, sheet_name=0, header=0, engine='openpyxl')
    except Exception as e:
        try:
            return pd.read_excel(archivo, sheet_name=0, header=None, engine='openpyxl')
        except:
            st.error(f"No se pudo leer el archivo. Error: {str(e)}")
            st.stop()

def leer_tabla_html(archivo):
    """Lee archivos que son tablas HTML (extensión .xls engañosa) usando SOLO librería estándar.
    No depende de lxml/html5lib (que no están instalados en el entorno).
    Devuelve un DataFrame con la fila de encabezados <th> como primera fila."""
    import html as _html
    archivo.seek(0)
    contenido = archivo.read()
    try:
        contenido = contenido.decode("utf-8")
    except UnicodeDecodeError:
        archivo.seek(0)
        contenido = archivo.read().decode("latin-1")
    
    filas = []
    for tr in re.findall(r"<tr[^>]*>(.*?)</tr>", contenido, re.IGNORECASE | re.DOTALL):
        celdas = re.findall(r"<t[dh][^>]*>(.*?)</t[dh]>", tr, re.IGNORECASE | re.DOTALL)
        if not celdas:
            continue
        fila = []
        for celda in celdas:
            texto = _html.unescape(celda)
            texto = re.sub(r"<[^>]+>", "", texto)
            texto = texto.replace("\xa0", "").replace("\r", " ").replace("\n", " ").replace("\t", " ")
            fila.append(texto.strip())
        filas.append(fila)
    
    if not filas:
        return pd.DataFrame()
    max_cols = max(len(f) for f in filas)
    filas_pad = [f + [""] * (max_cols - len(f)) for f in filas]
    return pd.DataFrame(filas_pad)

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
            texto = " ".join(df_temp.fillna("").astype(str).values.flatten()).upper()
            # 🔥 "0171" debe ser un número de cuenta de Banco Activo (20 dígitos aislados),
            # no cualquier coincidencia dentro de referencias de otros bancos
            es_cuenta_activo = bool(re.search(r'(?<![\d])0171\d{16}(?!\d)', texto))
            # Buscar específicamente Banco Activo
            if "BANCO ACTIVO" in texto or es_cuenta_activo or "NUMERO DE CUENTA" in texto:
                # Verificar si tiene el formato de Banco Activo (fechas en columna 0, conceptos en columna 2)
                if df_temp.shape[1] >= 7:
                    # Intentar detectar el patrón de Banco Activo
                    for i in range(min(10, len(df_temp))):
                        fila = df_temp.iloc[i]
                        if pd.notna(fila[0]) and pd.notna(fila[2]) and pd.notna(fila[6]):
                            # Verificar si la columna 0 tiene formato de fecha
                            fecha_str = str(fila[0]).strip()
                            if re.match(r'\d{2}/\d{2}/\d{4}', fecha_str):
                                archivo.seek(pos)
                                return "activo"
            elif "BANCAMIGA" in texto or "BANCAMIGA BANCO UNIVERSAL" in texto or "BANCA AMIGA" in texto or "AMIGA" in texto:
                archivo.seek(pos)
                return "bancamiga"
            elif "BANESCO" in texto:
                archivo.seek(pos)
                return "banesco"
            elif "MERCANTIL" in texto or "105 VEB" in texto:
                archivo.seek(pos)
                return "mercantil"
            elif (
                "BANCO PROVINCIAL" in texto
                or "BBVA" in texto
                or "CÓDIGO DE OPERACIÓN" in texto
                or "CODIGO DE OPERACION" in texto
                or "PRIMERA ORDENACIÓN" in texto
                or "F. OPERACIÓN" in texto
                or "CUENTA ACTUAL" in texto
            ):
                archivo.seek(pos)
                return "provincial"
            elif "BANCO DE VENEZUELA" in texto or "BDV" in texto:
                archivo.seek(pos)
                return "venezuela"
            elif "BNC" in texto:
                archivo.seek(pos)
                return "bnc"
            elif "TESORO" in texto or "BANCO DEL TESORO" in texto:
                archivo.seek(pos)
                return "tesoro"
            archivo.seek(pos)
            return None
        except Exception:
            pass
        archivo.seek(pos)
        return None
    except Exception:
        return None

def detectar_banco_por_nombre(nombre_archivo):
    """Detecta el banco por el nombre del archivo (fallback)"""
    nombre = nombre_archivo.upper()
    if "ACTIVO" in nombre:
        return "activo"
    elif "MERCANTIL" in nombre:
        return "mercantil"
    elif "TESORO" in nombre or "TESORERIA" in nombre or "TES" in nombre:
        return "tesoro"
    elif "BANCAMIGA" in nombre or "BANCAAMIGA" in nombre or "AMIGA" in nombre:
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
    elif "PROVINCIAL" in nombre or "BBVA" in nombre:
        return "provincial"
    elif "BNC" in nombre:
        return "bnc"
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
    """
    Detecta si un movimiento es una comisión bancaria.
    
    REGLAS:
    - Si tiene proveedor asociado → NO es comisión bancaria
    - Si es pago a personal (nómina, comisiones de ventas) → NO es comisión bancaria
    - Si es "COMISION PAGO A PROVEEDORES" → NO es comisión bancaria
    - Si contiene "FACTURA" → NO es comisión bancaria (es pago a proveedor)
    - Solo son comisiones bancarias: cargos del banco (ITF, mantenimiento, etc.)
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
    
    # 🔥 REGLA 4: "COMISION PAGO A PROVEEDORES" NO es comisión bancaria
    if "COMISION PAGO A PROVEEDORES" in texto_upper:
        return False
    
    # 🔥 REGLA 5: Si contiene "PAGO" y "PROVEEDOR" NO es comisión bancaria
    if "PAGO" in texto_upper and "PROVEEDOR" in texto_upper:
        return False
    
    # 🔥 REGLA 6: Si contiene "FACTURA" NO es comisión bancaria (es pago a proveedor)
    if "FACTURA" in texto_upper:
        return False
    
    # 🔥 REGLA 7: Si contiene "PAGO" + nombre de empresa/persona NO es comisión bancaria
    if "PAGO" in texto_upper and any(x in texto_upper for x in ["TRANSPORTE", "CHOFER", "VIATICOS", "MOCASA", "PIOTELLO"]):
        return False
    
    # 🔥 REGLA 8: SOLO SON COMISIONES BANCARIAS si coinciden con estas palabras
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
        "comision pago movil comercial",
        "comision x pago de nominas mb",
        "com pago otras ctas",
        "com pago otr bcos",
        "comis",
        "comis. cr.i",
        "sms",
        "servicio sms",
        "servicio sms plus",
        "sms plus",
        "domiciliacion j412438905",
        "distribuidora global",
        "emision edo",
        "retencion de impuesto",
        "com. trf",
        "com trf",
        "com transf",
        "com.serv",
        "emision de estado",
        "below minimum balance charges",
        "stament service",
        "statement service",
        "descuento de tarjeta",
        "descuento tarjeta",
        "desc. tarjeta",
        "desc tarjeta",
        "descuento tarjeta credito"
    ]
    
    for patron in palabras_comision_bancaria:
        if patron in texto:
            return True
    
    # 🔥 REGLA 9: Si contiene "COMISION" pero también "PROVEEDOR" o "FACTURA" NO es comisión bancaria
    if "COMISION" in texto_upper or "COMISIÓN" in texto_upper:
        if any(x in texto_upper for x in ["PROVEEDOR", "FACTURA", "PAGO A"]):
            return False
    
    # Si contiene "comision" pero no coincide con las reglas anteriores, NO es comisión bancaria
    if "comision" in texto or "comisión" in texto:
        return False
    
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
        # 🔥 APLICAR FILTRO DE FECHA PREDOMINANTE
        df_filtrado, fecha_pred, dict_conteo, total_filas, registros_excluidos = filtrar_por_fecha_predominante(
            df, columna_fecha_idx=0, nombre_banco="Banesco"
        )
        
        # Si no se pudo filtrar o está vacío, usar el df original
        if df_filtrado is None or df_filtrado.empty:
            df_filtrado = df
            st.warning("⚠️ Banesco: No se pudo filtrar por fecha predominante, se procesará todo el archivo.")
        
        # Guardar información de fechas en session_state
        if fecha_pred:
            _acumular_info_fechas("Banesco", fecha_pred.strftime('%d/%m/%Y'), len(df_filtrado), total_filas, registros_excluidos, dict_conteo)
        df_filtrado.columns = ["FECHA", "REFERENCIA", "DESCRIPCION", "MONTO_RAW", "BALANCE"]
        df_filtrado.columns = [str(c).strip().upper() for c in df_filtrado.columns]
        rename_map = {}
        for col in df_filtrado.columns:
            c = str(col).lower()
            if "fecha" in c: rename_map[col] = "FECHA"
            elif "referencia" in c: rename_map[col] = "REFERENCIA"
            elif "descrip" in c: rename_map[col] = "DESCRIPCION"
            elif "monto" in c: rename_map[col] = "MONTO_RAW"
        df_filtrado = df_filtrado.rename(columns=rename_map)
        for col in ["FECHA", "REFERENCIA", "DESCRIPCION", "MONTO_RAW"]:
            if col not in df_filtrado.columns:
                st.error(f"No existe columna: {col}")
                return pd.DataFrame()
        # Convertir fechas de manera robusta
        def parse_banesco_date(val):
            val_str = str(val).strip()
            if not val_str or val_str == "nan":
                return pd.NaT
            if len(val_str) >= 5 and val_str[:4].isdigit() and val_str[4] in ('/', '-'):
                return pd.to_datetime(val_str, dayfirst=False, errors="coerce")
            return pd.to_datetime(val_str, dayfirst=True, errors="coerce")

        df_filtrado["FECHA"] = df_filtrado["FECHA"].apply(parse_banesco_date)
        df_filtrado = df_filtrado[df_filtrado["FECHA"].notna()]
        
        # 🔥 PARSING DE MONTO ROBUSTO: soporta positivos sin "+", punto decimal
        # (345570.17), coma decimal venezolana (345.570,17) y enteros (316000)
        def limpiar_monto_banesco(valor):
            if valor is None or (isinstance(valor, float) and pd.isna(valor)):
                return None
            if isinstance(valor, (int, float, np.integer, np.floating)):
                return float(valor)
            val_str = str(valor).strip().replace(" ", "")
            if not val_str or val_str.lower() == "nan":
                return None
            es_negativo = val_str.startswith("-")
            val_str = val_str.lstrip("+-")
            if "," in val_str and "." in val_str:
                if val_str.rfind(",") > val_str.rfind("."):
                    val_str = val_str.replace(".", "").replace(",", ".")
                else:
                    val_str = val_str.replace(",", "")
            elif "," in val_str:
                val_str = val_str.replace(",", ".")
            try:
                numero = float(val_str)
            except ValueError:
                return None
            return -abs(numero) if es_negativo else numero
        
        df_filtrado["MONTO_LIMPIO"] = df_filtrado["MONTO_RAW"].apply(limpiar_monto_banesco)
        df_filtrado["TIPO"] = df_filtrado["MONTO_LIMPIO"].apply(
            lambda x: "NC" if x is not None and x > 0 else "ND" if x is not None else "ND"
        )
        df_filtrado["MONTO"] = df_filtrado["MONTO_LIMPIO"].abs()
        df_filtrado = df_filtrado[df_filtrado["MONTO"].notna()]
        df_filtrado = df_filtrado[df_filtrado["MONTO"] > 0]
        df_filtrado = df_filtrado[["FECHA", "REFERENCIA", "DESCRIPCION", "TIPO", "MONTO"]]
        st.success(f"Banesco OK: {len(df_filtrado)} movimientos")
        return df_filtrado
    except Exception as e:
        st.error(f"Error Banesco: {str(e)}")
        return pd.DataFrame()

def procesar_provincial(df):
    st.info("🔍 Procesando archivo de Provincial...")
    try:
        # 🔥 EXTRAER RESUMEN DEL ARCHIVO (Saldo Inicial/Final por día) — se muestra y se usa en el exportable
        resumen_archivo = extraer_resumen_provincial(df)
        if resumen_archivo.get("fechas") or resumen_archivo.get("saldo_inicial") or resumen_archivo.get("saldo_final"):
            detalle_resumen = " · ".join(
                f"{det['tipo']} {det['fecha']}: {formato_venezolano(det['monto'])}"
                for det in resumen_archivo.get("fechas", [])
            )
            if not detalle_resumen:
                detalle_resumen = f"Saldo Inicial: {formato_venezolano(resumen_archivo.get('saldo_inicial'))}"
                if resumen_archivo.get("saldo_final"):
                    detalle_resumen += f" · Saldo Final: {formato_venezolano(resumen_archivo.get('saldo_final'))}"
            st.info(f"📊 **Provincial - Resumen del archivo:** {detalle_resumen}")
        # 🔥 1. BUSCAR EL ENCABEZADO EN EL DF ORIGINAL (antes de filtrar por fecha)
        encabezado_idx = None
        for i in range(min(30, len(df))):
            fila = df.iloc[i]
            fila_str = [str(val) for val in fila.tolist()]
            texto_fila = " ".join(fila_str).upper()
            if "CONCEPTO" in texto_fila and "IMPORTE" in texto_fila:
                encabezado_idx = i
                break
        if encabezado_idx is None:
            st.error("❌ No se encontró la fila de encabezados en el archivo Provincial.")
            return pd.DataFrame()
        fila_encabezado = df.iloc[[encabezado_idx]].copy()
        df_datos = df.iloc[encabezado_idx + 1:].copy()
        
        # 🔥 2. APLICAR FILTRO DE FECHA PREDOMINANTE SOLO A LAS FILAS DE DATOS (post-encabezado)
        df_filtrado, fecha_pred, dict_conteo, total_filas, registros_excluidos = filtrar_por_fecha_predominante(
            df_datos, columna_fecha_idx=0, nombre_banco="Provincial"
        )
        
        if df_filtrado is None or df_filtrado.empty:
            df_filtrado = df_datos
            st.warning("⚠️ Provincial: No se pudo filtrar por fecha predominante, se procesará todo el archivo.")
        
        if fecha_pred:
            _acumular_info_fechas("Provincial", fecha_pred.strftime('%d/%m/%Y'), len(df_filtrado), total_filas, registros_excluidos, dict_conteo)
        # 🔥 3. REENSAMBLAR: encabezado + datos filtrados
        df_filtrado = pd.concat([fila_encabezado, df_filtrado], ignore_index=True)
        encabezado_idx = 0
        headers = df_filtrado.iloc[encabezado_idx].astype(str).str.strip().tolist()
        rename_map = {}
        for col in headers:
            col_clean = str(col).strip().upper()
            if "OPERAC" in col_clean or "FECHA" in col_clean: rename_map[col] = "FECHA"
            elif "F. VALOR" in col_clean: rename_map[col] = "FECHA_VALOR"
            elif "CÓDIGO" in col_clean or "CODIGO" in col_clean: rename_map[col] = "CODIGO"
            elif "Nº. DOC" in col_clean or "NRO DOC" in col_clean or "DOC" in col_clean: rename_map[col] = "REFERENCIA"
            elif "CONCEPTO" in col_clean: rename_map[col] = "DESCRIPCION"
            elif "IMPORTE" in col_clean: rename_map[col] = "MONTO"
        df_filtrado.columns = headers
        df_filtrado = df_filtrado.iloc[encabezado_idx + 1:].reset_index(drop=True)
        df_filtrado = df_filtrado.rename(columns=rename_map)
        if "FECHA" in df_filtrado.columns:
            df_filtrado["FECHA"] = df_filtrado["FECHA"].astype(str).str.strip()
            df_filtrado = df_filtrado[~df_filtrado["FECHA"].str.contains("FECHA|SALDO|Período", case=False, na=False)]
            # Convertir fechas de manera robusta (admite dd/mm/yyyy y datetime de Excel)
            df_filtrado["FECHA_DT"] = pd.to_datetime(df_filtrado["FECHA"], dayfirst=True, errors="coerce")
            mask = df_filtrado["FECHA_DT"].isna()
            if mask.any():
                df_filtrado.loc[mask, "FECHA_DT"] = pd.to_datetime(df_filtrado.loc[mask, "FECHA"].astype(str).str.strip(), dayfirst=True, errors="coerce")
            df_filtrado = df_filtrado[df_filtrado["FECHA_DT"].notna()]
            df_filtrado["FECHA"] = df_filtrado["FECHA_DT"].dt.strftime("%d/%m/%Y")
            df_filtrado = df_filtrado.drop(columns=["FECHA_DT"])
        else:
            return pd.DataFrame()
        if "MONTO" in df_filtrado.columns:
            df_filtrado["MONTO"] = df_filtrado["MONTO"].astype(str).str.replace(" ", "", regex=False).str.replace(".", "", regex=False).str.replace(",", ".", regex=False).str.replace("'", "", regex=False)
            df_filtrado["MONTO"] = pd.to_numeric(df_filtrado["MONTO"], errors="coerce")
            df_filtrado = df_filtrado[df_filtrado["MONTO"].notna()]
            df_filtrado["TIPO"] = df_filtrado["MONTO"].apply(lambda x: "NC" if x > 0 else "ND" if x < 0 else "")
            df_filtrado["MONTO"] = df_filtrado["MONTO"].abs()
            df_filtrado = df_filtrado[df_filtrado["MONTO"] > 0]
        else:
            return pd.DataFrame()
        if "REFERENCIA" not in df_filtrado.columns: df_filtrado["REFERENCIA"] = ""
        else: df_filtrado["REFERENCIA"] = df_filtrado["REFERENCIA"].astype(str).str.strip().str.replace("'", "", regex=False)
        if "DESCRIPCION" not in df_filtrado.columns: df_filtrado["DESCRIPCION"] = ""
        else: df_filtrado["DESCRIPCION"] = df_filtrado["DESCRIPCION"].astype(str).str.strip()
        df_filtrado["ES_COMISION"] = df_filtrado["DESCRIPCION"].str.contains("COMIS", case=False, na=False)
        df_resultado = df_filtrado[["FECHA", "REFERENCIA", "DESCRIPCION", "TIPO", "MONTO", "ES_COMISION"]].copy()
        st.success(f"✅ Provincial OK: {len(df_resultado)} movimientos")
        return df_resultado
    except Exception as e:
        st.error(f"❌ Error procesando Provincial: {str(e)}")
        return pd.DataFrame()

def procesar_bnc(df):
    st.info("Procesando archivo BNC...")
    try:
        # 🔥 1. BUSCAR EL ENCABEZADO EN EL DF ORIGINAL (antes de filtrar por fecha)
        encabezado = None
        for i in range(min(30, len(df))):
            fila = df.iloc[i].fillna("").astype(str)
            texto = " ".join(fila.tolist()).lower()
            if "fecha" in texto and ("descripcion" in texto or "descripción" in texto):
                encabezado = i
                break
        if encabezado is not None:
            fila_encabezado = df.iloc[[encabezado]].copy()
            df_datos = df.iloc[encabezado + 1:].copy()
        else:
            df_datos = df.copy()
            st.warning("⚠️ BNC: No se encontró la fila de encabezados, se procesará todo el archivo.")
        
        # 🔥 2. APLICAR FILTRO DE FECHA PREDOMINANTE SOLO A LAS FILAS DE DATOS
        df_filtrado, fecha_pred, dict_conteo, total_filas, registros_excluidos = filtrar_por_fecha_predominante(
            df_datos, columna_fecha_idx=0, nombre_banco="BNC"
        )
        
        if df_filtrado is None or df_filtrado.empty:
            df_filtrado = df_datos
            st.warning("⚠️ BNC: No se pudo filtrar por fecha predominante, se procesará todo el archivo.")
        
        if fecha_pred:
            _acumular_info_fechas("BNC", fecha_pred.strftime('%d/%m/%Y'), len(df_filtrado), total_filas, registros_excluidos, dict_conteo)
        # 🔥 3. REENSAMBLAR: encabezado + datos filtrados
        if encabezado is not None:
            df_filtrado = pd.concat([fila_encabezado, df_filtrado], ignore_index=True)
            encabezado = 0
        
        if encabezado is not None:
            headers = []
        for idx, col in enumerate(df_filtrado.iloc[encabezado]):
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
        df_filtrado.columns = headers_unicos
        df_filtrado = df_filtrado.iloc[encabezado + 1:].reset_index(drop=True)
        rename_map = {}
        for col in df_filtrado.columns:
            col_str = str(col).strip().lower()
            if "fecha" in col_str: rename_map[col] = "FECHA"
            elif "descripcion" in col_str or "descripción" in col_str or "concepto" in col_str: rename_map[col] = "DESCRIPCION"
            elif "referencia" in col_str: rename_map[col] = "REFERENCIA"
            elif "credito" in col_str or "haber" in col_str: rename_map[col] = "CREDITO"
            elif "debito" in col_str or "debe" in col_str: rename_map[col] = "DEBITO"
        df_filtrado = df_filtrado.rename(columns=rename_map)
        if "FECHA" in df_filtrado.columns:
            df_filtrado["FECHA"] = df_filtrado["FECHA"].apply(parsear_fecha_flexible)
            df_filtrado = df_filtrado[df_filtrado["FECHA"].notna()]
        df_filtrado["CREDITO"] = pd.to_numeric(df_filtrado.get("CREDITO", 0), errors="coerce").fillna(0) if "CREDITO" in df_filtrado.columns else 0
        df_filtrado["DEBITO"] = pd.to_numeric(df_filtrado.get("DEBITO", 0), errors="coerce").fillna(0) if "DEBITO" in df_filtrado.columns else 0
        df_filtrado["MONTO"] = df_filtrado["CREDITO"] - df_filtrado["DEBITO"]
        df_filtrado["TIPO"] = df_filtrado["MONTO"].apply(lambda x: "NC" if x > 0 else "ND")
        df_filtrado["MONTO"] = df_filtrado["MONTO"].abs()
        df_filtrado = df_filtrado[df_filtrado["MONTO"] != 0]
        st.success(f"Registros BNC OK: {len(df_filtrado)}")
        return df_filtrado
    except Exception as e:
        st.error(f"Error BNC: {e}")
        return pd.DataFrame()

def procesar_tesoro(df):
    st.info("Procesando Banco del Tesoro...")
    try:
        # 🔥 1. BUSCAR EL ENCABEZADO EN EL DF ORIGINAL (antes de filtrar por fecha)
        encabezado = None
        for i in range(min(20, len(df))):
            fila = df.iloc[i].astype(str)
            texto = " ".join(map(str, fila.tolist())).lower()
            if "fecha" in texto and "referencia" in texto and "concepto" in texto:
                encabezado = i
                break
        if encabezado is None:
            return pd.DataFrame()
        fila_encabezado = df.iloc[[encabezado]].copy()
        df_datos = df.iloc[encabezado + 1:].copy()
        
        # 🔥 2. APLICAR FILTRO DE FECHA PREDOMINANTE SOLO A LAS FILAS DE DATOS (fechas en columna 1)
        df_filtrado, fecha_pred, dict_conteo, total_filas, registros_excluidos = filtrar_por_fecha_predominante(
            df_datos, columna_fecha_idx=1, nombre_banco="Tesoro"
        )
        
        if df_filtrado is None or df_filtrado.empty:
            df_filtrado = df_datos
            st.warning("⚠️ Tesoro: No se pudo filtrar por fecha predominante, se procesará todo el archivo.")
        
        if fecha_pred:
            _acumular_info_fechas("Tesoro", fecha_pred.strftime('%d/%m/%Y'), len(df_filtrado), total_filas, registros_excluidos, dict_conteo)
        # 🔥 3. REENSAMBLAR: encabezado + datos filtrados
        df_filtrado = pd.concat([fila_encabezado, df_filtrado], ignore_index=True)
        encabezado = 0
        df_filtrado.columns = df_filtrado.iloc[encabezado]
        df_filtrado = df_filtrado.iloc[encabezado + 1:].reset_index(drop=True)
        df_filtrado.columns = [str(c).strip() for c in df_filtrado.columns]
        rename_map = {}
        for col in df_filtrado.columns:
            c = str(col).strip().lower()
            if "fecha" in c: rename_map[col] = "FECHA"
            elif "referencia" in c: rename_map[col] = "REFERENCIA"
            elif "concepto" in c: rename_map[col] = "DESCRIPCION"
            elif "débito" in c or "debito" in c: rename_map[col] = "DEBITO"
            elif "crédito" in c or "credito" in c: rename_map[col] = "CREDITO"
            elif "código" in c or "codigo" in c: rename_map[col] = "TIPO"
        df_filtrado = df_filtrado.rename(columns=rename_map)
        if "FECHA" not in df_filtrado.columns: return pd.DataFrame()
        df_filtrado["FECHA"] = df_filtrado["FECHA"].apply(parsear_fecha_flexible)
        df_filtrado = df_filtrado[df_filtrado["FECHA"].notna()]
        def limpiar_numero(valor):
            valor = str(valor).replace(".", "").replace(",", ".")
            try: return float(valor)
            except: return 0
        df_filtrado["CREDITO"] = df_filtrado.get("CREDITO", 0).apply(limpiar_numero) if "CREDITO" in df_filtrado.columns else 0
        df_filtrado["DEBITO"] = df_filtrado.get("DEBITO", 0).apply(limpiar_numero) if "DEBITO" in df_filtrado.columns else 0
        df_filtrado["MONTO"] = df_filtrado["CREDITO"] - df_filtrado["DEBITO"]
        df_filtrado["TIPO"] = df_filtrado["MONTO"].apply(lambda x: "NC" if x > 0 else "ND")
        df_filtrado["MONTO"] = df_filtrado["MONTO"].abs()
        df_filtrado = df_filtrado[df_filtrado["MONTO"] > 0]
        df_filtrado = df_filtrado[["FECHA", "REFERENCIA", "DESCRIPCION", "TIPO", "MONTO"]]
        st.success(f"Tesoro OK: {len(df_filtrado)} registros")
        return df_filtrado
    except Exception as e:
        st.error(f"Error Tesoro: {str(e)}")
        return pd.DataFrame()

def procesar_bancamiga(df):
    st.info("🔍 Procesando archivo de Bancamiga...")
    try:
        if isinstance(df.columns, pd.MultiIndex):
            df.columns = df.columns.get_level_values(-1)
        
        # 🔥 EXTRAER RESUMEN DEL ARCHIVO (Créditos Total, Débitos Total, Saldo Final)
        resumen_archivo = extraer_resumen_bancamiga(df)
        if any(v is not None for v in resumen_archivo.values()):
            creditos = formato_venezolano(resumen_archivo['creditos_total']) if resumen_archivo['creditos_total'] is not None else "N/A"
            debitos = formato_venezolano(resumen_archivo['debitos_total']) if resumen_archivo['debitos_total'] is not None else "N/A"
            saldo = formato_venezolano(resumen_archivo['saldo_final']) if resumen_archivo['saldo_final'] is not None else "N/A"
            st.info(f"""
            📊 **Bancamiga - Resumen del archivo:**
            - Créditos Total (Ingresos): {creditos} BS
            - Débitos Total (Egresos): {debitos} BS
            - Saldo Final: {saldo} BS
            """)
        
        columnas = [str(c).strip().upper() for c in df.columns]
        tiene_encabezados = "FECHA" in columnas and "REFERENCIA" in columnas
        
        # 🔥 1. BUSCAR EL ENCABEZADO EN EL DF ORIGINAL (antes de filtrar por fecha)
        encabezado_idx = None
        if not tiene_encabezados:
            for i in range(min(30, len(df))):
                fila = df.iloc[i]
                fila_str = [str(val) for val in fila.tolist()]
                texto_fila = " ".join(fila_str).upper()
                if "NRO" in texto_fila and "FECHA" in texto_fila and "REFERENCIA" in texto_fila:
                    encabezado_idx = i
                    break
        
        if tiene_encabezados:
            df_datos = df.copy()
        elif encabezado_idx is not None:
            fila_encabezado = df.iloc[[encabezado_idx]].copy()
            df_datos = df.iloc[encabezado_idx + 1:].copy()
        else:
            df_datos = df.copy()
            st.warning("⚠️ Bancamiga: No se encontró la fila de encabezados, se procesará todo el archivo.")
        
        # 🔥 2. APLICAR FILTRO DE FECHA PREDOMINANTE SOLO A LAS FILAS DE DATOS
        df_filtrado, fecha_pred, dict_conteo, total_filas, registros_excluidos = filtrar_por_fecha_predominante(
            df_datos, columna_fecha_idx=1, nombre_banco="Bancamiga"
        )
        
        if df_filtrado is None or df_filtrado.empty:
            df_filtrado = df_datos
            st.warning("⚠️ Bancamiga: No se pudo filtrar por fecha predominante, se procesará todo el archivo.")
        
        if fecha_pred:
            _acumular_info_fechas("Bancamiga", fecha_pred.strftime('%d/%m/%Y'), len(df_filtrado), total_filas, registros_excluidos, dict_conteo)
        # 🔥 3. REENSAMBLAR: encabezado + datos filtrados
        if not tiene_encabezados and encabezado_idx is not None:
            df_filtrado = pd.concat([fila_encabezado, df_filtrado], ignore_index=True)
        
        columnas = [str(c).strip().upper() for c in df_filtrado.columns]
        if "FECHA" in columnas and "REFERENCIA" in columnas:
            rename_map = {
                "NRO.": "NRO", "NRO": "NRO", "FECHA": "FECHA", "REFERENCIA": "REFERENCIA",
                "CONCEPTO": "DESCRIPCION", "DÉBITO": "DEBITO", "DEBITO": "DEBITO",
                "CRÉDITO": "CREDITO", "CREDITO": "CREDITO", "SALDO": "SALDO"
            }
            df_filtrado.columns = [rename_map.get(str(c).strip().upper(), str(c).strip().upper()) for c in df_filtrado.columns]
        else:
            encabezado_idx = None
            for i in range(min(30, len(df_filtrado))):
                fila = df_filtrado.iloc[i]
                fila_str = [str(val) for val in fila.tolist()]
                texto_fila = " ".join(fila_str).upper()
                if "NRO" in texto_fila and "FECHA" in texto_fila and "REFERENCIA" in texto_fila:
                    encabezado_idx = i
                    break
            if encabezado_idx is None: return pd.DataFrame()
            headers = df_filtrado.iloc[encabezado_idx].astype(str).str.strip().tolist()
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
            df_filtrado.columns = headers
            df_filtrado = df_filtrado.iloc[encabezado_idx + 1:].reset_index(drop=True)
            df_filtrado = df_filtrado.rename(columns=rename_map)
            
        if "FECHA" not in df_filtrado.columns: return pd.DataFrame()
        # Filtrar filas que no son movimientos
        fechas_str_col = df_filtrado["FECHA"].astype(str).str.strip()
        df_filtrado = df_filtrado[~fechas_str_col.str.contains("FECHA|SALDO|TOTAL|CRÉDITO|CREDITO|DÉBITO|DEBITO", case=False, na=False)]
        
        # Convertir fechas de manera robusta
        df_filtrado["FECHA_DT"] = df_filtrado["FECHA"].apply(parsear_fecha_flexible)
        df_filtrado = df_filtrado[df_filtrado["FECHA_DT"].notna()]
        df_filtrado["FECHA"] = df_filtrado["FECHA_DT"]
        df_filtrado = df_filtrado[df_filtrado["FECHA"].notna()]
        
        def limpiar_monto(val):
            val_str = str(val).strip().replace(" ", "")
            if not val_str or val_str == "nan":
                return 0.0
            if "," in val_str:
                val_str = val_str.replace(".", "").replace(",", ".")
            return pd.to_numeric(val_str, errors="coerce")

        df_filtrado["DEBITO"] = df_filtrado["DEBITO"].apply(limpiar_monto).fillna(0) if "DEBITO" in df_filtrado.columns else 0.0
        df_filtrado["CREDITO"] = df_filtrado["CREDITO"].apply(limpiar_monto).fillna(0) if "CREDITO" in df_filtrado.columns else 0.0
        
        df_filtrado["MONTO"] = df_filtrado["CREDITO"] - df_filtrado["DEBITO"]
        df_filtrado["TIPO"] = df_filtrado["MONTO"].apply(lambda x: "NC" if x > 0 else "ND" if x < 0 else "")
        df_filtrado["MONTO"] = df_filtrado["MONTO"].abs()
        df_filtrado = df_filtrado[df_filtrado["MONTO"] > 0]
        if "REFERENCIA" not in df_filtrado.columns: df_filtrado["REFERENCIA"] = ""
        else: df_filtrado["REFERENCIA"] = df_filtrado["REFERENCIA"].astype(str).str.strip().str.replace("'", "", regex=False)
        if "DESCRIPCION" not in df_filtrado.columns: df_filtrado["DESCRIPCION"] = ""
        else: df_filtrado["DESCRIPCION"] = df_filtrado["DESCRIPCION"].astype(str).str.strip()
        df_filtrado["ES_COMISION"] = df_filtrado["DESCRIPCION"].str.contains("Comisi", case=False, na=False)
        df_resultado = df_filtrado[["FECHA", "REFERENCIA", "DESCRIPCION", "TIPO", "MONTO", "ES_COMISION"]].copy()
        st.success(f"✅ Bancamiga OK: {len(df_resultado)} movimientos")
        return df_resultado
    except Exception as e:
        st.error(f"❌ Error Bancamiga: {str(e)}")
        return pd.DataFrame()

def procesar_banplus(df):
    st.info("🔍 Procesando archivo de BanPlus...")
    try:
        # 🔥 EXTRAER RESUMEN DEL ARCHIVO (Saldo Total, Saldo Inicial)
        resumen_archivo = extraer_resumen_banplus(df)
        if any(v is not None for v in resumen_archivo.values()):
            saldo_total = formato_venezolano(resumen_archivo['saldo_total']) if resumen_archivo['saldo_total'] is not None else "N/A"
            saldo_inicial = formato_venezolano(resumen_archivo['saldo_inicial']) if resumen_archivo['saldo_inicial'] is not None else "N/A"
            st.info(f"""
            📊 **BanPlus - Resumen del archivo:**
            - Saldo Total: {saldo_total} BS
            - Saldo Inicial: {saldo_inicial} BS
            """)
        
        columnas = [str(c).strip().upper() for c in df.columns]
        tiene_encabezados = "FECHA" in columnas and "REFERENCIA" in columnas
        
        # 🔥 1. BUSCAR EL ENCABEZADO EN EL DF ORIGINAL (antes de filtrar por fecha)
        encabezado_idx = None
        if not tiene_encabezados:
            for i in range(min(30, len(df))):
                fila = df.iloc[i]
                fila_str = [str(val) for val in fila.tolist()]
                texto_fila = " ".join(fila_str).upper()
                if "FECHA" in texto_fila and ("REFERENCIA" in texto_fila or "REF" in texto_fila or "DESCRIPCION" in texto_fila or "CONCEPTO" in texto_fila):
                    encabezado_idx = i
                    break
        
        if tiene_encabezados:
            df_datos = df.copy()
        elif encabezado_idx is not None:
            fila_encabezado = df.iloc[[encabezado_idx]].copy()
            df_datos = df.iloc[encabezado_idx + 1:].copy()
        else:
            df_datos = df.copy()
            st.warning("⚠️ BanPlus: No se encontró la fila de encabezados, se procesará todo el archivo.")
        
        # 🔥 2. APLICAR FILTRO DE FECHA PREDOMINANTE SOLO A LAS FILAS DE DATOS
        df_filtrado, fecha_pred, dict_conteo, total_filas, registros_excluidos = filtrar_por_fecha_predominante(
            df_datos, columna_fecha_idx=0, nombre_banco="BanPlus"
        )
        
        if df_filtrado is None or df_filtrado.empty:
            df_filtrado = df_datos
            st.warning("⚠️ BanPlus: No se pudo filtrar por fecha predominante, se procesará todo el archivo.")
        
        if fecha_pred:
            _acumular_info_fechas("BanPlus", fecha_pred.strftime('%d/%m/%Y'), len(df_filtrado), total_filas, registros_excluidos, dict_conteo)
        # 🔥 3. REENSAMBLAR: encabezado + datos filtrados
        if not tiene_encabezados and encabezado_idx is not None:
            df_filtrado = pd.concat([fila_encabezado, df_filtrado], ignore_index=True)
        
        columnas = [str(c).strip().upper() for c in df_filtrado.columns]
        if "FECHA" in columnas and "REFERENCIA" in columnas:
            rename_map = {
                "FECHA": "FECHA", "REFERENCIA": "REFERENCIA", "DESCRIPCION": "DESCRIPCION",
                "CONCEPTO": "DESCRIPCION", "DEBITO": "DEBITO", "DEBITOS": "DEBITO",
                "DÉBITO": "DEBITO", "CREDITO": "CREDITO", "CREDITOS": "CREDITO",
                "CRÉDITO": "CREDITO", "SALDO": "SALDO"
            }
            df_filtrado.columns = [rename_map.get(str(c).strip().upper(), str(c).strip().upper()) for c in df_filtrado.columns]
        else:
            encabezado_idx = None
            for i in range(min(30, len(df_filtrado))):
                fila = df_filtrado.iloc[i]
                fila_str = [str(val) for val in fila.tolist()]
                texto_fila = " ".join(fila_str).upper()
                if "FECHA" in texto_fila and ("REFERENCIA" in texto_fila or "REF" in texto_fila or "DESCRIPCION" in texto_fila or "CONCEPTO" in texto_fila):
                    encabezado_idx = i
                    break
            if encabezado_idx is None: return pd.DataFrame()
            headers = df_filtrado.iloc[encabezado_idx].astype(str).str.strip().tolist()
            rename_map = {}
            for col in headers:
                col_clean = str(col).strip().upper()
                if "FECHA" in col_clean: rename_map[col] = "FECHA"
                elif "REFERENCIA" in col_clean or "REF" in col_clean: rename_map[col] = "REFERENCIA"
                elif "CONCEPTO" in col_clean or "DESCRIP" in col_clean: rename_map[col] = "DESCRIPCION"
                elif "DÉBITO" in col_clean or "DEBITO" in col_clean or "EGRESO" in col_clean or "CARGO" in col_clean: rename_map[col] = "DEBITO"
                elif "CRÉDITO" in col_clean or "CREDITO" in col_clean or "INGRESO" in col_clean or "ABONO" in col_clean: rename_map[col] = "CREDITO"
                elif "SALDO" in col_clean: rename_map[col] = "SALDO"
            df_filtrado.columns = headers
            df_filtrado = df_filtrado.iloc[encabezado_idx + 1:].reset_index(drop=True)
            df_filtrado = df_filtrado.rename(columns=rename_map)
            
        if "FECHA" not in df_filtrado.columns: return pd.DataFrame()
        df_filtrado["FECHA"] = df_filtrado["FECHA"].astype(str).str.strip()
        df_filtrado = df_filtrado[~df_filtrado["FECHA"].str.contains("FECHA|SALDO|TOTAL|CRÉDITO|CREDITO|DÉBITO|DEBITO", case=False, na=False)]
        # 🔥 FIX (2026-08-21): parseo robusto sin perder filas
        # El Excel guarda fechas como datetime -> "2026-08-07 00:00:00" (ISO)
        # El formato estricto "%d/%m/%Y" las convertia a NaT y se perdian TODAS las filas
        df_filtrado["FECHA_DT"] = df_filtrado["FECHA"].apply(parsear_fecha_flexible)
        df_filtrado["FECHA"] = df_filtrado["FECHA_DT"]
        df_filtrado = df_filtrado[df_filtrado["FECHA"].notna()]
        
        def mono_limpiar_monto_banplus(valor):
            if valor is None or pd.isna(valor):
                return 0.0
            if isinstance(valor, (int, float)):
                return float(valor)
            valor_str = str(valor).strip()
            if not valor_str:
                return 0.0
            valor_str = valor_str.replace('$', '').replace('Bs.', '').replace('Bs', '').replace(' ', '').strip()
            try:
                return float(valor_str)
            except ValueError:
                pass
            has_comma = ',' in valor_str
            has_dot = '.' in valor_str
            if has_comma and has_dot:
                pos_comma = valor_str.rfind(',')
                pos_dot = valor_str.rfind('.')
                if pos_dot > pos_comma:
                    valor_limpio = valor_str.replace(',', '')
                else:
                    valor_limpio = valor_str.replace('.', '').replace(',', '.')
            elif has_comma:
                valor_limpio = valor_str.replace(',', '.')
            elif has_dot:
                valor_limpio = valor_str
            else:
                valor_limpio = valor_str
            try:
                return float(valor_limpio)
            except ValueError:
                return 0.0
        
        df_filtrado["DEBITO"] = df_filtrado["DEBITO"].apply(mono_limpiar_monto_banplus) if "DEBITO" in df_filtrado.columns else 0.0
        df_filtrado["CREDITO"] = df_filtrado["CREDITO"].apply(mono_limpiar_monto_banplus) if "CREDITO" in df_filtrado.columns else 0.0
        
        df_filtrado["MONTO"] = df_filtrado["CREDITO"] - df_filtrado["DEBITO"]
        df_filtrado["TIPO"] = df_filtrado["MONTO"].apply(lambda x: "NC" if x > 0 else "ND" if x < 0 else "")
        df_filtrado["MONTO"] = df_filtrado["MONTO"].abs()
        df_filtrado = df_filtrado[df_filtrado["MONTO"] > 0]
        if "REFERENCIA" not in df_filtrado.columns: df_filtrado["REFERENCIA"] = ""
        else: df_filtrado["REFERENCIA"] = df_filtrado["REFERENCIA"].astype(str).str.strip().str.replace("'", "", regex=False)
        if "DESCRIPCION" not in df_filtrado.columns: df_filtrado["DESCRIPCION"] = ""
        else: df_filtrado["DESCRIPCION"] = df_filtrado["DESCRIPCION"].astype(str).str.strip()
        df_filtrado["ES_COMISION"] = df_filtrado["DESCRIPCION"].str.contains("Comisi|sms|servicio sms|sms plus", case=False, na=False)
        df_resultado = df_filtrado[["FECHA", "REFERENCIA", "DESCRIPCION", "TIPO", "MONTO", "ES_COMISION"]].copy()
        st.success(f"✅ BanPlus OK: {len(df_resultado)} movimientos")
        return df_resultado
    except Exception as e:
        st.error(f"❌ Error BanPlus: {str(e)}")
        return pd.DataFrame()

def procesar_venezuela_simple(df):
    st.info("🔍 Procesando Banco de Venezuela (MODO SIMPLE)...")
    try:
        # 🔥 APLICAR FILTRO DE FECHA PREDOMINANTE (BDV tiene fechas en columna 3)
        df_filtrado, fecha_pred, dict_conteo, total_filas, registros_excluidos = filtrar_por_fecha_predominante(
            df, columna_fecha_idx=3, nombre_banco="Banco de Venezuela"
        )
        
        if df_filtrado is None or df_filtrado.empty:
            df_filtrado = df
            st.warning("⚠️ Banco de Venezuela: No se pudo filtrar por fecha predominante, se procesará todo el archivo.")
        
        if fecha_pred:
            _acumular_info_fechas("Banco de Venezuela", fecha_pred.strftime('%d/%m/%Y'), len(df_filtrado), total_filas, registros_excluidos, dict_conteo)
        col_fecha = 3
        col_ref = 1
        col_desc = 2
        col_tipo = 4
        col_credito = 5
        col_debito = 6
        
        movimientos = []
        # 🔧 CORRECCIÓN (2026-08-21): filtrar_por_fecha_predominante YA eliminó el encabezado,
        # así que el primer registro de datos es la fila 0. Antes se arrancaba en 1 y se perdía
        # el primer movimiento (p.ej. el PAGO A PROVEEDORES de 14.500,00 del BDV).
        for idx in range(0, len(df_filtrado)):
            try:
                fila = df_filtrado.iloc[idx]
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
                    # 🔥 CORREGIDO: Si no tiene tipo definido, determinar por crédito/débito
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
# 🔥 FUNCIÓN PARA PROCESAR BANCO ACTIVO - VERSIÓN CORREGIDA
# =========================================================

def procesar_banco_activo(df):
    """
    Procesa archivo de Banco Activo con formato específico.
    El archivo tiene formato: Número de Cuenta en primera fila, luego columnas:
    Fecha | Fecha Valor | Concepto | # | Débito | Crédito | Saldo
    """
    st.info("🔍 Procesando archivo de Banco Activo...")
    
    try:
        # 🔥 APLICAR FILTRO DE FECHA PREDOMINANTE (Banco Activo tiene fechas en columna 0)
        df_filtrado, fecha_pred, dict_conteo, total_filas, registros_excluidos = filtrar_por_fecha_predominante(
            df, columna_fecha_idx=0, nombre_banco="Banco Activo"
        )
        
        if df_filtrado is None or df_filtrado.empty:
            df_filtrado = df
            st.warning("⚠️ Banco Activo: No se pudo filtrar por fecha predominante, se procesará todo el archivo.")
        
        if fecha_pred:
            _acumular_info_fechas("Banco Activo", fecha_pred.strftime('%d/%m/%Y'), len(df_filtrado), total_filas, registros_excluidos, dict_conteo)
        # Mostrar información del archivo
        st.write("📊 **Información del archivo:**")
        st.write(f"- Número de filas: {len(df_filtrado)}")
        st.write(f"- Número de columnas: {len(df_filtrado.columns)}")
        
        # Mostrar primeras filas para debug
        st.write("👁️ **Primeras 15 filas del archivo:**")
        st.dataframe(df_filtrado.head(15))
        
        # Saltar la primera fila (Número de Cuenta) y empezar desde la fila 1
        df_datos = df_filtrado.iloc[1:].copy().reset_index(drop=True)
        
        # 🔥 DETECTAR EL NÚMERO DE COLUMNAS Y ASIGNAR NOMBRES CORRECTAMENTE
        num_columnas = df_datos.shape[1]
        st.write(f"📋 **El archivo tiene {num_columnas} columnas**")
        
        if num_columnas == 8:
            # Si tiene 8 columnas, la columna extra probablemente es una columna vacía o de índice
            # Asignamos nombres y luego eliminamos la columna extra o la ignoramos
            df_datos.columns = ["FECHA", "FECHA_VALOR", "DESCRIPCION", "NRO", "DEBITO", "CREDITO", "SALDO", "EXTRA"]
            # Eliminar la columna extra
            df_datos = df_datos.drop(columns=["EXTRA"])
        elif num_columnas == 7:
            df_datos.columns = ["FECHA", "FECHA_VALOR", "DESCRIPCION", "NRO", "DEBITO", "CREDITO", "SALDO"]
        else:
            st.error(f"❌ El archivo tiene {num_columnas} columnas, se esperaban 7 u 8.")
            return pd.DataFrame()
        
        st.write("📋 **Estructura de datos asignada:**")
        st.write(f"- Columnas: {df_datos.columns.tolist()}")
        
        # 🔥 PROCESAR FECHAS DE BANCO ACTIVO
        # Convertir fechas (formato DD/MM/YYYY)
        def parse_fecha(val):
            if pd.isna(val):
                return pd.NaT
            val_str = str(val).strip()
            if not val_str:
                return pd.NaT
            # Intentar diferentes formatos
            try:
                return pd.to_datetime(val_str, format="%d/%m/%Y", errors="coerce")
            except:
                try:
                    return pd.to_datetime(val_str, dayfirst=True, errors="coerce")
                except:
                    return pd.NaT
        
        df_datos["FECHA_DT"] = df_datos["FECHA"].apply(parse_fecha)
        df_datos = df_datos[df_datos["FECHA_DT"].notna()]
        
        # Procesar débito y crédito
        def limpiar_monto(val):
            val_str = str(val).strip().replace(" ", "")
            if not val_str or val_str == "nan" or val_str == "":
                return 0.0
            # Limpiar formato venezolano (puntos como separadores de miles, coma decimal)
            if "," in val_str:
                val_str = val_str.replace(".", "").replace(",", ".")
            try:
                return float(val_str)
            except:
                return 0.0
        
        df_datos["DEBITO"] = df_datos["DEBITO"].apply(limpiar_monto)
        df_datos["CREDITO"] = df_datos["CREDITO"].apply(limpiar_monto)
        
        # Determinar tipo y monto
        df_datos["MONTO"] = df_datos["CREDITO"] - df_datos["DEBITO"]
        df_datos["TIPO"] = df_datos["MONTO"].apply(lambda x: "NC" if x > 0 else "ND" if x < 0 else "")
        df_datos["MONTO"] = df_datos["MONTO"].abs()
        
        # Eliminar filas con monto 0
        df_datos = df_datos[df_datos["MONTO"] > 0]
        
        # Asegurar que existe columna REFERENCIA
        df_datos["REFERENCIA"] = df_datos["NRO"].astype(str).str.strip()
        
        # Asegurar que existe columna DESCRIPCION
        df_datos["DESCRIPCION"] = df_datos["DESCRIPCION"].astype(str).str.strip()
        
        # 🔥 DETECTAR COMISIONES DE BANCO ACTIVO
        palabras_comision = [
            "CARGO POR MANTENIMIENTO",
            "CARGO EMISION EDO DE CUENTA",
            "CARGO SERVICIO SMS",
            "COMISION",
            "COMISIÓN",
            "MANTENIMIENTO",
            "SMS",
            "EMISION EDO",
            "CARGO POR",
            "COM EDO",
            "COM MOV"
        ]
        
        df_datos["ES_COMISION"] = df_datos["DESCRIPCION"].str.contains('|'.join(palabras_comision), case=False, na=False)
        
        # También detectar comisiones por monto pequeño (típico de cargos bancarios)
        mascara_monto_pequeno = df_datos["MONTO"] < 1000
        df_datos.loc[mascara_monto_pequeno, "ES_COMISION"] = (
            df_datos.loc[mascara_monto_pequeno, "ES_COMISION"] | 
            df_datos.loc[mascara_monto_pequeno, "DESCRIPCION"].str.contains("CARGO|COM|MANTENIMIENTO|SMS", case=False, na=False)
        )
        
        # 🔥 DEBUG: Mostrar cuántas comisiones se detectaron
        num_comisiones = df_datos["ES_COMISION"].sum()
        st.info(f"💳 Se detectaron {num_comisiones} comisiones en el archivo de Banco Activo")
        
        # Mostrar las comisiones detectadas
        if num_comisiones > 0:
            st.write("📋 **Comisiones detectadas:**")
            st.dataframe(df_datos[df_datos["ES_COMISION"] == True][["FECHA", "REFERENCIA", "DESCRIPCION", "MONTO"]])
        
        # Seleccionar solo las columnas necesarias
        df_resultado = df_datos[["FECHA_DT", "REFERENCIA", "DESCRIPCION", "TIPO", "MONTO", "ES_COMISION"]].copy()
        df_resultado = df_resultado.rename(columns={"FECHA_DT": "FECHA"})
        
        # Mostrar resultados
        st.success(f"✅ Banco Activo OK: {len(df_resultado)} movimientos detectados")
        st.dataframe(df_resultado.head(10))
        
        return df_resultado
        
    except Exception as e:
        st.error(f"❌ Error procesando Banco Activo: {str(e)}")
        import traceback
        st.code(traceback.format_exc())
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
    df_temp = df.copy()
    
    # Si es Mercantil o no hay columnas de texto esperadas, mapeamos por índice
    if banco == "mercantil" or "FECHA" not in [str(c).strip().upper() for c in df_temp.columns]:
        for idx, fila in df_temp.iterrows():
            try:
                if len(fila) < 8: continue
                fecha = fila.iloc[3]
                if pd.isna(fecha): continue
                fecha_str = str(fecha).strip().replace(".0", "")
                # 🔧 CORRECCIÓN (2026-08-21): el archivo Mercantil trae la fecha como ddmmyyyy
                # sin separadores y con día de 1-2 dígitos (p.ej. "7082026" = 07/08/2026).
                # Se normaliza a dd/mm/yyyy para que el movimiento no se pierda en el
                # consolidado MULTIBANCO.
                if re.fullmatch(r"\d{7,8}", fecha_str):
                    anio = fecha_str[-4:]
                    mes = fecha_str[-6:-4]
                    dia = fecha_str[:-6].zfill(2)
                    fecha_str = f"{dia}/{mes}/{anio}"
                
                tipo = str(fila.iloc[5]).strip()
                descripcion = str(fila.iloc[6]).strip()
                referencia = str(fila.iloc[4]).strip()
                monto = fila.iloc[7]
                # 🔧 CORRECCIÓN (2026-08-21): normalizar el monto del Mercantil (formato VES
                # "136.207,80" → 136207.8) para que _sumar_creditos_convertidos y el REPORTE
                # consolidado lo usen correctamente.
                if isinstance(monto, str):
                    monto_clean = monto.strip().replace(" ", "").replace(".", "").replace(",", ".")
                    try:
                        monto = float(monto_clean)
                    except ValueError:
                        monto = 0.0
                else:
                    try:
                        monto = float(monto)
                    except (TypeError, ValueError):
                        monto = 0.0
                # 🔧 CORRECCIÓN (2026-08-21): en Mercantil la columna 9 es el Nº de transacción,
                # NO un indicador de comisión. Usarla como flag hacía que _sumar_creditos_convertidos
                # descartara TODOS los movimientos (ingresos = 0). La clasificación de comisiones
                # la hace procesar_archivo con patrones_comision_por_banco("mercantil").
                es_comision_flag = False
                
                fila_convertida = ["", "", "", fecha_str, referencia, tipo, descripcion, monto, "", es_comision_flag]
                datos_convertidos.append(fila_convertida)
            except:
                continue
    else:
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
# API TASA BCV AUTOMÁTICA (lab.geocenso.com) + HISTÓRICO PERSISTENTE
# =========================================================

BCV_API_URL = "https://lab.geocenso.com/tasas/api_bcv.php"
BCV_API_HEADERS = {"X-API-Key": "2a0ad52cc3e2c0632180e790f5b322cfe7a2281362776a0d"}
BCV_TASAS_AUTO_JSON = os.path.join(os.path.dirname(os.path.abspath(__file__)), "tasas_bcv_auto.json")
_tasas_auto = None

def _tasas_auto_data():
    """Registro automático persistente de tasas obtenidas de la API (sobrevive reinicios)."""
    global _tasas_auto
    if _tasas_auto is None:
        _tasas_auto = {}
        try:
            if os.path.exists(BCV_TASAS_AUTO_JSON):
                with open(BCV_TASAS_AUTO_JSON, "r", encoding="utf-8") as f:
                    datos = json.load(f)
                if isinstance(datos, dict):
                    _tasas_auto = {k: float(v) for k, v in datos.items() if k and v}
        except Exception:
            _tasas_auto = {}
    return _tasas_auto

def _tasa_recordar(fecha_str, tasa):
    """Guarda la tasa resuelta por la API para esa fecha en el registro automático."""
    if not fecha_str or not tasa:
        return
    auto = _tasas_auto_data()
    if auto.get(fecha_str) != tasa:
        auto[fecha_str] = tasa
        try:
            with open(BCV_TASAS_AUTO_JSON, "w", encoding="utf-8") as f:
                json.dump(auto, f, ensure_ascii=False, indent=1)
        except Exception:
            pass

def _bcv_api_consultar():
    """Consulta la tasa vigente del BCV en la API remota. Devuelve dict con 'usd' y 'fecha', o None si falla."""
    try:
        resp = requests.get(BCV_API_URL, headers=BCV_API_HEADERS, timeout=10)
        resp.raise_for_status()
        data = resp.json()
        usd = float(data.get("usd") or 0)
        if usd <= 0:
            return None
        f_obj = pd.to_datetime(str(data.get("fecha", "")), errors="coerce")
        if pd.isna(f_obj):
            return None
        return {"usd": usd, "fecha": f_obj.date()}
    except Exception:
        return None

@st.cache_data(ttl=600, show_spinner=False)
def _bcv_api_tasa_cache():
    """Tasa BCV vigente en caché corta (10 min) para no golpear la API en cada transacción."""
    return _bcv_api_consultar()

def _tasa_con_origen_por_fecha(fecha_obj, usar_api=False):
    """Resuelve la tasa BCV para la fecha SELECCIONADA consultando en este orden:
    1) histórico local canónico (código) → 2) registro automático persistente de la API
    3) fin de semana: última tasa publicada antes → 4) API en vivo para el día/desfase reciente.
    Devuelve (tasa, origen) o (None, "")."""
    if fecha_obj is None:
        fecha_obj = date.today()
    fecha_str = fecha_obj.strftime("%d/%m/%Y")

    # 1) Histórico canónico del código (fechas ya registradas)
    tasa = obtener_tasa_bcv_fecha(fecha_obj)
    if tasa is not None:
        return tasa, "Histórico local"

    # 2) Registro automático persistente (tasas que la API devolvió antes para esa fecha)
    auto = _tasas_auto_data().get(fecha_str)
    if auto is not None:
        return auto, "API registrada automáticamente"

    if not usar_api:
        return None, ""

    hoy = date.today()
    if fecha_obj > hoy:
        return None, ""

    # 3) Sábado/domingo: el BCV no publica, la tasa no cambia el fin de semana
    if fecha_obj.weekday() >= 5:
        for i in range(1, 4):
            prev = fecha_obj - timedelta(days=i)
            t_prev, _ = _tasa_con_origen_por_fecha(prev, True)
            if t_prev is not None:
                _tasa_recordar(fecha_str, t_prev)
                return t_prev, f"Fin de semana (última publicada {prev.strftime('%d/%m/%Y')})"
        # si no hay conocida previa, cae al bloque siguiente (API vigente reciente)

    # 4) API en vivo: tasa vigente del BCV (el día publicado y desfases recientes de hasta 3 días)
    api = _bcv_api_tasa_cache()
    if api:
        dias_desfase = (api["fecha"] - fecha_obj).days
        if 0 <= dias_desfase <= 3:
            _tasa_recordar(fecha_str, api["usd"])
            if dias_desfase == 0:
                return api["usd"], f"API BCV en vivo ({api['fecha'].strftime('%d/%m/%Y')})"
            return api["usd"], f"Tasa vigente BCV del {api['fecha'].strftime('%d/%m/%Y')} (día sin publicar)"
    return None, ""

def _tasa_con_origen(fecha=None, usar_api=False):
    """Igual que _tasa_con_origen_por_fecha pero con respaldo final a la última tasa conocida."""
    if fecha is None:
        fecha = date.today()
    tasa, origen = _tasa_con_origen_por_fecha(fecha, usar_api)
    if tasa is not None:
        return tasa, origen
    api = _bcv_api_tasa_cache()
    if api:
        return api["usd"], f"Última conocida (API {api['fecha'].strftime('%d/%m/%Y')})"
    hoy = date.today()
    for i in range(0, 15):
        t_prev = obtener_tasa_bcv_fecha(hoy - timedelta(days=i))
        if t_prev is not None:
            return t_prev, "Última conocida (histórico)"
    return 757.5406, "Última conocida (valor fijo)"

# =========================================================
# OBTENER TASA BCV HISTORICA LOCAL
# =========================================================

def obtener_tasa_bcv_fecha(fecha_obj):
    """Histórico local canónico: devuelve la tasa oficial ya registrada en el código para esa fecha exacta."""
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
        "01/07/2026": 633.3644, "02/07/2026": 639.7029, "03/07/2026": 652.9726,
        "04/07/2026": 667.0500, "05/07/2026": 667.0500, "06/07/2026": 667.0500,
        "07/07/2026": 674.9305, "08/07/2026": 685.9427, "09/07/2026": 700.2249,
        "10/07/2026": 709.6935, "11/07/2026": 721.3456, "12/07/2026": 721.3456,
        "13/07/2026": 721.3456, "14/07/2026": 723.9990, "15/07/2026": 725.7470,
        "16/07/2026": 727.4512, "17/07/2026": 732.4787, "18/07/2026": 736.9339,
        "19/07/2026": 736.9339, "20/07/2026": 736.9339, "21/07/2026": 737.2321,
        "22/07/2026": 737.2321, "23/07/2026": 737.8816, "24/07/2026": 742.2292,
        "25/07/2026": 742.2292, "26/07/2026": 742.2292, "27/07/2026": 742.2292,
        "28/07/2026": 742.8105, "29/07/2026": 744.2264, "30/07/2026": 745.6371,
        "31/07/2026": 746.6297,
        "01/08/2026": 748.7864, "02/08/2026": 748.7864, "03/08/2026": 748.7864,
        "04/08/2026": 752.0943, "05/08/2026": 755.1552, "06/08/2026": 755.9001,
        "07/08/2026": 756.7083, "08/08/2026": 757.5406, "09/08/2026": 757.5406,
        "10/08/2026": 757.5406, "11/08/2026": 761.2167, "12/08/2026": 764.3486,
        "13/08/2026": 766.8603, "14/08/2026": 771.0714, "15/08/2026": 772.5441,
        "16/08/2026": 772.5441, "17/08/2026": 772.5441, "18/08/2026": 773.3125,
        "19/08/2026": 775.3356, "20/08/2026": 777.4161, "21/08/2026": 779.9522,
        "22/08/2026": 784.6633, "23/08/2026": 784.6633, "24/08/2026": 784.6633,
        "25/08/2026": 785.0693, "26/08/2026": 787.5196, "27/08/2026": 791.3248,
        "28/08/2026": 791.6667, "29/08/2026": 794.9917, "30/08/2026": 794.9917,
        "31/08/2026": 794.9917,
        "01/09/2026": 798.3260, "02/09/2026": 801.1752, "03/09/2026": 804.8109,
        "04/09/2026": 807.3862,
    }
    fecha_str = fecha_obj.strftime("%d/%m/%Y")
    return tasas_bcv_local.get(fecha_str, None)

def obtener_tasa_por_fecha(fecha_obj, usar_api=False):
    """Tasa BCV de una fecha (resolución automática por fecha seleccionada)."""
    tasa, _ = _tasa_con_origen_por_fecha(fecha_obj, usar_api)
    return tasa

# =========================================================
# 🔥 DETECCIÓN DE COMISIONES POR BANCO (FUNCIÓN ÚNICA)
# =========================================================

def patrones_comision_por_banco(banco):
    """Devuelve la lista de patrones de comisión de texto para un banco"""
    if banco == "venezuela":
        return [
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
            "COMISION X PAGO DE NOMINAS MB",
            "DOMICILIACION J412438905",
            "DISTRIBUIDORA GLOBAL",
            "DOMICILIACION",
            "COM TRANSF LINEA OT BCOS CTA P",
            "COM TRANSF LINEA OT BCOS",
            "COM TRANSF LINEA",
            "COM TRANSF EN LINEA",
            "COM TRANSF"
        ]
    elif banco == "banesco":
        return ["COMISION", "COMIS", "CARGO", "ITF", "IMPUESTO"]
    elif banco == "bnc":
        return ["COMISION", "COMIS", "CARGO", "ITF"]
    elif banco == "provincial":
        return ["COMIS", "COM TRF", "COM TRF TERCERO", "COMISION COBRADA", "COMISION POR"]
    elif banco == "bancamiga":
        return ["COMISI"]
    elif banco == "mercantil":
        return [
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
            "COMISION POR TRANSFERENCIA",
            "EMISION EDO",
            "RETENCION DE IMPUESTO",
            "DESCUENTO DE TARJETA",
            "DESC. TARJETA",
            "DESCUENTO TARJETA",
            "DESCUENTO TARJETA CREDITO",
            "DESC TARJETA"
        ]
    elif banco == "tesoro":
        return [
            "BELOW MINIMUM BALANCE CHARGES",
            "STAMENT SERVICE",
            "STATEMENT SERVICE",
            "COMIS",
            "COMISION",
            "CARGO BANCARIO",
            "CARGO POR SERVICIO"
        ]
    elif banco == "activo":
        return [
            "CARGO POR MANTENIMIENTO",
            "CARGO EMISION EDO DE CUENTA",
            "CARGO SERVICIO SMS",
            "COMISION",
            "COMISIÓN",
            "MANTENIMIENTO",
            "SMS",
            "EMISION EDO",
            "COM EDO",
            "COM MOV"
        ]
    elif banco == "banplus":
        return ["COMISION", "COMIS", "SMS", "CARGO", "MANTENIMIENTO"]
    return []

def detectar_comision_por_banco(descripcion, referencia="", tipo="", monto_bs=0, banco=""):
    """
    Detecta si un movimiento es comisión bancaria según reglas específicas del banco.
    En modo multibanco se aplica la unión de patrones de TODOS los bancos.
    """
    descripcion_upper = str(descripcion or "").upper()
    referencia_upper = str(referencia or "").upper()
    tipo_upper = str(tipo or "").upper()

    # 🔥 REGLA GENÉRICA: "COMISION PAGO A PROVEEDORES" (y variantes) SIEMPRE
    # es una comisión bancaria, sin importar el banco.
    if any(x in descripcion_upper for x in [
        "COMISION PAGO A PROVEEDORES",
        "COMISION PAGO A PROVEEDOR",
        "COM PAGO A PROVEEDORES",
        "COM. PAGO A PROVEEDORES",
        "COM PAGO PROVEEDORES"
    ]):
        return True

    if banco == "multibanco":
        for b in ["venezuela", "banesco", "bnc", "provincial", "bancamiga", "mercantil", "tesoro", "activo", "banplus"]:
            for patron in patrones_comision_por_banco(b):
                if patron in descripcion_upper:
                    return True
        return False

    # BANCO DE VENEZUELA - REGLAS ESPECÍFICAS
    if banco == "venezuela":
        # 🔥 REGLA 1: Detectar por descripción (comisiones de BDV)
        for patron in patrones_comision_por_banco("venezuela"):
            if patron in descripcion_upper:
                return True

        # 🔥 REGLA 2: Detectar por referencia (comisiones de BDV tienen referencias específicas)
        if referencia_upper.startswith(("970", "972", "067")):
            if any(palabra in descripcion_upper for palabra in ["COM", "PAGO OTRAS", "PAGO OTR", "COMISION"]):
                return True

        # 🔥 REGLA 3: Si el tipo es ND y la descripción contiene "COM" es una comisión
        if tipo_upper == "ND":
            if "COM" in descripcion_upper or "PAGO OTR" in descripcion_upper:
                return True

        # 🔥 REGLA 4: Comisiones específicas de BDV por monto pequeño
        if tipo_upper == "ND":
            if monto_bs < 1000 and ("COM" in descripcion_upper or "PAGO OTR" in descripcion_upper):
                return True

    # Resto de bancos: matching de patrones de texto
    for patron in patrones_comision_por_banco(banco):
        if patron in descripcion_upper:
            return True

    return False

# =========================================================
# 🔥 PROCESAMIENTO PRINCIPAL - CLASIFICACIÓN (CORREGIDO PARA BDV)
# =========================================================

def procesar_archivo(df, usar_api=False, banco=""):
    ingresos = []
    egresos = []
    comisiones = []
    registros_procesados = set()
    
    tipos_ingresos = ["NC", "C", "CREDITO", "ABONO", "DP", "DEP"]
    tipos_egresos = ["ND", "D", "DEBITO", "DEBIT"]
    cache_tasas = {}

    for _, fila in df.iterrows():
        try:
            if len(fila) < 10: continue
            fecha_raw = str(fila[3]).strip()
            if fecha_raw.lower() == "nan": continue
            fecha_raw = fecha_raw.replace(".0", "")
            if len(fecha_raw) == 7 and fecha_raw.isdigit():
                fecha = f"0{fecha_raw[0]}/{fecha_raw[1:3]}/{fecha_raw[3:]}"
            elif len(fecha_raw) == 8 and fecha_raw.isdigit():
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

            # 🔥 DETECCIÓN DE COMISIONES POR BANCO (función única con soporte multibanco)
            es_comision_banco = detectar_comision_por_banco(
                descripcion, referencia, tipo, monto_bs, banco
            )

            # Si es comisión del banco, clasificar como comisión
            if es_comision_banco:
                comisiones.append(registro)
            # Si es comisión bancaria detectada por función genérica
            elif es_comision(descripcion):
                comisiones.append(registro)
            # Clasificar por tipo de movimiento
            elif tipo in tipos_ingresos:
                ingresos.append(registro)
            elif tipo in tipos_egresos:
                egresos.append(registro)
            else:
                # 🔥 NUEVO: Si no tiene tipo definido, clasificar por monto (crédito = ingreso, débito = egreso)
                if monto_bs > 0:
                    ingresos.append(registro)
                else:
                    egresos.append(registro)
        except:
            continue
    return ingresos, egresos, comisiones

# =========================================================
# INTERFAZ PRINCIPAL - EJECUCIÓN
# =========================================================

# =========================================================
# 🔥 MONOBANCO FUNCTIONS (NAMESPACED TO AVOID OVERWRITING)
# =========================================================

def mono_leer_excel_sin_encabezados(archivo):
    """Lee archivo Excel sin encabezados detectando el engine correcto"""
    nombre = archivo.name.lower()
    
    try:
        if nombre.endswith('.xls') and not nombre.endswith('.xlsx'):
            try:
                import xlrd
                # Intentar leer con xlrd
                return pd.read_excel(archivo, sheet_name=0, header=None, engine='xlrd')
            except Exception as e:
                # Si falla, intentar leer como HTML o texto
                print(f"[Reader Warning] Error leyendo como Excel, intentando como HTML: {str(e)}")
                
                archivo.seek(0)
                
                try:
                    tablas = pd.read_html(archivo)
                    
                    if len(tablas) > 0:
                        return tablas[0]
                    
                except Exception:
                    pass
                
                archivo.seek(0)
                
                contenido = archivo.read()
                try:
                    contenido = contenido.decode("utf-8")
                except UnicodeDecodeError:
                    archivo.seek(0)
                    contenido = archivo.read().decode("latin-1")
                
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

def mono_leer_excel_con_encabezados(archivo):
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
# 🔥 DETECCIÓN DE BANCO POR CONTENIDO DEL ARCHIVO (VERSIÓN MONOBANCO)
# =========================================================

def mono_detectar_banco_por_contenido(archivo):
    """
    Detecta el banco leyendo el contenido del archivo, no solo el nombre.
    Soporta formatos binarios de Excel, HTML o texto de forma robusta.
    """
    try:
        # Guardar la posición actual
        pos = archivo.tell()
        archivo.seek(0)
        
        # Leer usando la función robusta mono_leer_excel_sin_encabezados
        df_temp = mono_leer_excel_sin_encabezados(archivo)
        
        if df_temp is not None and not df_temp.empty:
            # Incluir encabezados/columnas de forma robusta (soporta MultiIndex)
            columnas_texto = ""
            if isinstance(df_temp.columns, pd.MultiIndex):
                for lvl in df_temp.columns.levels:
                    columnas_texto += " " + " ".join([str(x) for x in lvl if pd.notna(x)])
            else:
                columnas_texto = " ".join([str(x) for x in df_temp.columns if pd.notna(x)])
            
            # Convertir las primeras 40 filas a string para buscar
            df_sub = df_temp.head(40)
            texto_valores = " ".join([str(val) for val in df_sub.values.flatten() if pd.notna(val)])
            texto = (columnas_texto + " " + texto_valores).upper()
            
            # Restablecer la posición del archivo
            archivo.seek(pos)
            
            # Detectar por contenido - Banco Activo primero
            # 🔥 "0171" debe ser un número de cuenta de Banco Activo (20 dígitos aislados),
            # no cualquier coincidencia dentro de referencias de otros bancos
            es_cuenta_activo = bool(re.search(r'(?<![\d])0171\d{16}(?!\d)', texto))
            if "BANCO ACTIVO" in texto or es_cuenta_activo or "NUMERO DE CUENTA" in texto:
                # Verificar si tiene el formato de Banco Activo
                if df_temp.shape[1] >= 7:
                    for i in range(min(10, len(df_temp))):
                        fila = df_temp.iloc[i]
                        if pd.notna(fila[0]) and pd.notna(fila[2]) and pd.notna(fila[6]):
                            fecha_str = str(fila[0]).strip()
                            if re.match(r'\d{2}/\d{2}/\d{4}', fecha_str):
                                return "activo"
                return "activo"
            elif "BANCAMIGA" in texto or "BANCAMIGA BANCO UNIVERSAL" in texto or "BANCA AMIGA" in texto or "AMIGA" in texto:
                return "bancamiga"
            elif "BANESCO" in texto:
                return "banesco"
            elif "MERCANTIL" in texto or "105 VEB" in texto:
                return "mercantil"
            elif (
                "BANCO PROVINCIAL" in texto
                or "BBVA" in texto
                or "CÓDIGO DE OPERACIÓN" in texto
                or "CODIGO DE OPERACION" in texto
                or "PRIMERA ORDENACIÓN" in texto
                or "F. OPERACIÓN" in texto
                or "CUENTA ACTUAL" in texto
            ):
                return "provincial"
            elif "BANCO DE VENEZUELA" in texto or "BDV" in texto:
                return "venezuela"
            elif (
                "TIPO DE MOVIMIENTO" in texto
                or "TODAL DÉBITO" in texto
                or "TODAL DEBITO" in texto
                or "SALDO PROMEDIO" in texto
            ):
                # 🔥 Encabezado típico de los estados de cuenta del Banco de Venezuela
                return "venezuela"
            elif "BNC" in texto or "BANCO NACIONAL DE CREDITO" in texto:
                return "bnc"
            elif "BANPLUS" in texto or "BAN PLUS" in texto:
                return "banplus"
            elif "TESORO" in texto or "BANCO DEL TESORO" in texto:
                return "tesoro"
                
        # Restablecer si no se detectó nada
        archivo.seek(pos)
        return None
        
    except Exception as e:
        try:
            archivo.seek(pos)
        except Exception:
            pass
        return None

def mono_detectar_banco_por_nombre(nombre_archivo):
    """Detecta el banco por el nombre del archivo (fallback)"""
    nombre = nombre_archivo.upper()

    # Detectar por número de cuenta en el nombre (20 dígitos continuos o separados)
    clean_name = re.sub(r'[\s\-_]', '', nombre_archivo)
    match_banco = re.search(r'(0102|0105|0108|0134|0138|0163|0172|0174|0191)\d{16}', clean_name)
    if match_banco:
        codigo = match_banco.group(1)
        if codigo == "0102":
            return "venezuela"
        elif codigo == "0105":
            return "mercantil"
        elif codigo == "0108":
            return "provincial"
        elif codigo == "0134":
            return "banesco"
        elif codigo == "0138":
            return "activo"
        elif codigo == "0163":
            return "tesoro"
        elif codigo == "0172":
            return "bancamiga"
        elif codigo == "0174":
            return "banplus"
        elif codigo == "0191":
            return "bnc"

    if "ACTIVO" in nombre:
        return "activo"
    elif "MERCANTIL" in nombre:
        return "mercantil"
    elif "TESORO" in nombre or "TESORERIA" in nombre or "TES" in nombre:
        return "tesoro"
    elif "BANCAMIGA" in nombre or "BANCAAMIGA" in nombre or "AMIGA" in nombre:
        return "bancamiga"
    elif "BANPLUS" in nombre:
        return "banplus"
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
    elif "PROVINCIAL" in nombre or "BBVA" in nombre:
        return "provincial"
    elif "BNC" in nombre:
        return "bnc"
    return "mercantil"

# =========================================================
# FUNCIONES ORIGINALES (NO MODIFICADAS)
# =========================================================

def mono_convertir_monto(valor):
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

def mono_limpiar_monto_banplus(valor):
    if valor is None or pd.isna(valor):
        return 0.0
    if isinstance(valor, (int, float)):
        return float(valor)
    
    valor_str = str(valor).strip()
    if not valor_str:
        return 0.0
    
    valor_str = valor_str.replace('$', '').replace('Bs.', '').replace('Bs', '').replace(' ', '').strip()
    
    try:
        return float(valor_str)
    except ValueError:
        pass
    
    has_comma = ',' in valor_str
    has_dot = '.' in valor_str
    
    if has_comma and has_dot:
        pos_comma = valor_str.rfind(',')
        pos_dot = valor_str.rfind('.')
        if pos_dot > pos_comma:
            valor_limpio = valor_str.replace(',', '')
        else:
            valor_limpio = valor_str.replace('.', '').replace(',', '.')
    elif has_comma:
        valor_limpio = valor_str.replace(',', '.')
    elif has_dot:
        valor_limpio = valor_str
    else:
        valor_limpio = valor_str
        
    try:
        return float(valor_limpio)
    except ValueError:
        return 0.0

# =========================================================
# CALCULAR USD SEGÚN TASA
# =========================================================

def mono_calcular_usd(monto_bs, tasa):
    try:
        if monto_bs is None or tasa is None or tasa == 0:
            return None
        return round(abs(monto_bs) / abs(tasa), 2)
    except:
        return None

# =========================================================
# 🔥 DETECTAR COMISIONES - VERSIÓN MEJORADA CON PALABRAS CLAVE
# =========================================================

def mono_es_comision(texto, proveedor=None):
    """
    Versión monobanco de la función es_comision.
    Detecta si un movimiento es una comisión bancaria.
    """
    return es_comision(texto, proveedor)

# =========================================================
# 🔥 FUNCIÓN MEJORADA: ENRIQUECER EGRESOS CON IPAGO (CRUCE FLEXIBLE)
# =========================================================

def mono_enriquecer_egresos_con_ipago(df_egresos, df_ipago):
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

def mono_procesar_banesco(df):
    st.info("Procesando Banesco...")
    try:
        # 🔥 APLICAR FILTRO DE FECHA PREDOMINANTE
        df_filtrado, fecha_pred, dict_conteo, total_filas, registros_excluidos = filtrar_por_fecha_predominante(
            df, columna_fecha_idx=0, nombre_banco="Banesco"
        )
        
        if df_filtrado is None or df_filtrado.empty:
            df_filtrado = df
            st.warning("⚠️ Banesco: No se pudo filtrar por fecha predominante, se procesará todo el archivo.")
        
        if fecha_pred:
            st.session_state.info_fechas_por_banco["Banesco"] = {
                'fecha': fecha_pred.strftime('%d/%m/%Y'),
                'registros': len(df_filtrado),
                'total_original': total_filas,
                'excluidos': registros_excluidos,
                'detalle_fechas': dict_conteo
            }
        
        df_filtrado.columns = ["FECHA", "REFERENCIA", "DESCRIPCION", "MONTO_RAW", "BALANCE"]
        df_filtrado.columns = [str(c).strip().upper() for c in df_filtrado.columns]

        rename_map = {}
        for col in df_filtrado.columns:
            c = str(col).lower()
            if "fecha" in c:
                rename_map[col] = "FECHA"
            elif "referencia" in c:
                rename_map[col] = "REFERENCIA"
            elif "descrip" in c:
                rename_map[col] = "DESCRIPCION"
            elif "monto" in c:
                rename_map[col] = "MONTO_RAW"

        df_filtrado = df_filtrado.rename(columns=rename_map)

        for col in ["FECHA", "REFERENCIA", "DESCRIPCION", "MONTO_RAW"]:
            if col not in df_filtrado.columns:
                st.error(f"No existe columna: {col}")
                return pd.DataFrame()

        # Convertir fechas de manera robusta
        def parse_banesco_date(val):
            val_str = str(val).strip()
            if not val_str or val_str == "nan":
                return pd.NaT
            if len(val_str) >= 5 and val_str[:4].isdigit() and val_str[4] in ('/', '-'):
                return pd.to_datetime(val_str, dayfirst=False, errors="coerce")
            return pd.to_datetime(val_str, dayfirst=True, errors="coerce")

        df_filtrado["FECHA"] = df_filtrado["FECHA"].apply(parse_banesco_date)
        df_filtrado = df_filtrado[df_filtrado["FECHA"].notna()]

        # 🔥 PARSING DE MONTO ROBUSTO: soporta positivos sin "+", punto decimal
        # (345570.17), coma decimal venezolana (345.570,17) y enteros (316000)
        def limpiar_monto_banesco(valor):
            if valor is None or (isinstance(valor, float) and pd.isna(valor)):
                return None
            if isinstance(valor, (int, float, np.integer, np.floating)):
                return float(valor)
            val_str = str(valor).strip().replace(" ", "")
            if not val_str or val_str.lower() == "nan":
                return None
            es_negativo = val_str.startswith("-")
            val_str = val_str.lstrip("+-")
            if "," in val_str and "." in val_str:
                if val_str.rfind(",") > val_str.rfind("."):
                    val_str = val_str.replace(".", "").replace(",", ".")
                else:
                    val_str = val_str.replace(",", "")
            elif "," in val_str:
                val_str = val_str.replace(",", ".")
            try:
                numero = float(val_str)
            except ValueError:
                return None
            return -abs(numero) if es_negativo else numero
        
        df_filtrado["MONTO_LIMPIO"] = df_filtrado["MONTO_RAW"].apply(limpiar_monto_banesco)
        df_filtrado["TIPO"] = df_filtrado["MONTO_LIMPIO"].apply(
            lambda x: "NC" if x is not None and x > 0 else "ND" if x is not None else "ND"
        )
        df_filtrado["MONTO"] = df_filtrado["MONTO_LIMPIO"].abs()
        df_filtrado = df_filtrado[df_filtrado["MONTO"].notna()]
        df_filtrado = df_filtrado[df_filtrado["MONTO"] > 0]

        df_filtrado = df_filtrado[["FECHA", "REFERENCIA", "DESCRIPCION", "TIPO", "MONTO"]]
        st.success(f"Banesco OK: {len(df_filtrado)} movimientos")
        st.dataframe(df_filtrado.head())
        return df_filtrado

    except Exception as e:
        st.error(f"Error Banesco: {str(e)}")
        return pd.DataFrame()

# =========================================================
# PROCESAR PROVINCIAL - VERSIÓN MEJORADA PARA FORMATO ESPECÍFICO
# =========================================================

def mono_procesar_provincial(df):
    """
    Procesa archivo de Provincial con formato específico.
    El archivo tiene un formato de texto con columnas:
    F. Operación | F. Valor | Código | Nº. Doc. | Concepto | Importe | Oficina
    """
    st.info("🔍 Procesando archivo de Provincial (formato especial)...")
    
    try:
        # 🔥 1. BUSCAR EL ENCABEZADO EN EL DF ORIGINAL (antes de filtrar por fecha)
        encabezado_idx = None
        for i in range(min(30, len(df))):
            fila = df.iloc[i]
            fila_str = [str(val) for val in fila.tolist()]
            texto_fila = " ".join(fila_str).upper()
            if "CONCEPTO" in texto_fila and "IMPORTE" in texto_fila:
                encabezado_idx = i
                break
        
        if encabezado_idx is None:
            st.error("❌ No se encontró la fila de encabezados en el archivo Provincial.")
            return pd.DataFrame()
        
        fila_encabezado = df.iloc[[encabezado_idx]].copy()
        df_datos = df.iloc[encabezado_idx + 1:].copy()
        
        # 🔥 2. APLICAR FILTRO DE FECHA PREDOMINANTE SOLO A LAS FILAS DE DATOS (post-encabezado)
        df_filtrado, fecha_pred, dict_conteo, total_filas, registros_excluidos = filtrar_por_fecha_predominante(
            df_datos, columna_fecha_idx=0, nombre_banco="Provincial"
        )
        
        if df_filtrado is None or df_filtrado.empty:
            df_filtrado = df_datos
            st.warning("⚠️ Provincial: No se pudo filtrar por fecha predominante, se procesará todo el archivo.")
        
        if fecha_pred:
            st.session_state.info_fechas_por_banco["Provincial"] = {
                'fecha': fecha_pred.strftime('%d/%m/%Y'),
                'registros': len(df_filtrado),
                'total_original': total_filas,
                'excluidos': registros_excluidos,
                'detalle_fechas': dict_conteo
            }
        
        # 🔥 3. REENSAMBLAR: encabezado + datos filtrados
        df_filtrado = pd.concat([fila_encabezado, df_filtrado], ignore_index=True)
        encabezado_idx = 0
        
        # Mostrar información del archivo
        st.write("📊 **Información del archivo:**")
        st.write(f"- Número de filas: {len(df_filtrado)}")
        st.write(f"- Número de columnas: {len(df_filtrado.columns)}")
        
        # Mostrar primeras filas para debug
        st.write("👁️ **Primeras 15 filas del archivo:**")
        st.dataframe(df_filtrado.head(15))
        
        st.write(f"✅ Encabezados encontrados en la fila {encabezado_idx}")
        
        # Obtener los encabezados
        headers = df_filtrado.iloc[encabezado_idx].astype(str).str.strip().tolist()
        st.write("📋 **Encabezados detectados:**", headers)
        
        # Limpiar y mapear encabezados
        rename_map = {}
        for col in headers:
            col_clean = str(col).strip().upper()
            if "OPERAC" in col_clean or "FECHA" in col_clean:
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
        df_filtrado.columns = headers
        df_filtrado = df_filtrado.iloc[encabezado_idx + 1:].reset_index(drop=True)
        
        # Renombrar columnas
        df_filtrado = df_filtrado.rename(columns=rename_map)
        
        # Verificar columnas necesarias
        if "FECHA" not in df_filtrado.columns:
            # Intentar encontrar fecha en otra columna
            for col in df_filtrado.columns:
                if "FECHA" in str(col).upper():
                    df_filtrado = df_filtrado.rename(columns={col: "FECHA"})
                    break
        
        if "FECHA" in df_filtrado.columns:
            # Procesar fechas - convertir a string primero
            df_filtrado["FECHA"] = df_filtrado["FECHA"].astype(str).str.strip()
            # Eliminar filas con fechas vacías o que sean encabezados
            df_filtrado = df_filtrado[~df_filtrado["FECHA"].str.contains("FECHA|SALDO|Período", case=False, na=False)]
            # Convertir fechas de manera robusta (admite dd/mm/yyyy y datetime de Excel)
            df_filtrado["FECHA_DT"] = pd.to_datetime(df_filtrado["FECHA"], dayfirst=True, errors="coerce")
            mask = df_filtrado["FECHA_DT"].isna()
            if mask.any():
                df_filtrado.loc[mask, "FECHA_DT"] = pd.to_datetime(df_filtrado.loc[mask, "FECHA"].astype(str).str.strip(), dayfirst=True, errors="coerce")
            df_filtrado = df_filtrado[df_filtrado["FECHA_DT"].notna()]
            df_filtrado["FECHA"] = df_filtrado["FECHA_DT"].dt.strftime("%d/%m/%Y")
            df_filtrado = df_filtrado.drop(columns=["FECHA_DT"])
        else:
            st.error("❌ No se encontró columna FECHA en el archivo Provincial.")
            return pd.DataFrame()
        
        # Procesar el monto
        if "MONTO" in df_filtrado.columns:
            # Limpiar el monto (quitar espacios, puntos, comas) - convertir a string primero
            df_filtrado["MONTO"] = df_filtrado["MONTO"].astype(str).str.replace(" ", "", regex=False)
            df_filtrado["MONTO"] = df_filtrado["MONTO"].str.replace(".", "", regex=False)
            df_filtrado["MONTO"] = df_filtrado["MONTO"].str.replace(",", ".", regex=False)
            df_filtrado["MONTO"] = df_filtrado["MONTO"].str.replace("'", "", regex=False)
            
            # Convertir directamente a numérico (sin filtro regex)
            df_filtrado["MONTO"] = pd.to_numeric(df_filtrado["MONTO"], errors="coerce")
            
            # Eliminar filas con monto NaN
            df_filtrado = df_filtrado[df_filtrado["MONTO"].notna()]
            
            # Si el monto es negativo, es un ND (débito), si es positivo es NC (crédito)
            df_filtrado["TIPO"] = df_filtrado["MONTO"].apply(lambda x: "NC" if x > 0 else "ND" if x < 0 else "")
            
            # Tomar valor absoluto
            df_filtrado["MONTO"] = df_filtrado["MONTO"].abs()
            
            # Eliminar filas con monto 0
            df_filtrado = df_filtrado[df_filtrado["MONTO"] > 0]
        else:
            st.error("❌ No se encontró columna MONTO en el archivo Provincial.")
            return pd.DataFrame()
        
        # Asegurar que existe columna REFERENCIA
        if "REFERENCIA" not in df_filtrado.columns:
            df_filtrado["REFERENCIA"] = ""
        else:
            df_filtrado["REFERENCIA"] = df_filtrado["REFERENCIA"].astype(str).str.strip()
            # Limpiar referencias (quitar comillas simples)
            df_filtrado["REFERENCIA"] = df_filtrado["REFERENCIA"].str.replace("'", "", regex=False)
        
        # Asegurar que existe columna DESCRIPCION
        if "DESCRIPCION" not in df_filtrado.columns:
            df_filtrado["DESCRIPCION"] = ""
        else:
            df_filtrado["DESCRIPCION"] = df_filtrado["DESCRIPCION"].astype(str).str.strip()
        
        # 🔥 DETECTAR COMISIONES DE PROVINCIAL
        df_filtrado["ES_COMISION"] = df_filtrado["DESCRIPCION"].str.contains("COMIS", case=False, na=False)
        
        # 🔥 DEBUG: Mostrar cuántas comisiones se detectaron
        num_comisiones = df_filtrado["ES_COMISION"].sum()
        st.info(f"💳 Se detectaron {num_comisiones} comisiones en el archivo Provincial")
        
        # Mostrar las comisiones detectadas
        if num_comisiones > 0:
            st.write("📋 **Comisiones detectadas:**")
            st.dataframe(df_filtrado[df_filtrado["ES_COMISION"] == True][["FECHA", "REFERENCIA", "DESCRIPCION", "MONTO"]])
        
        # Seleccionar solo las columnas necesarias
        df_resultado = df_filtrado[["FECHA", "REFERENCIA", "DESCRIPCION", "TIPO", "MONTO", "ES_COMISION"]].copy()
        
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

def mono_procesar_bnc(df):
    st.info("Procesando archivo BNC...")
    
    # 🔥 1. BUSCAR EL ENCABEZADO EN EL DF ORIGINAL (antes de filtrar por fecha)
    encabezado = None
    for i in range(min(30, len(df))):
        fila = df.iloc[i].fillna("").astype(str)
        texto = " ".join(fila.tolist()).lower()
        if "fecha" in texto and ("descripcion" in texto or "descripción" in texto):
            encabezado = i
            break

    if encabezado is not None:
        fila_encabezado = df.iloc[[encabezado]].copy()
        df_datos = df.iloc[encabezado + 1:].copy()
    else:
        df_datos = df.copy()
        st.warning("⚠️ BNC: No se encontró la fila de encabezados, se procesará todo el archivo.")
    
    # 🔥 2. APLICAR FILTRO DE FECHA PREDOMINANTE SOLO A LAS FILAS DE DATOS
    df_filtrado, fecha_pred, dict_conteo, total_filas, registros_excluidos = filtrar_por_fecha_predominante(
        df_datos, columna_fecha_idx=0, nombre_banco="BNC"
    )
    
    if df_filtrado is None or df_filtrado.empty:
        df_filtrado = df_datos
        st.warning("⚠️ BNC: No se pudo filtrar por fecha predominante, se procesará todo el archivo.")
    
    if fecha_pred:
        st.session_state.info_fechas_por_banco["BNC"] = {
            'fecha': fecha_pred.strftime('%d/%m/%Y'),
            'registros': len(df_filtrado),
            'total_original': total_filas,
            'excluidos': registros_excluidos,
            'detalle_fechas': dict_conteo
        }
    
    # 🔥 3. REENSAMBLAR: encabezado + datos filtrados
    if encabezado is not None:
        df_filtrado = pd.concat([fila_encabezado, df_filtrado], ignore_index=True)
        encabezado = 0
    else:
        encabezado = None
        for i in range(min(30, len(df_filtrado))):
            fila = df_filtrado.iloc[i].fillna("").astype(str)
            texto = " ".join(fila.tolist()).lower()
            if "fecha" in texto and ("descripcion" in texto or "descripción" in texto):
                encabezado = i
                break

    if encabezado is None:
        st.error("No se encontró encabezado válido en BNC")
        return pd.DataFrame()

    headers = []
    for idx, col in enumerate(df_filtrado.iloc[encabezado]):
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

    df_filtrado.columns = headers_unicos
    df_filtrado = df_filtrado.iloc[encabezado + 1:].reset_index(drop=True)

    rename_map = {}
    for col in df_filtrado.columns:
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

    df_filtrado = df_filtrado.rename(columns=rename_map)

    if "FECHA" in df_filtrado.columns:
        df_filtrado["FECHA"] = pd.to_datetime(df_filtrado["FECHA"], dayfirst=True, errors="coerce")
        df_filtrado = df_filtrado[df_filtrado["FECHA"].notna()]

    df_filtrado["CREDITO"] = pd.to_numeric(df_filtrado.get("CREDITO", 0), errors="coerce").fillna(0) if "CREDITO" in df_filtrado.columns else 0
    df_filtrado["DEBITO"] = pd.to_numeric(df_filtrado.get("DEBITO", 0), errors="coerce").fillna(0) if "DEBITO" in df_filtrado.columns else 0

    df_filtrado["MONTO"] = df_filtrado["CREDITO"] - df_filtrado["DEBITO"]
    df_filtrado["TIPO"] = df_filtrado["MONTO"].apply(lambda x: "NC" if x > 0 else "ND")
    df_filtrado["MONTO"] = df_filtrado["MONTO"].abs()
    df_filtrado = df_filtrado[df_filtrado["MONTO"] != 0]

    st.success(f"Registros BNC: {len(df_filtrado)}")
    st.dataframe(df_filtrado.head())
    return df_filtrado

# =========================================================
# PROCESAR TESORO
# =========================================================

def mono_procesar_tesoro(df):
    st.info("Procesando Banco del Tesoro...")
    
    try:
        # 🔥 1. BUSCAR EL ENCABEZADO EN EL DF ORIGINAL (antes de filtrar por fecha)
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
        fila_encabezado = df.iloc[[encabezado]].copy()
        df_datos = df.iloc[encabezado + 1:].copy()
    except Exception:
        fila_encabezado = df.iloc[[0]].copy()
        df_datos = df.iloc[1:].copy()
    
    # 🔥 2. APLICAR FILTRO DE FECHA PREDOMINANTE SOLO A LAS FILAS DE DATOS (fechas en columna 1)
    df_filtrado, fecha_pred, dict_conteo, total_filas, registros_excluidos = filtrar_por_fecha_predominante(
        df_datos, columna_fecha_idx=1, nombre_banco="Tesoro"
    )
    
    if df_filtrado is None or df_filtrado.empty:
        df_filtrado = df_datos
        st.warning("⚠️ Tesoro: No se pudo filtrar por fecha predominante, se procesará todo el archivo.")
    
    if fecha_pred:
        st.session_state.info_fechas_por_banco["Tesoro"] = {
            'fecha': fecha_pred.strftime('%d/%m/%Y'),
            'registros': len(df_filtrado),
            'total_original': total_filas,
            'excluidos': registros_excluidos,
            'detalle_fechas': dict_conteo
        }
    
    # 🔥 3. REENSAMBLAR: encabezado + datos filtrados
    df_filtrado = pd.concat([fila_encabezado, df_filtrado], ignore_index=True)
    
    try:
        encabezado = 0
        df_filtrado.columns = df_filtrado.iloc[encabezado]
        df_filtrado = df_filtrado.iloc[encabezado + 1:].reset_index(drop=True)
        df_filtrado.columns = [str(c).strip() for c in df_filtrado.columns]

        rename_map = {}
        for col in df_filtrado.columns:
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

        df_filtrado = df_filtrado.rename(columns=rename_map)

        if "FECHA" not in df_filtrado.columns:
            st.error("No existe columna FECHA")
            return pd.DataFrame()

        df_filtrado["FECHA"] = pd.to_datetime(df_filtrado["FECHA"], dayfirst=True, errors="coerce")
        df_filtrado = df_filtrado[df_filtrado["FECHA"].notna()]

        def limpiar_numero(valor):
            valor = str(valor).replace(".", "").replace(",", ".")
            try:
                return float(valor)
            except:
                return 0

        df_filtrado["CREDITO"] = df_filtrado.get("CREDITO", 0).apply(limpiar_numero) if "CREDITO" in df_filtrado.columns else 0
        df_filtrado["DEBITO"] = df_filtrado.get("DEBITO", 0).apply(limpiar_numero) if "DEBITO" in df_filtrado.columns else 0

        df_filtrado["MONTO"] = df_filtrado["CREDITO"] - df_filtrado["DEBITO"]
        df_filtrado["TIPO"] = df_filtrado["MONTO"].apply(lambda x: "NC" if x > 0 else "ND")
        df_filtrado["MONTO"] = df_filtrado["MONTO"].abs()
        df_filtrado = df_filtrado[df_filtrado["MONTO"] > 0]

        df_filtrado = df_filtrado[["FECHA", "REFERENCIA", "DESCRIPCION", "TIPO", "MONTO"]]
        st.success(f"Tesoro OK: {len(df_filtrado)} registros")
        st.dataframe(df_filtrado.head())
        return df_filtrado

    except Exception as e:
        st.error(f"Error Tesoro: {str(e)}")
        return pd.DataFrame()

# =========================================================
# PROCESAR BANCAMIGA - VERSIÓN MEJORADA CON DETECCIÓN DE ENCABEZADOS
# =========================================================

def mono_procesar_bancamiga(df):
    """
    Procesa archivo de Bancamiga con formato específico.
    El archivo tiene columnas: Nro. | Fecha | Referencia | Concepto | Débito | Crédito | Saldo
    """
    st.info("🔍 Procesando archivo de Bancamiga...")
    
    try:
        if isinstance(df.columns, pd.MultiIndex):
            df.columns = df.columns.get_level_values(-1)
        
        columnas_orig = [str(c).strip().upper() for c in df.columns]
        tiene_encabezados = "FECHA" in columnas_orig and "REFERENCIA" in columnas_orig
        
        # 🔥 1. BUSCAR EL ENCABEZADO EN EL DF ORIGINAL (antes de filtrar por fecha)
        encabezado_idx = None
        if not tiene_encabezados:
            for i in range(min(30, len(df))):
                fila = df.iloc[i]
                fila_str = [str(val) for val in fila.tolist()]
                texto_fila = " ".join(fila_str).upper()
                if "NRO" in texto_fila and "FECHA" in texto_fila and "REFERENCIA" in texto_fila:
                    encabezado_idx = i
                    break
        
        if tiene_encabezados:
            df_datos = df.copy()
        elif encabezado_idx is not None:
            fila_encabezado = df.iloc[[encabezado_idx]].copy()
            df_datos = df.iloc[encabezado_idx + 1:].copy()
        else:
            df_datos = df.copy()
            st.warning("⚠️ Bancamiga: No se encontró la fila de encabezados, se procesará todo el archivo.")
        
        # 🔥 2. APLICAR FILTRO DE FECHA PREDOMINANTE SOLO A LAS FILAS DE DATOS
        df_filtrado, fecha_pred, dict_conteo, total_filas, registros_excluidos = filtrar_por_fecha_predominante(
            df_datos, columna_fecha_idx=1, nombre_banco="Bancamiga"
        )
        
        if df_filtrado is None or df_filtrado.empty:
            df_filtrado = df_datos
            st.warning("⚠️ Bancamiga: No se pudo filtrar por fecha predominante, se procesará todo el archivo.")
        
        if fecha_pred:
            st.session_state.info_fechas_por_banco["Bancamiga"] = {
                'fecha': fecha_pred.strftime('%d/%m/%Y'),
                'registros': len(df_filtrado),
                'total_original': total_filas,
                'excluidos': registros_excluidos,
                'detalle_fechas': dict_conteo
            }
        
        # 🔥 3. REENSAMBLAR: encabezado + datos filtrados
        if not tiene_encabezados and encabezado_idx is not None:
            df_filtrado = pd.concat([fila_encabezado, df_filtrado], ignore_index=True)
        
        # Mostrar información del archivo
        st.write("📊 **Información del archivo:**")
        st.write(f"- Número de filas: {len(df_filtrado)}")
        st.write(f"- Número de columnas: {len(df_filtrado.columns)}")
        
        # Mostrar primeras filas para debug
        st.write("👁️ **Primeras 15 filas del archivo:**")
        st.dataframe(df_filtrado.head(15))
        
        # 🔥 VERIFICAR SI YA TIENE ENCABEZADOS
        columnas = [str(c).strip().upper() for c in df_filtrado.columns]
        
        if "FECHA" in columnas and "REFERENCIA" in columnas:
            st.success("✅ Encabezados ya presentes.")
            
            # Mapeo de columnas
            rename_map = {
                "NRO.": "NRO",
                "NRO": "NRO",
                "FECHA": "FECHA",
                "REFERENCIA": "REFERENCIA",
                "CONCEPTO": "DESCRIPCION",
                "DÉBITO": "DEBITO",
                "DEBITO": "DEBITO",
                "CRÉDITO": "CREDITO",
                "CREDITO": "CREDITO",
                "SALDO": "SALDO"
            }
            
            df_filtrado.columns = [rename_map.get(str(c).strip().upper(), str(c).strip().upper())
                          for c in df_filtrado.columns]
            
        else:
            st.info("🔍 Buscando fila de encabezados...")
            
            # Buscar la fila que contiene los encabezados
            encabezado_idx = None
            for i in range(min(30, len(df_filtrado))):
                fila = df_filtrado.iloc[i]
                # Convertir TODOS los valores a string para evitar errores
                fila_str = [str(val) for val in fila.tolist()]
                texto_fila = " ".join(fila_str).upper()
                
                # Buscar columnas que contengan "NRO" o "FECHA" o "REFERENCIA" o "CONCEPTO"
                if "NRO" in texto_fila and "FECHA" in texto_fila and "REFERENCIA" in texto_fila:
                    encabezado_idx = i
                    break
            
            if encabezado_idx is None:
                st.error("❌ No se encontró la fila de encabezados en el archivo Bancamiga.")
                return pd.DataFrame()
            
            st.write(f"✅ Encabezados encontrados en la fila {encabezado_idx}")
            
            # Obtener los encabezados
            headers = df_filtrado.iloc[encabezado_idx].astype(str).str.strip().tolist()
            st.write("📋 **Encabezados detectados:**", headers)
            
            # Limpiar y mapear encabezados
            rename_map = {}
            for col in headers:
                col_clean = str(col).strip().upper()
                if "NRO" in col_clean or "Nº" in col_clean:
                    rename_map[col] = "NRO"
                elif "FECHA" in col_clean:
                    rename_map[col] = "FECHA"
                elif "REFERENCIA" in col_clean:
                    rename_map[col] = "REFERENCIA"
                elif "CONCEPTO" in col_clean:
                    rename_map[col] = "DESCRIPCION"
                elif "DÉBITO" in col_clean or "DEBITO" in col_clean:
                    rename_map[col] = "DEBITO"
                elif "CRÉDITO" in col_clean or "CREDITO" in col_clean:
                    rename_map[col] = "CREDITO"
                elif "SALDO" in col_clean:
                    rename_map[col] = "SALDO"
            
            st.write("📋 **Mapeo de columnas:**", rename_map)
            
            # Asignar encabezados al DataFrame
            df_filtrado.columns = headers
            df_filtrado = df_filtrado.iloc[encabezado_idx + 1:].reset_index(drop=True)
            
            # Renombrar columnas
            df_filtrado = df_filtrado.rename(columns=rename_map)
        
        # Verificar columnas necesarias
        if "FECHA" not in df_filtrado.columns:
            st.error("❌ No se encontró columna FECHA en el archivo Bancamiga.")
            return pd.DataFrame()
        
        # 🔥 PROCESAR FECHAS DE BANCAMIGA DE MANERA ROBUSTA
        # Filtrar filas que no son movimientos
        fechas_str_col = df_filtrado["FECHA"].astype(str).str.strip()
        df_filtrado = df_filtrado[
            ~fechas_str_col.str.contains(
                "FECHA|SALDO|TOTAL|CRÉDITO|CREDITO|DÉBITO|DEBITO",
                case=False,
                na=False
            )
        ]

        # Convertir fechas de manera robusta
        df_filtrado["FECHA_DT"] = pd.to_datetime(df_filtrado["FECHA"], dayfirst=True, errors="coerce")
        mask = df_filtrado["FECHA_DT"].isna()
        if mask.any():
            df_filtrado.loc[mask, "FECHA_DT"] = pd.to_datetime(
                df_filtrado.loc[mask, "FECHA"].astype(str).str.strip(),
                dayfirst=True,
                errors="coerce"
            )
        
        df_filtrado = df_filtrado[df_filtrado["FECHA"].notna()]
        
        # Procesar débito y crédito
        def limpiar_monto(val):
            val_str = str(val).strip().replace(" ", "")
            if not val_str or val_str == "nan":
                return 0.0
            if "," in val_str:
                val_str = val_str.replace(".", "").replace(",", ".")
            return pd.to_numeric(val_str, errors="coerce")

        if "DEBITO" in df_filtrado.columns:
            df_filtrado["DEBITO"] = df_filtrado["DEBITO"].apply(limpiar_monto).fillna(0)
        else:
            df_filtrado["DEBITO"] = 0.0
        
        if "CREDITO" in df_filtrado.columns:
            df_filtrado["CREDITO"] = df_filtrado["CREDITO"].apply(limpiar_monto).fillna(0)
        else:
            df_filtrado["CREDITO"] = 0.0
        
        # Determinar tipo y monto
        df_filtrado["MONTO"] = df_filtrado["CREDITO"] - df_filtrado["DEBITO"]
        df_filtrado["TIPO"] = df_filtrado["MONTO"].apply(lambda x: "NC" if x > 0 else "ND" if x < 0 else "")
        df_filtrado["MONTO"] = df_filtrado["MONTO"].abs()
        
        # Eliminar filas con monto 0
        df_filtrado = df_filtrado[df_filtrado["MONTO"] > 0]
        
        # Asegurar que existe columna REFERENCIA
        if "REFERENCIA" not in df_filtrado.columns:
            df_filtrado["REFERENCIA"] = ""
        else:
            df_filtrado["REFERENCIA"] = df_filtrado["REFERENCIA"].astype(str).str.strip()
            # Limpiar referencias (quitar comillas simples)
            df_filtrado["REFERENCIA"] = df_filtrado["REFERENCIA"].str.replace("'", "", regex=False)
        
        # Asegurar que existe columna DESCRIPCION
        if "DESCRIPCION" not in df_filtrado.columns:
            df_filtrado["DESCRIPCION"] = ""
        else:
            df_filtrado["DESCRIPCION"] = df_filtrado["DESCRIPCION"].astype(str).str.strip()
        
        # 🔥 DETECTAR COMISIONES DE BANCAMIGA
        # Las comisiones tienen "Comisión" en la descripción
        df_filtrado["ES_COMISION"] = df_filtrado["DESCRIPCION"].str.contains("Comisi", case=False, na=False)
        
        # 🔥 DEBUG: Mostrar cuántas comisiones se detectaron
        num_comisiones = df_filtrado["ES_COMISION"].sum()
        st.info(f"💳 Se detectaron {num_comisiones} comisiones en el archivo Bancamiga")
        
        # Mostrar las comisiones detectadas
        if num_comisiones > 0:
            st.write("📋 **Comisiones detectadas:**")
            st.dataframe(df_filtrado[df_filtrado["ES_COMISION"] == True][["FECHA", "REFERENCIA", "DESCRIPCION", "MONTO"]])
        
        # Seleccionar solo las columnas necesarias
        df_resultado = df_filtrado[["FECHA", "REFERENCIA", "DESCRIPCION", "TIPO", "MONTO", "ES_COMISION"]].copy()
        
        # Mostrar resultados
        st.success(f"✅ Bancamiga OK: {len(df_resultado)} movimientos detectados")
        st.dataframe(df_resultado.head(10))
        
        return df_resultado
        
    except Exception as e:
        st.error(f"❌ Error procesando Bancamiga: {str(e)}")
        import traceback
        st.code(traceback.format_exc())
        return pd.DataFrame()

# =========================================================
# OBTENER TASA BCV
# =========================================================

def mono_obtener_tasa_por_fecha(fecha_obj, usar_api=False):
    # ✅ MISMO SISTEMA QUE MULTIBANCO: histórico local + registro automático + API BCV según la fecha
    tasa, _ = _tasa_con_origen_por_fecha(fecha_obj, usar_api)
    return tasa

# =========================================================
# CONVERTIR A FORMATO MERCANTIL - INCLUYE FLAG DE COMISIONES
# =========================================================

def mono_convertir_a_formato_mercantil(df, banco):
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

def mono_procesar_venezuela_simple(df):
    """Procesa el archivo del BDV usando índices fijos (sin encabezados)"""
    st.info("🔍 Procesando Banco de Venezuela (MODO SIMPLE)...")
    
    try:
        # 🔥 APLICAR FILTRO DE FECHA PREDOMINANTE
        df_filtrado, fecha_pred, dict_conteo, total_filas, registros_excluidos = filtrar_por_fecha_predominante(
            df, columna_fecha_idx=3, nombre_banco="Banco de Venezuela"
        )
        
        if df_filtrado is None or df_filtrado.empty:
            df_filtrado = df
            st.warning("⚠️ Banco de Venezuela: No se pudo filtrar por fecha predominante, se procesará todo el archivo.")
        
        if fecha_pred:
            st.session_state.info_fechas_por_banco["Banco de Venezuela"] = {
                'fecha': fecha_pred.strftime('%d/%m/%Y'),
                'registros': len(df_filtrado),
                'total_original': total_filas,
                'excluidos': registros_excluidos,
                'detalle_fechas': dict_conteo
            }
        
        # Mostrar información del archivo
        st.write("📊 **Información del archivo:**")
        st.write(f"- Número de filas: {len(df_filtrado)}")
        st.write(f"- Número de columnas: {len(df_filtrado.columns)}")
        
        # Mostrar primeras filas
        st.write("👁️ **Primeras 10 filas del archivo (sin encabezados):**")
        st.dataframe(df_filtrado.head(10))
        
        # Índices fijos para el formato BDV
        col_fecha = 3
        col_ref = 1
        col_desc = 2
        col_tipo = 4
        col_credito = 5
        col_debito = 6
        
        movimientos = []
        filas_procesadas = 0
        
        # 🔥 CORREGIDO: Empezar desde la fila 0 porque filtrar_por_fecha_predominante
        # ya eliminó los encabezados. Empezar en 1 perdía siempre el primer movimiento.
        for idx in range(0, len(df_filtrado)):
            try:
                fila = df_filtrado.iloc[idx]
                
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

def mono_convertir_venezuela_a_formato_mercantil(df):
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

def mono_procesar_banplus(df):
    """
    Procesa archivo de Banplus con formato HTML/Excel.
    Columnas: Fecha | Referencia | Descripción | Débito | Crédito | Saldo
    """
    st.info("🔍 Procesando archivo de Banplus...")
    
    # 🔥 1. BUSCAR EL ENCABEZADO EN EL DF ORIGINAL (antes de filtrar por fecha)
    encabezado_idx = None
    cols_upper = [str(c).strip().upper() for c in df.columns]
    if not ("FECHA" in cols_upper and "REFERENCIA" in cols_upper):
        for i in range(min(15, len(df))):
            fila = df.iloc[i].astype(str).str.strip().str.upper().tolist()
            fila_str = " ".join(fila)
            if "FECHA" in fila_str and "REFERENCIA" in fila_str:
                encabezado_idx = i
                break
    
    if encabezado_idx is not None:
        fila_encabezado = df.iloc[[encabezado_idx]].copy()
        df_datos = df.iloc[encabezado_idx + 1:].copy()
    else:
        df_datos = df.copy()
        st.warning("⚠️ Banplus: No se encontró la fila de encabezados, se procesará todo el archivo.")
    
    # 🔥 2. APLICAR FILTRO DE FECHA PREDOMINANTE SOLO A LAS FILAS DE DATOS
    df_filtrado, fecha_pred, dict_conteo, total_filas, registros_excluidos = filtrar_por_fecha_predominante(
        df_datos, columna_fecha_idx=0, nombre_banco="Banplus"
    )
    
    if df_filtrado is None or df_filtrado.empty:
        df_filtrado = df_datos
        st.warning("⚠️ Banplus: No se pudo filtrar por fecha predominante, se procesará todo el archivo.")
    
    if fecha_pred:
        st.session_state.info_fechas_por_banco["Banplus"] = {
            'fecha': fecha_pred.strftime('%d/%m/%Y'),
            'registros': len(df_filtrado),
            'total_original': total_filas,
            'excluidos': registros_excluidos,
            'detalle_fechas': dict_conteo
        }
    
    # 🔥 3. REENSAMBLAR: encabezado + datos filtrados
    if encabezado_idx is not None:
        df_filtrado = pd.concat([fila_encabezado, df_filtrado], ignore_index=True)
    
    try:
        # Mostrar información del archivo
        st.write("📊 **Información del archivo:**")
        st.write(f"- Número de filas: {len(df_filtrado)}")
        st.write(f"- Número de columnas: {len(df_filtrado.columns)}")
        
        # Encontrar la fila de encabezados si no está en las columnas
        encabezado_idx = None
        cols_upper = [str(c).strip().upper() for c in df_filtrado.columns]
        if "FECHA" in cols_upper and "REFERENCIA" in cols_upper:
            pass
        else:
            for i in range(min(15, len(df_filtrado))):
                fila = df_filtrado.iloc[i].astype(str).str.strip().str.upper().tolist()
                fila_str = " ".join(fila)
                if "FECHA" in fila_str and "REFERENCIA" in fila_str:
                    encabezado_idx = i
                    break
            if encabezado_idx is not None:
                df_filtrado.columns = df_filtrado.iloc[encabezado_idx].tolist()
                df_filtrado = df_filtrado.iloc[encabezado_idx + 1:].reset_index(drop=True)
        
        # Limpiar columnas
        df_filtrado.columns = [str(c).strip().upper() for c in df_filtrado.columns]
        
        rename_map = {}
        for col in df_filtrado.columns:
            col_clean = str(col).strip().upper()
            if "FECHA" in col_clean: rename_map[col] = "FECHA"
            elif "REFERENCIA" in col_clean: rename_map[col] = "REFERENCIA"
            elif "DESCRIP" in col_clean: rename_map[col] = "DESCRIPCION"
            elif "DÉBITO" in col_clean or "DEBITO" in col_clean or "DEB" in col_clean: rename_map[col] = "DEBITO"
            elif "CRÉDITO" in col_clean or "CREDITO" in col_clean or "CRE" in col_clean: rename_map[col] = "CREDITO"
            elif "SALDO" in col_clean: rename_map[col] = "SALDO"
            
        df_filtrado = df_filtrado.rename(columns=rename_map)
        
        # Filtrar filas vacías o totales
        if "FECHA" in df_filtrado.columns:
            df_filtrado["FECHA"] = df_filtrado["FECHA"].astype(str).str.strip()
            df_filtrado = df_filtrado[~df_filtrado["FECHA"].str.contains("FECHA|SALDO|Período|Total", case=False, na=False)]
            # Convertir fechas de manera robusta (admite dd/mm/yyyy y datetime de Excel)
            df_filtrado["FECHA_DT"] = pd.to_datetime(df_filtrado["FECHA"], dayfirst=True, errors="coerce")
            mask = df_filtrado["FECHA_DT"].isna()
            if mask.any():
                df_filtrado.loc[mask, "FECHA_DT"] = pd.to_datetime(df_filtrado.loc[mask, "FECHA"].astype(str).str.strip(), dayfirst=True, errors="coerce")
            df_filtrado = df_filtrado[df_filtrado["FECHA_DT"].notna()]
            df_filtrado["FECHA"] = df_filtrado["FECHA_DT"].dt.strftime("%d/%m/%Y")
            df_filtrado = df_filtrado.drop(columns=["FECHA_DT"])
            
        # Reemplazar valores vacíos o nulos en Débito y Crédito
        for col in ["DEBITO", "CREDITO"]:
            if col in df_filtrado.columns:
                df_filtrado[col] = df_filtrado[col].apply(mono_limpiar_monto_banplus).fillna(0.0)
                
        datos_normalizados = []
        for idx, fila in df_filtrado.iterrows():
            fecha_str = str(fila["FECHA"])
            referencia = str(fila.get("REFERENCIA", "")).strip().replace("'", "")
            descripcion = str(fila.get("DESCRIPCION", "")).strip()
            
            debito = float(fila.get("DEBITO", 0.0))
            credito = float(fila.get("CREDITO", 0.0))
            
            if credito > 0:
                tipo = "NC"
                monto = credito
            elif debito > 0:
                tipo = "ND"
                monto = debito
            else:
                continue
                
            datos_normalizados.append({
                "FECHA": fecha_str,
                "REFERENCIA": referencia,
                "TIPO": tipo,
                "DESCRIPCION": descripcion,
                "MONTO": monto
            })
            
        return pd.DataFrame(datos_normalizados)
    except Exception as e:
        st.error(f"❌ Error procesando archivo Banplus: {e}")
        return pd.DataFrame()

# =========================================================
# 🔥 PROCESAMIENTO PRINCIPAL - CON DETECCIÓN DIRECTA PARA TODOS LOS BANCOS
# =========================================================

def mono_procesar_archivo(df, usar_api=False, banco=""):
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
    
    tipos_ingresos = ["NC", "C", "CREDITO", "ABONO", "DP", "DEP"]
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
            if len(fecha_raw) == 7 and fecha_raw.isdigit():
                fecha = f"0{fecha_raw[0]}/{fecha_raw[1:3]}/{fecha_raw[3:]}"
            elif len(fecha_raw) == 8 and fecha_raw.isdigit():
                fecha = f"{fecha_raw[0:2]}/{fecha_raw[2:4]}/{fecha_raw[4:]}"
            else:
                fecha = fecha_raw

            tipo = str(fila[5]).strip().upper()
            descripcion = str(fila[6]).strip()
            referencia = str(fila[4]).strip()

            monto_bs = mono_convertir_monto(fila[7])
            if monto_bs is None or monto_bs == 0:
                continue
            
            fecha_obj = pd.to_datetime(fecha, dayfirst=True, errors="coerce")
            if pd.isna(fecha_obj):
                continue
            
            fecha_key = fecha_obj.strftime("%d/%m/%Y")
            # ✅ MISMA LÓGICA QUE MULTIBANCO (procesar_archivo): tasa por fecha del movimiento
            tasa = cache_tasas.get(fecha_key) or mono_obtener_tasa_por_fecha(fecha_obj, usar_api) or 1.0
            cache_tasas[fecha_key] = tasa

            monto_usd = mono_calcular_usd(monto_bs, tasa)
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

            # 🔥 CORREGIDO: Se eliminó la deduplicación por clave, porque descartaba
            # movimientos legítimos repetidos (mismo monto/referencia/descripción),
            # causando que ingresos (p. ej. Pago Móvil) no se tomaran todos del archivo.

            # 🔥 DETECCIÓN DE COMISIONES POR BANCO (función única con soporte multibanco)
            es_comision_banco = detectar_comision_por_banco(
                descripcion, referencia, tipo, monto_bs, banco
            )

            # Si es comisión del banco, clasificar como comisión
            if es_comision_banco:
                comisiones.append(registro)
            # Si es comisión bancaria detectada por función genérica
            elif mono_es_comision(descripcion):
                comisiones.append(registro)
            # Clasificar por tipo de movimiento
            elif tipo in tipos_ingresos:
                ingresos.append(registro)
            elif tipo in tipos_egresos:
                egresos.append(registro)
            else:
                # 🔥 NUEVO: Si no tiene tipo definido, clasificar por monto (crédito = ingreso, débito = egreso)
                if monto_bs > 0:
                    ingresos.append(registro)
                else:
                    egresos.append(registro)

        except Exception as e:
            continue

    return ingresos, egresos, comisiones

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
    if st.button("📊 SALDOS BANCARIOS MULTIBANCO", use_container_width=True, type="primary" if st.session_state.seccion_activa == "consolidado" else "secondary", key="nav_btn_consolidado"):
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
archivo_banplus = None
archivo_activo = None
saldo_manual_tesoro = 0.0

archivo = None

with st.sidebar:
    st.markdown(
        """
        <div style="display: flex; justify-content: center; align-items: center; background-color: #f1f3f5; padding: 12px; border-radius: 10px; border: 1px solid #e9ecef; margin-bottom: 5px;">
            <img src="https://raw.githubusercontent.com/pelobravo/clasificador-excel/main/LOGO.jpeg" style="width: 100px; border-radius: 8px;">
        </div>
        """,
        unsafe_allow_html=True
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

        with st.expander("🏦 BanPlus", expanded=False):
            archivo_banplus = st.file_uploader(
                "Cargar BanPlus",
                type=["xlsx", "xls", "xlsm"],
                accept_multiple_files=True,
                key="uploader_banplus"
            )

        with st.expander("🏦 Banco Activo", expanded=False):
            archivo_activo = st.file_uploader(
                "Cargar Banco Activo",
                type=["xlsx", "xls", "xlsm"],
                accept_multiple_files=True,
                key="uploader_activo"
            )

        with st.expander("🏦 Banco del Tesoro", expanded=False):
            saldo_manual_tesoro = st.number_input(
                "Saldo manual Banco del Tesoro (VES)",
                min_value=0.0,
                value=0.0,
                step=100.0,
                key="saldo_manual_tesoro"
            )

        with st.expander("💵 Banco Efectivo (Manual)", expanded=False):
            saldo_manual_efectivo = st.number_input(
                "Saldo manual Banco Efectivo (USD)",
                min_value=0.0,
                value=0.0,
                step=100.0,
                key="saldo_manual_efectivo"
            )

        with st.expander("🪙 Banco Binance (Manual)", expanded=False):
            saldo_manual_binance = st.number_input(
                "Saldo manual Banco Binance (USD)",
                min_value=0.0,
                value=0.0,
                step=100.0,
                key="saldo_manual_binance"
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
        "🌐 Tasa BCV automática desde la API (recomendado)",
        value=True,
        help="Descarga automáticamente la tasa oficial del BCV (API lab.geocenso.com) para la fecha de hoy "
             "y días recientes, sin necesidad de actualizarla manualmente. Si la API no responde, "
             "se usa la tasa local más reciente."
    )

    st.markdown("---")

    procesar = st.button(
        "🚀 Procesar",
        type="primary",
        use_container_width=True
    )

# =========================================================
# MAIN BODY CONDITIONAL ROUTING
# =========================================================

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
    if not archivo_banplus: st.session_state.saldo_banplus = 0.0
    if not archivo_activo: st.session_state.saldo_activo = 0.0
    st.session_state.saldo_tesoro = st.session_state.get("saldo_manual_tesoro", 0.0)
    st.session_state.saldo_efectivo = st.session_state.get("saldo_manual_efectivo", 0.0)
    st.session_state.saldo_binance = st.session_state.get("saldo_manual_binance", 0.0)

    # Selector de Moneda para los KPIs
    col_mon1, col_mon2 = st.columns([1.5, 3.5])
    with col_mon1:
        st.radio(
            "💱 Mostrar KPIs en:",
            options=["Dólares ($)", "Bolívares (Bs.)"],
            horizontal=True,
            key="selector_moneda_kpis"
        )
    
    # Determinar moneda seleccionada
    moneda_kpi = "USD" if st.session_state.get("selector_moneda_kpis", "Dólares ($)") == "Dólares ($)" else "VES"

    # Renderizado de KPIs
    tasa_dia, origen_tasa_dia = obtener_tasa_bcv_con_origen(fecha_fin, usar_api)
    
    # Efectivo y Binance se ingresan en USD, los convertimos a VES usando la tasa del día
    st.session_state.saldo_efectivo = st.session_state.get("saldo_manual_efectivo", 0.0) * tasa_dia
    st.session_state.saldo_binance = st.session_state.get("saldo_manual_binance", 0.0) * tasa_dia

    total_ves = (
        st.session_state.saldo_banesco + st.session_state.saldo_bnc + 
        st.session_state.saldo_mercantil + st.session_state.saldo_venezuela + 
        st.session_state.saldo_provincial + st.session_state.saldo_bancamiga + 
        st.session_state.saldo_banplus + st.session_state.saldo_activo +
        st.session_state.saldo_tesoro +
        st.session_state.saldo_efectivo +
        st.session_state.saldo_binance
    )
    total_usd = total_ves / tasa_dia if tasa_dia > 0 else 0.0
    
    # 🔥 CALCULAR AUTOMÁTICAMENTE los ingresos totales
    # 🔧 CORRECCIÓN (2026-08-13): usar el total consolidado real de TODOS los bancos
    # (antes priorizaba solo los créditos del archivo de Venezuela, dejando fuera los demás bancos)
    total_ingresos_ves = st.session_state.get("total_ingresos_consolidado", 0.0)
    total_ingresos_usd = total_ingresos_ves / tasa_dia if tasa_dia > 0 else 0.0
    
    total_egresos_ves = st.session_state.get("total_egresos_ipago_ves", 0.0)
    total_egresos_usd = total_egresos_ves / tasa_dia if tasa_dia > 0 else 0.0

    bancos_con_saldo = []
    if st.session_state.saldo_banesco > 0: bancos_con_saldo.append(f"Banesco: Bs. {formato_venezolano(st.session_state.saldo_banesco)}")
    if st.session_state.saldo_bnc > 0: bancos_con_saldo.append(f"BNC: Bs. {formato_venezolano(st.session_state.saldo_bnc)}")
    if st.session_state.saldo_mercantil > 0: bancos_con_saldo.append(f"Mercantil: Bs. {formato_venezolano(st.session_state.saldo_mercantil)}")
    if st.session_state.saldo_venezuela > 0: bancos_con_saldo.append(f"BDV: Bs. {formato_venezolano(st.session_state.saldo_venezuela)}")
    if st.session_state.saldo_provincial > 0: bancos_con_saldo.append(f"Provincial: Bs. {formato_venezolano(st.session_state.saldo_provincial)}")
    if st.session_state.saldo_bancamiga > 0: bancos_con_saldo.append(f"Bancamiga: Bs. {formato_venezolano(st.session_state.saldo_bancamiga)}")
    if st.session_state.saldo_banplus > 0: bancos_con_saldo.append(f"BanPlus: Bs. {formato_venezolano(st.session_state.saldo_banplus)}")
    if st.session_state.saldo_activo > 0: bancos_con_saldo.append(f"Banco Activo: Bs. {formato_venezolano(st.session_state.saldo_activo)}")
    if st.session_state.saldo_tesoro > 0: bancos_con_saldo.append(f"Tesoro: Bs. {formato_venezolano(st.session_state.saldo_tesoro)}")
    saldo_ef_usd = st.session_state.get("saldo_manual_efectivo", 0.0)
    if st.session_state.saldo_efectivo > 0: bancos_con_saldo.append(f"Efectivo: Bs. {formato_venezolano(st.session_state.saldo_efectivo)} (${saldo_ef_usd:,.2f})")
    saldo_bin_usd = st.session_state.get("saldo_manual_binance", 0.0)
    if st.session_state.saldo_binance > 0: bancos_con_saldo.append(f"Binance: Bs. {formato_venezolano(st.session_state.saldo_binance)} (${saldo_bin_usd:,.2f})")

    kpi_subtitle_text = " | ".join(bancos_con_saldo) if bancos_con_saldo else "Sin saldos cargados"

    if moneda_kpi == "USD":
        val_saldos = f"${total_usd:,.2f}"
        sub_saldos = f"Bs. {formato_venezolano(total_ves)}"
        
        val_ingresos = f"${total_ingresos_usd:,.2f}"
        sub_ingresos = f"Bs. {formato_venezolano(total_ingresos_ves)}"
        
        val_egresos = f"${total_egresos_usd:,.2f}"
        sub_egresos = f"Bs. {formato_venezolano(total_egresos_ves)}"
        
        title_saldos = "Total Saldos Bancos (USD)"
        title_ingresos = "Total Ingresos Bancos (USD)"
        title_egresos = "Total Egresos iPago (USD)"
        
        title_extra = "Total Equivalente (VES)"
        val_extra = f"Bs. {formato_venezolano(total_ves)}"
        sub_extra = "Saldos convertidos a tasa oficial"
    else:
        val_saldos = f"Bs. {formato_venezolano(total_ves)}"
        sub_saldos = f"${total_usd:,.2f} USD"
        
        val_ingresos = f"Bs. {formato_venezolano(total_ingresos_ves)}"
        sub_ingresos = f"${total_ingresos_usd:,.2f} USD"
        
        val_egresos = f"Bs. {formato_venezolano(total_egresos_ves)}"
        sub_egresos = f"${total_egresos_usd:,.2f} USD"
        
        title_saldos = "Total Saldos Bancos (VES)"
        title_ingresos = "Total Ingresos Bancos (VES)"
        title_egresos = "Total Egresos iPago (VES)"
        
        title_extra = "Total Equivalente (USD)"
        val_extra = f"${total_usd:,.2f}"
        sub_extra = "Convertido al tipo de cambio oficial"

    st.markdown(f"""
    <div class="kpi-container">
        <div class="kpi-card">
            <div class="kpi-title">{title_saldos}</div>
            <div class="kpi-value">{val_saldos}</div>
            <div class="kpi-subtitle">{kpi_subtitle_text if moneda_kpi == 'VES' else sub_saldos}</div>
        </div>
        <div class="kpi-card">
            <div class="kpi-title">{title_ingresos}</div>
            <div class="kpi-value">{val_ingresos}</div>
            <div class="kpi-subtitle">{sub_ingresos}</div>
        </div>
        <div class="kpi-card">
            <div class="kpi-title">{title_egresos}</div>
            <div class="kpi-value">{val_egresos}</div>
            <div class="kpi-subtitle">{sub_egresos}</div>
        </div>
        <div class="kpi-card">
            <div class="kpi-title">Tasa BCV ({fecha_fin.strftime('%d/%m/%Y')})</div>
            <div class="kpi-value">{tasa_dia:.4f} VES/USD</div>
            <div class="kpi-subtitle">{origen_tasa_dia}</div>
        </div>
        <div class="kpi-card">
            <div class="kpi-title">{title_extra}</div>
            <div class="kpi-value">{val_extra}</div>
            <div class="kpi-subtitle">{sub_extra}</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # =========================================================
    # RESUMEN DE FECHAS PROCESADAS
    # =========================================================
    if st.session_state.info_fechas_por_banco:
        with st.expander("📅 Resumen de Fechas Procesadas por Banco", expanded=False):
            data_fechas = []
            for banco, info in st.session_state.info_fechas_por_banco.items():
                fechas_excluidas = []
                fechas_procesadas = info.get('fechas') or ([info.get('fecha')] if info.get('fecha') else [])
                for fecha, count in info.get('detalle_fechas', {}).items():
                    if fecha not in fechas_procesadas:
                        fechas_excluidas.append(f"{fecha} ({count})")
                
                data_fechas.append({
                    "Banco": banco,
                    "Fecha Procesada": ", ".join(fechas_procesadas) if fechas_procesadas else 'No detectada',
                    "Registros Procesados": info.get('registros', 0),
                    "Registros Excluidos": info.get('excluidos', 0),
                    "Total Original": info.get('total_original', 0),
                    "Fechas Excluidas": ", ".join(fechas_excluidas) or "Ninguna"
                })
            
            df_fechas = pd.DataFrame(data_fechas)
            st.dataframe(df_fechas, use_container_width=True)
            
            # Descargar reporte de fechas
            csv = df_fechas.to_csv(index=False)
            st.download_button(
                label="📥 Descargar Resumen de Fechas (CSV)",
                data=csv,
                file_name=f"resumen_fechas_procesadas_{date.today().strftime('%Y%m%d')}.csv",
                mime="text/csv"
            )

    # =========================================================
    # LEER ARCHIVOS CARGADOS
    # =========================================================
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

    # 🔧 CORRECCIÓN (2026-08-13): reiniciar acumulador de ingresos por banco en cada procesamiento
    st.session_state.creditos_por_banco = {}

    # 🔥 REINICIAR acumuladores de resumen de saldos y fechas por banco (multibanco)
    st.session_state.info_fechas_por_banco = {}
    st.session_state.resumen_bancos = {b: {"saldo_inicial": 0.0, "saldo_final": 0.0, "creditos_total": 0.0, "debitos_total": 0.0} for b in _BANCOS_RESUMEN}

    # 1. Banesco
    if archivo_banesco:
        st.session_state.saldo_banesco = 0.0
        datos_banesco = []
        for idx, arch in enumerate(archivo_banesco, 1):
            try:
                nombre = arch.name.lower()
                if nombre.endswith(".xlsx") or nombre.endswith(".xlsm"):
                    df_raw = pd.read_excel(arch, engine="openpyxl", header=None)
                else:
                    df_raw = pd.read_html(arch)[0]
            
                saldo_arch = obtener_saldo_banco(df_raw, "banesco")
                datos_banesco.append({
                    "fecha": _fecha_archivo(arch.name),
                    "saldo": saldo_arch,
                    "resumen": extraer_resumen_banco(df_raw, "banesco"),
                })
            
            
                df_normalizado = procesar_banesco(df_raw)
                df_convertido = convertir_a_formato_mercantil(df_normalizado, "banesco")
                if not df_convertido.empty:
                    list_df_convertidos.append(df_convertido)
                    st.session_state.creditos_por_banco["banesco"] = st.session_state.creditos_por_banco.get("banesco", 0.0) + _sumar_creditos_convertidos(df_convertido)
                    if "Banesco" not in bancos_procesados:
                        bancos_procesados.append("Banesco")
            except Exception as e:
                st.error(f"❌ Error leyendo Banesco ({arch.name}): {e}")
        if datos_banesco:
            saldo_ultimo_banesco, resumen_banesco = _fijar_saldo_ultimo_dia(datos_banesco)
            st.session_state.saldo_banesco = saldo_ultimo_banesco
            _acumular_resumen_banco("Banesco", resumen_banesco)
            saldos_detalle_excel.append(("Banesco", saldo_ultimo_banesco))
    else:
        saldos_detalle_excel.append(("Banesco", 0.0))

    # 2. BNC
    if archivo_bnc:
        st.session_state.saldo_bnc = 0.0
        datos_bnc = []
        for idx, arch in enumerate(archivo_bnc, 1):
            try:
                df_raw = leer_excel_con_encabezados(arch)
                encabezado = None
                for i in range(min(30, len(df_raw))):
                    fila = df_raw.iloc[i].fillna("").astype(str)
                    texto = " ".join(fila.tolist()).lower()
                    if "fecha" in texto and ("descripcion" in texto or "descripción" in texto):
                        encabezado = i
                        break
            
                saldo_arch = obtener_saldo_banco(df_raw, "bnc", encabezado)
                datos_bnc.append({
                    "fecha": _fecha_archivo(arch.name),
                    "saldo": saldo_arch,
                    "resumen": extraer_resumen_banco(df_raw, "bnc"),
                })
            
            
                df_normalizado = procesar_bnc(df_raw)
                df_convertido = convertir_a_formato_mercantil(df_normalizado, "bnc")
                if not df_convertido.empty:
                    list_df_convertidos.append(df_convertido)
                    st.session_state.creditos_por_banco["bnc"] = st.session_state.creditos_por_banco.get("bnc", 0.0) + _sumar_creditos_convertidos(df_convertido)
                    if "BNC" not in bancos_procesados:
                        bancos_procesados.append("BNC")
            except Exception as e:
                st.error(f"❌ Error leyendo BNC ({arch.name}): {e}")
        if datos_bnc:
            saldo_ultimo_bnc, resumen_bnc = _fijar_saldo_ultimo_dia(datos_bnc)
            st.session_state.saldo_bnc = saldo_ultimo_bnc
            _acumular_resumen_banco("BNC", resumen_bnc)
            saldos_detalle_excel.append(("BNC", saldo_ultimo_bnc))
    else:
        saldos_detalle_excel.append(("BNC", 0.0))

    # 3. Mercantil
    if archivo_mercantil:
        st.session_state.saldo_mercantil = 0.0
        datos_mercantil = []
        for idx, arch in enumerate(archivo_mercantil, 1):
            try:
                df_raw = leer_excel_sin_encabezados(arch)
                df_raw = preparar_df_con_encabezado_dinamico(df_raw)
                saldo_arch = obtener_saldo_banco(df_raw, "mercantil")
                datos_mercantil.append({
                    "fecha": _fecha_archivo(arch.name),
                    "saldo": saldo_arch,
                    "resumen": extraer_resumen_banco(df_raw, "mercantil"),
                })
            
            
                df_convertido = convertir_a_formato_mercantil(df_raw, "mercantil")
                if not df_convertido.empty:
                    list_df_convertidos.append(df_convertido)
                    st.session_state.creditos_por_banco["mercantil"] = st.session_state.creditos_por_banco.get("mercantil", 0.0) + _sumar_creditos_convertidos(df_convertido)
                    if "Mercantil" not in bancos_procesados:
                        bancos_procesados.append("Mercantil")
            except Exception as e:
                st.error(f"❌ Error leyendo Mercantil ({arch.name}): {e}")
        if datos_mercantil:
            saldo_ultimo_mercantil, resumen_mercantil = _fijar_saldo_ultimo_dia(datos_mercantil)
            st.session_state.saldo_mercantil = saldo_ultimo_mercantil
            _acumular_resumen_banco("Mercantil", resumen_mercantil)
            saldos_detalle_excel.append(("Mercantil", saldo_ultimo_mercantil))
    else:
        saldos_detalle_excel.append(("Mercantil", 0.0))

    # 4. BDV
    if archivo_venezuela:
        st.session_state.saldo_venezuela = 0.0
        st.session_state.total_creditos_venezuela = 0.0  # 🔥 Inicializar
        
        datos_venezuela = []
        for idx, arch in enumerate(archivo_venezuela, 1):
            try:
                df_raw = leer_excel_sin_encabezados(arch)
                
                # 🔥 RESUMEN DEL ESTADO DE CUENTA BDV (extractor dedicado):
                # lee Saldo Inicial/Final y Créditos/Débitos Total directamente del reporte
                # (columnas "Saldo Inicial", "Saldo Final", "Total Crédito", "Todal Débito"),
                # en vez de adivinar con escáneres de texto (que devolvían montos errados
                # o 0 en Créditos/Débitos para el Banco de Venezuela).
                resumen_raw = extraer_resumen_banco(df_raw, "venezuela")
                
                # Créditos del archivo original: prioridad al Total Crédito reportado (col K);
                # respaldo: suma de la columna de créditos (col 5)
                total_creditos_raw = float(resumen_raw.get("creditos_total") or 0.0)
                if total_creditos_raw <= 0.001:
                    col_credito = 5  # Columna de créditos en BDV
                    for i in range(1, len(df_raw)):
                        try:
                            fila = df_raw.iloc[i]
                            if pd.notna(fila[col_credito]):
                                valor_str = str(fila[col_credito]).strip()
                                valor_str = valor_str.replace(".", "").replace(",", ".")
                                if valor_str and valor_str != "0" and valor_str != "0.0":
                                    valor = float(valor_str)
                                    if valor > 0:
                                        total_creditos_raw += valor
                        except:
                            continue
                
                # Acumular en session_state
                st.session_state.total_creditos_venezuela += total_creditos_raw
                
                # Saldo final: prioridad al extractor BDV; respaldo: escáner genérico
                saldo_arch = float(resumen_raw.get("saldo_final") or 0.0)
                if saldo_arch <= 0.001:
                    saldo_arch = obtener_saldo_banco(df_raw, "venezuela")
                datos_venezuela.append({
                    "fecha": _fecha_archivo(arch.name),
                    "saldo": saldo_arch,
                    "resumen": resumen_raw,
                })
                
                
                df_normalizado = procesar_venezuela_simple(df_raw)
                df_convertido = convertir_venezuela_a_formato_mercantil(df_normalizado)
                if not df_convertido.empty:
                    list_df_convertidos.append(df_convertido)
                    st.session_state.creditos_por_banco["venezuela"] = st.session_state.creditos_por_banco.get("venezuela", 0.0) + _sumar_creditos_convertidos(df_convertido)
                    if "Venezuela" not in bancos_procesados:
                        bancos_procesados.append("Venezuela")
            except Exception as e:
                st.error(f"❌ Error leyendo BDV ({arch.name}): {e}")
                # No dejar el banco en silencio: se limpian los acumuladores para que el
                # error sea visible (el saldo NO debe quedar con el valor de otro día)
                st.session_state.saldo_venezuela = 0.0
                st.session_state.total_creditos_venezuela = 0.0
        if datos_venezuela:
            saldo_ultimo_venezuela, resumen_venezuela = _fijar_saldo_ultimo_dia(datos_venezuela)
            st.session_state.saldo_venezuela = saldo_ultimo_venezuela
            _acumular_resumen_banco("Banco de Venezuela (BDV)", resumen_venezuela)
            saldos_detalle_excel.append(("Banco de Venezuela (BDV)", saldo_ultimo_venezuela))
    else:
        saldos_detalle_excel.append(("Banco de Venezuela (BDV)", 0.0))

    # 5. Provincial
    if archivo_provincial:
        st.session_state.saldo_provincial = 0.0
        datos_provincial = []
        for idx, arch in enumerate(archivo_provincial, 1):
            try:
                df_raw = leer_excel_sin_encabezados(arch)
                saldo_arch = obtener_saldo_banco(df_raw, "provincial")
                datos_provincial.append({
                    "fecha": _fecha_archivo(arch.name),
                    "saldo": saldo_arch,
                    "resumen": extraer_resumen_banco(df_raw, "provincial"),
                })
            
            
                df_normalizado = procesar_provincial(df_raw)
                df_convertido = convertir_a_formato_mercantil(df_normalizado, "provincial")
                if not df_convertido.empty:
                    list_df_convertidos.append(df_convertido)
                    st.session_state.creditos_por_banco["provincial"] = st.session_state.creditos_por_banco.get("provincial", 0.0) + _sumar_creditos_convertidos(df_convertido)
                    if "Provincial" not in bancos_procesados:
                        bancos_procesados.append("Provincial")
            except Exception as e:
                st.error(f"❌ Error leyendo Provincial ({arch.name}): {e}")
        if datos_provincial:
            saldo_ultimo_provincial, resumen_provincial = _fijar_saldo_ultimo_dia(datos_provincial)
            st.session_state.saldo_provincial = saldo_ultimo_provincial
            _acumular_resumen_banco("Provincial", resumen_provincial)
            saldos_detalle_excel.append(("Provincial", saldo_ultimo_provincial))
    else:
        saldos_detalle_excel.append(("Provincial", 0.0))

    # 6. Bancamiga
    if archivo_bancamiga:
        st.session_state.saldo_bancamiga = 0.0
        datos_bancamiga = []
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
                        try:
                            df_raw = pd.read_html(arch, decimal=',', thousands='.')[0]
                        except Exception:
                            arch.seek(0)
                            df_raw = leer_tabla_html(arch)
                
                if isinstance(df_raw.columns, pd.MultiIndex):
                    df_raw.columns = df_raw.columns.get_level_values(-1)
            
                saldo_arch = obtener_saldo_banco(df_raw, "bancamiga")
                datos_bancamiga.append({
                    "fecha": _fecha_archivo(arch.name),
                    "saldo": saldo_arch,
                    "resumen": extraer_resumen_banco(df_raw, "bancamiga"),
                })
            
            
                df_normalizado = procesar_bancamiga(df_raw)
                df_convertido = convertir_a_formato_mercantil(df_normalizado, "bancamiga")
                if not df_convertido.empty:
                    list_df_convertidos.append(df_convertido)
                    st.session_state.creditos_por_banco["bancamiga"] = st.session_state.creditos_por_banco.get("bancamiga", 0.0) + _sumar_creditos_convertidos(df_convertido)
                    if "Bancamiga" not in bancos_procesados:
                        bancos_procesados.append("Bancamiga")
            except Exception as e:
                st.error(f"❌ Error leyendo Bancamiga ({arch.name}): {e}")
        if datos_bancamiga:
            saldo_ultimo_bancamiga, resumen_bancamiga = _fijar_saldo_ultimo_dia(datos_bancamiga)
            st.session_state.saldo_bancamiga = saldo_ultimo_bancamiga
            _acumular_resumen_banco("Bancamiga", resumen_bancamiga)
            saldos_detalle_excel.append(("Bancamiga", saldo_ultimo_bancamiga))
    else:
        saldos_detalle_excel.append(("Bancamiga", 0.0))

    # 6.5. BanPlus
    if archivo_banplus:
        st.session_state.saldo_banplus = 0.0
        datos_banplus = []
        for idx, arch in enumerate(archivo_banplus, 1):
            try:
                nombre = arch.name.lower()
                if nombre.endswith(".xlsx") or nombre.endswith(".xlsm"):
                    df_raw = pd.read_excel(arch, engine="openpyxl", header=None)
                else:
                    try:
                        df_raw = pd.read_excel(arch, header=None)
                    except Exception:
                        arch.seek(0)
                        try:
                            df_raw = pd.read_html(arch)[0]
                        except Exception:
                            arch.seek(0)
                            df_raw = leer_tabla_html(arch)
            
                saldo_arch = obtener_saldo_banco(df_raw, "banplus")
                datos_banplus.append({
                    "fecha": _fecha_archivo(arch.name),
                    "saldo": saldo_arch,
                    "resumen": extraer_resumen_banco(df_raw, "banplus"),
                })
            
            
                df_normalizado = procesar_banplus(df_raw)
                df_convertido = convertir_a_formato_mercantil(df_normalizado, "banplus")
                if not df_convertido.empty:
                    list_df_convertidos.append(df_convertido)
                    st.session_state.creditos_por_banco["banplus"] = st.session_state.creditos_por_banco.get("banplus", 0.0) + _sumar_creditos_convertidos(df_convertido)
                    if "BanPlus" not in bancos_procesados:
                        bancos_procesados.append("BanPlus")
            except Exception as e:
                st.error(f"❌ Error leyendo BanPlus ({arch.name}): {e}")
        if datos_banplus:
            saldo_ultimo_banplus, resumen_banplus = _fijar_saldo_ultimo_dia(datos_banplus)
            st.session_state.saldo_banplus = saldo_ultimo_banplus
            _acumular_resumen_banco("BanPlus", resumen_banplus)
            saldos_detalle_excel.append(("BanPlus", saldo_ultimo_banplus))
    else:
        saldos_detalle_excel.append(("BanPlus", 0.0))

    # 6.75. Banco Activo
    if archivo_activo:
        st.session_state.saldo_activo = 0.0
        datos_activo = []
        for idx, arch in enumerate(archivo_activo, 1):
            try:
                nombre = arch.name.lower()
                if nombre.endswith(".xlsx") or nombre.endswith(".xlsm"):
                    df_raw = pd.read_excel(arch, engine="openpyxl", header=None)
                else:
                    try:
                        df_raw = pd.read_excel(arch, header=None)
                    except Exception:
                        arch.seek(0)
                        try:
                            df_raw = pd.read_html(arch)[0]
                        except Exception:
                            arch.seek(0)
                            df_raw = leer_tabla_html(arch)
                
                if isinstance(df_raw.columns, pd.MultiIndex):
                    df_raw.columns = df_raw.columns.get_level_values(-1)
            
                saldo_arch = obtener_saldo_banco(df_raw, "activo")
                datos_activo.append({
                    "fecha": _fecha_archivo(arch.name),
                    "saldo": saldo_arch,
                    "resumen": extraer_resumen_banco(df_raw, "activo"),
                })
            
            
                df_normalizado = procesar_banco_activo(df_raw)
                df_convertido = convertir_a_formato_mercantil(df_normalizado, "activo")
                if not df_convertido.empty:
                    list_df_convertidos.append(df_convertido)
                    st.session_state.creditos_por_banco["activo"] = st.session_state.creditos_por_banco.get("activo", 0.0) + _sumar_creditos_convertidos(df_convertido)
                    if "Activo" not in bancos_procesados:
                        bancos_procesados.append("Activo")
            except Exception as e:
                st.error(f"❌ Error leyendo Banco Activo ({arch.name}): {e}")
        if datos_activo:
            saldo_ultimo_activo, resumen_activo = _fijar_saldo_ultimo_dia(datos_activo)
            st.session_state.saldo_activo = saldo_ultimo_activo
            _acumular_resumen_banco("Banco Activo", resumen_activo)
            saldos_detalle_excel.append(("Banco Activo", saldo_ultimo_activo))
    else:
        saldos_detalle_excel.append(("Banco Activo", 0.0))

    # 7. Tesoro (Manual)
    saldos_detalle_excel.append(("Banco del Tesoro", st.session_state.saldo_tesoro))

    # 7.5. Efectivo (Manual)
    saldos_detalle_excel.append(("Banco Efectivo", st.session_state.saldo_efectivo))

    # 7.6. Binance (Manual)
    saldos_detalle_excel.append(("Banco Binance", st.session_state.saldo_binance))

    st.session_state.saldos_detalle_excel = saldos_detalle_excel

    # 🔥 RE-RENDER ÚNICO TRAS LEER ARCHIVOS: los KPIs superiores (Total Saldos Bancos,
    # Ingresos, etc.) se dibujan ANTES de este bloque de lectura, por lo que en la
    # primera interacción mostrarían valores del run anterior (0 o de otro día). Un rerun
    # inmediato (solo cuando se cargaron archivos y NO se pulsó "Procesar") refresca las
    # tarjetas con los saldos recién leídos. El flag evita bucles de rerun.
    hay_archivos_bancos = bool(
        archivo_banesco or archivo_bnc or archivo_mercantil
        or archivo_venezuela or archivo_provincial
        or archivo_bancamiga or archivo_banplus or archivo_activo
    )
    if not procesar and hay_archivos_bancos and not st.session_state.get("_rerun_saldos"):
        st.session_state._rerun_saldos = True
        st.rerun()
    elif st.session_state.get("_rerun_saldos"):
        st.session_state._rerun_saldos = False

    # Recalcular el total consolidado con los datos extraídos
    st.session_state.saldo_efectivo = st.session_state.get("saldo_manual_efectivo", 0.0) * tasa_dia
    st.session_state.saldo_binance = st.session_state.get("saldo_manual_binance", 0.0) * tasa_dia

    total_ves = (
        st.session_state.saldo_banesco + st.session_state.saldo_bnc + 
        st.session_state.saldo_mercantil + st.session_state.saldo_venezuela + 
        st.session_state.saldo_provincial + st.session_state.saldo_bancamiga + 
        st.session_state.saldo_banplus + st.session_state.saldo_activo +
        st.session_state.saldo_tesoro +
        st.session_state.saldo_efectivo +
        st.session_state.saldo_binance
    )
    total_usd = total_ves / tasa_dia if tasa_dia > 0 else 0.0

    # =========================================================
    # SALDOS DEL ESTADO DE CUENTA POR BANCO (CAPTURADOS DEL ARCHIVO)
    # =========================================================
    if st.session_state.get("resumen_bancos"):
        with st.expander("📊 Saldos del Estado de Cuenta por Banco", expanded=False):
            data_saldos = []
            for banco_s, info_s in st.session_state.resumen_bancos.items():
                if not any(info_s.values()):
                    continue
                data_saldos.append({
                    "Banco": banco_s,
                    "Saldo Inicial (VES)": info_s.get("saldo_inicial", 0) or 0,
                    "Saldo Final (VES)": info_s.get("saldo_final", 0) or 0,
                    "Créditos Total (VES)": info_s.get("creditos_total", 0) or 0,
                    "Débitos Total (VES)": info_s.get("debitos_total", 0) or 0,
                })
            if data_saldos:
                st.dataframe(pd.DataFrame(data_saldos), use_container_width=True)
                st.caption("Valores capturados de las filas de saldo de cada archivo (Saldo Inicial/Final, Créditos/Débitos Total). También se incluyen en el exportable Excel.")

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

                    # 🔧 CORRECCIÓN (2026-08-13): numérico de 7 dígitos (Mercantil compacto "6082026" -> 06/08/2026)
                    if len(val_str) == 7 and val_str.isdigit():
                        dt = pd.to_datetime(val_str.zfill(8), format="%d%m%Y", errors="coerce")
                        if pd.notna(dt):
                            fechas.append(dt)
                            continue
                
                    # Parseo flexible general
                    parsed = False
                    for fmt in ["%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"]:
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
            
            # Ajustar automáticamente el rango si las fechas consolidadas están fuera del rango seleccionado
            if not fechas_convertidas.empty:
                min_file_date = fechas_convertidas.min()
                max_file_date = fechas_convertidas.max()
                if pd.notna(min_file_date) and pd.notna(max_file_date):
                    if fecha_inicio_dt > min_file_date or fecha_fin_dt < max_file_date:
                        fecha_inicio_dt = min_file_date
                        fecha_fin_dt = max_file_date
                        st.info(f"💡 **Rango de fechas ajustado automáticamente** al contenido de los archivos: {min_file_date.strftime('%d/%m/%Y')} al {max_file_date.strftime('%d/%m/%Y')}")
            
            df_original = df_original[(fechas_convertidas >= fecha_inicio_dt) & (fechas_convertidas <= fecha_fin_dt)]
            
            # 🔥 MODIFICAR: Calcular la suma de ingresos para los KPIs
            # Usar el total de créditos del archivo original si existe
            total_ingresos_ves_calc = st.session_state.get('total_creditos_venezuela', 0.0)
            
            # Si no hay créditos raw (otro banco o falló), usar el cálculo del procesamiento
            if total_ingresos_ves_calc == 0.0 and not df_original.empty:
                tipos_ingresos = ["NC", "C", "CREDITO", "ABONO", "DP", "DEP"]
                tipo_col = df_original.iloc[:, 5].astype(str).str.strip().str.upper()
                ingresos_filas = df_original[tipo_col.isin(tipos_ingresos)]
                
                for val in ingresos_filas.iloc[:, 7]:
                    try:
                        if isinstance(val, (int, float)):
                            total_ingresos_ves_calc += abs(float(val))
                        else:
                            val_str = str(val).strip()
                            if "," in val_str and "." in val_str:
                                if val_str.find(".") < val_str.find(","):
                                    val_str = val_str.replace(".", "").replace(",", ".")
                                else:
                                    val_str = val_str.replace(",", "")
                            elif "," in val_str:
                                val_str = val_str.replace(",", ".")
                            total_ingresos_ves_calc += abs(float(val_str))
                    except:
                        pass
            
            st.session_state.total_ingresos_consolidado = total_ingresos_ves_calc
            
            # Calcular la suma de egresos del archivo iPago
            total_egresos_ipago_ves_calc = 0.0
            if df_ipago is not None and not df_ipago.empty:
                try:
                    fechas_ipago = pd.to_datetime(df_ipago['Fecha Pago'], errors='coerce')
                    df_ipago_filtrado = df_ipago[(fechas_ipago >= fecha_inicio_dt) & (fechas_ipago <= fecha_fin_dt)]
                    # Sumar la columna Monto convirtiendo a float por seguridad
                    monto_sum = 0.0
                    for val in df_ipago_filtrado['Monto']:
                        try:
                            if isinstance(val, (int, float)):
                                monto_sum += abs(float(val))
                            else:
                                val_str = str(val).strip()
                                if "," in val_str and "." in val_str:
                                    if val_str.find(".") < val_str.find(","):
                                        val_str = val_str.replace(".", "").replace(",", ".")
                                    else:
                                        val_str = val_str.replace(",", "")
                                elif "," in val_str:
                                    val_str = val_str.replace(",", ".")
                                monto_sum += abs(float(val_str))
                        except:
                            pass
                    total_egresos_ipago_ves_calc = monto_sum
                except Exception as e:
                    st.warning(f"⚠️ Error al calcular egresos de iPago: {e}")
            st.session_state.total_egresos_ipago_ves = total_egresos_ipago_ves_calc
            
            st.success(f"📅 Movimientos consolidados de {', '.join(bancos_procesados)} filtrados del {fecha_inicio_dt.strftime('%d/%m/%Y')} al {fecha_fin_dt.strftime('%d/%m/%Y')} ({len(df_original)} registros)")
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
                    # 🔧 CORRECCIÓN (2026-08-13): TOTAL INGRESOS = suma real de los ingresos de TODOS los bancos
                    # (antes se usaba solo los créditos del archivo de Venezuela, dejando fuera Bancamiga, BanPlus, etc.)
                    if not df_ingresos.empty:
                        st.session_state.total_ingresos_consolidado = float(df_ingresos["MONTO BS"].sum())
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

                        hoja_resumen["B3"] = "CONCILIACIÓN BANCARIA - RESUMEN DE SALDOS"
                        hoja_resumen["B3"].font = Font(bold=True, size=11, color="555555")
                        hoja_resumen["B3"].alignment = alineacion_izquierda

                        # Datos de la Empresa / Reporte (A la derecha superior)
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

                        # Cabeceras tabla (5 columnas: BANCOS, TOTAL (VES), CONVERSIÓN (USD), INGRESOS (VES), INGRESOS (USD))
                        headers_r = ["BANCOS", "TOTAL (VES)", "CONVERSIÓN (USD)", "INGRESOS (VES)", "INGRESOS (USD)"]
                        for col_num, header in enumerate(headers_r, 2):
                            cell = hoja_resumen.cell(row=8, column=col_num)
                            cell.value = header
                            cell.fill = azul_oscuro
                            cell.font = blanco_bold
                            cell.alignment = alineacion_centro
                            cell.border = borde_fino

                        # Datos
                        bancos_data = st.session_state.get("saldos_detalle_excel", [
                            ("Banesco", st.session_state.saldo_banesco),
                            ("BNC", st.session_state.saldo_bnc),
                            ("Mercantil", st.session_state.saldo_mercantil),
                            ("Banco de Venezuela (BDV)", st.session_state.saldo_venezuela),
                            ("Provincial", st.session_state.saldo_provincial),
                            ("Bancamiga", st.session_state.saldo_bancamiga),
                            ("BanPlus", st.session_state.saldo_banplus),
                            ("Banco Activo", st.session_state.saldo_activo),
                            ("Banco del Tesoro", st.session_state.saldo_tesoro),
                            ("Banco Efectivo", st.session_state.saldo_efectivo),
                            ("Banco Binance", st.session_state.saldo_binance)
                        ])

                        fila_r = 9
                        mapeo_ingresos_banco = {
                            "Banesco": "banesco", "BNC": "bnc", "Mercantil": "mercantil",
                            "Banco de Venezuela (BDV)": "venezuela", "Provincial": "provincial",
                            "Bancamiga": "bancamiga", "BanPlus": "banplus", "Banco Activo": "activo",
                            "Banco del Tesoro": "tesoro", "Banco Efectivo": "efectivo", "Banco Binance": "binance",
                        }
                        for banco_n, saldo_v in bancos_data:
                            cell_b = hoja_resumen.cell(row=fila_r, column=2, value=banco_n)
                            cell_b.border = borde_fino
                            cell_b.alignment = alineacion_izquierda
                        
                            cell_s = hoja_resumen.cell(row=fila_r, column=3, value=saldo_v)
                            cell_s.border = borde_fino
                            cell_s.number_format = '#,##0.00'
                            cell_s.alignment = alineacion_derecha
                        
                            usd_v = saldo_v / tasa_dia if tasa_dia > 0 else 0.0
                            cell_u = hoja_resumen.cell(row=fila_r, column=4, value=usd_v)
                            cell_u.border = borde_fino
                            cell_u.number_format = '$#,##0.00'
                            cell_u.alignment = alineacion_derecha

                            # 🔧 CORRECCIÓN (2026-08-13): mostrar los INGRESOS de cada banco en VES y USD
                            nombre_base = str(banco_n).split(" - Cuenta")[0]
                            clave_ing = mapeo_ingresos_banco.get(nombre_base, "")
                            ing_ves = st.session_state.creditos_por_banco.get(clave_ing, 0.0) if clave_ing else 0.0
                            ing_usd = ing_ves / tasa_dia if tasa_dia > 0 else 0.0
                            cell_iv = hoja_resumen.cell(row=fila_r, column=5, value=ing_ves)
                            cell_iv.border = borde_fino
                            cell_iv.number_format = '#,##0.00'
                            cell_iv.alignment = alineacion_derecha

                            cell_i = hoja_resumen.cell(row=fila_r, column=6, value=ing_usd)
                            cell_i.border = borde_fino
                            cell_i.number_format = '$#,##0.00'
                            cell_i.alignment = alineacion_derecha
                        
                            if fila_r % 2 == 0:
                                for col in range(2, 7):
                                    hoja_resumen.cell(row=fila_r, column=col).fill = gris_claro
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

                        # Total Ingresos Archivos
                        fila_r += 1
                        cell_total_ing_lbl = hoja_resumen.cell(row=fila_r, column=2, value="TOTAL INGRESOS ARCHIVOS")
                        cell_total_ing_lbl.font = negro_bold
                        cell_total_ing_lbl.border = borde_fino
                        cell_total_ing_lbl.fill = amarillo
                    
                        tot_ing_ves_export = st.session_state.get("total_ingresos_consolidado", 0.0)
                        cell_total_ing_ves = hoja_resumen.cell(row=fila_r, column=3, value=tot_ing_ves_export)
                        cell_total_ing_ves.font = negro_bold
                        cell_total_ing_ves.border = borde_fino
                        cell_total_ing_ves.number_format = '#,##0.00'
                        cell_total_ing_ves.fill = amarillo
                        cell_total_ing_ves.alignment = alineacion_derecha
                    
                        tot_ing_usd_export = tot_ing_ves_export / tasa_dia if tasa_dia > 0 else 0.0
                        cell_total_ing_usd = hoja_resumen.cell(row=fila_r, column=4, value=tot_ing_usd_export)
                        cell_total_ing_usd.font = negro_bold
                        cell_total_ing_usd.border = borde_fino
                        cell_total_ing_usd.number_format = '$#,##0.00'
                        cell_total_ing_usd.fill = amarillo
                        cell_total_ing_usd.alignment = alineacion_derecha

                        # Total Egresos iPago
                        fila_r += 1
                        cell_total_egr_lbl = hoja_resumen.cell(row=fila_r, column=2, value="TOTAL EGRESOS IPAGO")
                        cell_total_egr_lbl.font = negro_bold
                        cell_total_egr_lbl.border = borde_fino
                        cell_total_egr_lbl.fill = amarillo
                    
                        tot_egr_ves_export = st.session_state.get("total_egresos_ipago_ves", 0.0)
                        cell_total_egr_ves = hoja_resumen.cell(row=fila_r, column=3, value=tot_egr_ves_export)
                        cell_total_egr_ves.font = negro_bold
                        cell_total_egr_ves.border = borde_fino
                        cell_total_egr_ves.number_format = '#,##0.00'
                        cell_total_egr_ves.fill = amarillo
                        cell_total_egr_ves.alignment = alineacion_derecha
                    
                        tot_egr_usd_export = tot_egr_ves_export / tasa_dia if tasa_dia > 0 else 0.0
                        cell_total_egr_usd = hoja_resumen.cell(row=fila_r, column=4, value=tot_egr_usd_export)
                        cell_total_egr_usd.font = negro_bold
                        cell_total_egr_usd.border = borde_fino
                        cell_total_egr_usd.number_format = '$#,##0.00'
                        cell_total_egr_usd.fill = amarillo
                        cell_total_egr_usd.alignment = alineacion_derecha

                        # 🔥 SECCIÓN: RESUMEN DE FECHAS PROCESADAS
                        if st.session_state.info_fechas_por_banco:
                            fila_r += 2
                            hoja_resumen.merge_cells(start_row=fila_r, start_column=2, end_row=fila_r, end_column=4)
                            cell_fechas_titulo = hoja_resumen.cell(row=fila_r, column=2, value="RESUMEN DE FECHAS PROCESADAS POR BANCO")
                            cell_fechas_titulo.font = Font(bold=True, size=11, color="1E3A5F")
                            cell_fechas_titulo.alignment = alineacion_centro
                            fila_r += 1
                            
                            # Encabezados de la tabla de fechas
                            headers_fechas = ["BANCO", "FECHA PROCESADA", "REGISTROS", "EXCLUIDOS", "TOTAL ORIGINAL"]
                            for col_num, header in enumerate(headers_fechas, 2):
                                cell = hoja_resumen.cell(row=fila_r, column=col_num)
                                cell.value = header
                                cell.fill = azul_oscuro
                                cell.font = blanco_bold
                                cell.alignment = alineacion_centro
                                cell.border = borde_fino
                            fila_r += 1
                            
                            for banco, info in st.session_state.info_fechas_por_banco.items():
                                hoja_resumen.cell(row=fila_r, column=2, value=banco).border = borde_fino
                                fechas_dias = info.get('fechas') or ([info.get('fecha')] if info.get('fecha') else [])
                                hoja_resumen.cell(row=fila_r, column=3, value=", ".join(fechas_dias) if fechas_dias else 'No detectada').border = borde_fino
                                hoja_resumen.cell(row=fila_r, column=4, value=info.get('registros', 0)).border = borde_fino
                                hoja_resumen.cell(row=fila_r, column=5, value=info.get('excluidos', 0)).border = borde_fino
                                hoja_resumen.cell(row=fila_r, column=6, value=info.get('total_original', 0)).border = borde_fino
                                
                                # Resaltar filas con exclusiones
                                if info.get('excluidos', 0) > 0:
                                    for col in range(2, 7):
                                        hoja_resumen.cell(row=fila_r, column=col).fill = PatternFill(start_color="FFE699", end_color="FFE699", fill_type="solid")
                                
                                # Mostrar detalle de fechas excluidas
                                detalle_fechas = info.get('detalle_fechas', {})
                                if detalle_fechas:
                                    fechas_dias = info.get('fechas') or ([info.get('fecha')] if info.get('fecha') else [])
                                    fechas_excluidas = [f"{fecha} ({count})" for fecha, count in detalle_fechas.items() if fecha not in fechas_dias]
                                    if fechas_excluidas:
                                        fila_r += 1
                                        for col in range(2, 7):
                                            hoja_resumen.cell(row=fila_r, column=col).fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
                                        hoja_resumen.cell(row=fila_r, column=2, value="  ⚠️ Fechas excluidas:").border = borde_fino
                                        hoja_resumen.cell(row=fila_r, column=3, value=", ".join(fechas_excluidas)).border = borde_fino
                                        hoja_resumen.merge_cells(start_row=fila_r, start_column=3, end_row=fila_r, end_column=6)
                                
                                fila_r += 1

                        # 🔥 SECCIÓN: SALDOS POR BANCO (ESTADO DE CUENTA)
                        # Estos valores (Saldo Inicial/Final, Créditos/Débitos Total) los toma el sistema Validador.
                        resumen_bancos = st.session_state.get("resumen_bancos", {})
                        if resumen_bancos:
                            fila_r += 2
                            hoja_resumen.merge_cells(start_row=fila_r, start_column=2, end_row=fila_r, end_column=6)
                            cell_saldos_titulo = hoja_resumen.cell(row=fila_r, column=2, value="SALDOS POR BANCO (ESTADO DE CUENTA)")
                            cell_saldos_titulo.font = Font(bold=True, size=11, color="1E3A5F")
                            cell_saldos_titulo.alignment = alineacion_centro
                            fila_r += 1

                            headers_saldos = ["BANCO", "SALDO INICIAL (VES)", "SALDO FINAL (VES)", "CRÉDITOS TOTAL (VES)", "DÉBITOS TOTAL (VES)"]
                            for col_num, header in enumerate(headers_saldos, 2):
                                cell = hoja_resumen.cell(row=fila_r, column=col_num)
                                cell.value = header
                                cell.fill = azul_oscuro
                                cell.font = blanco_bold
                                cell.alignment = alineacion_centro
                                cell.border = borde_fino
                            fila_r += 1

                            tot_ini_s = tot_fin_s = tot_cred_s = tot_deb_s = 0.0
                            for banco_n, info_s in resumen_bancos.items():
                                ini_s = info_s.get("saldo_inicial", 0) or 0
                                fin_s = info_s.get("saldo_final", 0) or 0
                                cred_s = info_s.get("creditos_total", 0) or 0
                                deb_s = info_s.get("debitos_total", 0) or 0
                                tot_ini_s += ini_s
                                tot_fin_s += fin_s
                                tot_cred_s += cred_s
                                tot_deb_s += deb_s
                                hoja_resumen.cell(row=fila_r, column=2, value=banco_n).border = borde_fino
                                for col_num, val_s in [(3, ini_s), (4, fin_s), (5, cred_s), (6, deb_s)]:
                                    cell_s = hoja_resumen.cell(row=fila_r, column=col_num, value=val_s)
                                    cell_s.border = borde_fino
                                    cell_s.number_format = '#,##0.00'
                                    cell_s.alignment = alineacion_derecha
                                if fila_r % 2 == 0:
                                    for col in range(2, 7):
                                        hoja_resumen.cell(row=fila_r, column=col).fill = gris_claro
                                fila_r += 1

                            # Totales (etiqueta en col B, VES en C y USD en D: formato que lee el Validador)
                            fila_r += 1
                            hoja_resumen.cell(row=fila_r, column=2, value="TOTAL SALDO INICIAL BANCOS").font = negro_bold
                            hoja_resumen.cell(row=fila_r, column=2).fill = amarillo
                            hoja_resumen.cell(row=fila_r, column=2).border = borde_fino
                            c_ini_ves = hoja_resumen.cell(row=fila_r, column=3, value=tot_ini_s)
                            c_ini_ves.number_format = '#,##0.00'
                            c_ini_ves.fill = amarillo
                            c_ini_ves.border = borde_fino
                            c_ini_usd = hoja_resumen.cell(row=fila_r, column=4, value=(tot_ini_s / tasa_dia if tasa_dia > 0 else 0.0))
                            c_ini_usd.number_format = '$#,##0.00'
                            c_ini_usd.fill = amarillo
                            c_ini_usd.border = borde_fino

                            fila_r += 1
                            hoja_resumen.cell(row=fila_r, column=2, value="TOTAL SALDO FINAL ESTADO CUENTA").font = negro_bold
                            hoja_resumen.cell(row=fila_r, column=2).fill = amarillo
                            hoja_resumen.cell(row=fila_r, column=2).border = borde_fino
                            c_fin_ves = hoja_resumen.cell(row=fila_r, column=3, value=tot_fin_s)
                            c_fin_ves.number_format = '#,##0.00'
                            c_fin_ves.fill = amarillo
                            c_fin_ves.border = borde_fino
                            c_fin_usd = hoja_resumen.cell(row=fila_r, column=4, value=(tot_fin_s / tasa_dia if tasa_dia > 0 else 0.0))
                            c_fin_usd.number_format = '$#,##0.00'
                            c_fin_usd.fill = amarillo
                            c_fin_usd.border = borde_fino

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
                                "FECHA", "REFERENCIA", "DESCRIPCIÓN", "DESCRIPCIÓN ORIGINAL",
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
        st.session_state.total_ingresos_consolidado = 0.0
        st.session_state.total_egresos_ipago_ves = 0.0
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
    # =========================================================

    df_ipago = None

    if archivo:
        st.info(f"📄 Archivo: **{archivo.name}** - {archivo.size/1024:.1f} KB")

        try:
            # =========================================================
            # 🔥 DETECCIÓN DE BANCO - NUEVO ORDEN
            # =========================================================
        
            # 1. Detectar SIEMPRE por el nombre primero
            banco = mono_detectar_banco_por_nombre(archivo.name)
        
            # 2. Solo si no se reconoce (queda como "mercantil"), intentar por contenido
            if banco == "mercantil":
                banco_contenido = mono_detectar_banco_por_contenido(archivo)
                if banco_contenido:
                    banco = banco_contenido
        
            st.success(f"🏦 **Banco detectado:** {banco.upper()}")
        
            if banco == "mercantil":
                df_original = mono_leer_excel_sin_encabezados(archivo)
            
            elif banco == "banesco":
                try:
                    nombre = archivo.name.lower()
                    if nombre.endswith(".xlsx") or nombre.endswith(".xlsm"):
                        df_raw = pd.read_excel(archivo, engine="openpyxl", header=None)
                    else:
                        df_raw = pd.read_html(archivo)[0]
                    df_normalizado = mono_procesar_banesco(df_raw)
                    df_original = mono_convertir_a_formato_mercantil(df_normalizado, banco)
                except Exception as e:
                    st.error(f"Error leyendo Banesco: {str(e)}")
                    st.stop()
            
            elif banco == "tesoro":
                try:
                    df_raw = pd.read_excel(archivo, engine="openpyxl")
                    df_normalizado = mono_procesar_tesoro(df_raw)
                    df_original = mono_convertir_a_formato_mercantil(df_normalizado, banco)
                except Exception as e:
                    st.error(f"Error leyendo Tesoro: {str(e)}")
                    st.stop()
            
            elif banco == "bancamiga":
                try:
                    # 🔥 CARGA DE BANCAMIGA
                    nombre = archivo.name.lower()
                
                    if nombre.endswith(".xlsx") or nombre.endswith(".xlsm"):
                        df_raw = pd.read_excel(archivo, engine="openpyxl", header=None)
                    else:
                        try:
                            # Intentar leer como Excel .xls real
                            df_raw = pd.read_excel(archivo, header=None)
                        except Exception:
                            # Si realmente es HTML disfrazado de .xls
                            archivo.seek(0)
                            df_raw = pd.read_html(archivo, decimal=',', thousands='.')[0]
                    
                    if isinstance(df_raw.columns, pd.MultiIndex):
                        df_raw.columns = df_raw.columns.get_level_values(-1)
                
                    df_normalizado = mono_procesar_bancamiga(df_raw)
                    if df_normalizado.empty:
                        st.error("No se pudieron procesar los datos de Bancamiga.")
                        st.stop()
                    df_original = mono_convertir_a_formato_mercantil(df_normalizado, banco)
                except Exception as e:
                    st.error(f"Error leyendo Bancamiga: {str(e)}")
                    st.stop()
            
            elif banco == "provincial":
                try:
                    df_raw = mono_leer_excel_sin_encabezados(archivo)
                    df_normalizado = mono_procesar_provincial(df_raw)
                    if df_normalizado.empty:
                        st.error("No se pudieron procesar los datos de Provincial.")
                        st.stop()
                    df_original = mono_convertir_a_formato_mercantil(df_normalizado, banco)
                except Exception as e:
                    st.error(f"Error leyendo Provincial: {str(e)}")
                    st.stop()
            
            elif banco == "venezuela":
                # 🔥 LEER SIN ENCABEZADOS (igual que Mercantil)
                df_raw = mono_leer_excel_sin_encabezados(archivo)
                df_normalizado = mono_procesar_venezuela_simple(df_raw)
                if df_normalizado.empty:
                    st.stop()
            
                # Venezuela trabaja directamente con el dataframe normalizado
                df_original = mono_convertir_venezuela_a_formato_mercantil(df_normalizado)
            
                # Fechas para Venezuela
                fechas_convertidas = pd.to_datetime(
                    df_normalizado["FECHA"],
                    dayfirst=True,
                    errors="coerce"
                )
            
            elif banco == "banplus":
                try:
                    df_raw = mono_leer_excel_sin_encabezados(archivo)
                    df_normalizado = mono_procesar_banplus(df_raw)
                    if df_normalizado.empty:
                        st.error("No se pudieron procesar los datos de Banplus.")
                        st.stop()
                    df_original = mono_convertir_a_formato_mercantil(df_normalizado, banco)
                except Exception as e:
                    st.error(f"Error leyendo Banplus: {str(e)}")
                    st.stop()
            
            elif banco == "bnc":
                df_raw = leer_excel_con_encabezados(archivo)
                df_normalizado = mono_procesar_bnc(df_raw)
                df_original = mono_convertir_a_formato_mercantil(df_normalizado, banco)
            
            elif banco == "activo":
                try:
                    df_raw = mono_leer_excel_sin_encabezados(archivo)
                    df_normalizado = procesar_banco_activo(df_raw)
                    if df_normalizado.empty:
                        st.error("No se pudieron procesar los datos de Banco Activo.")
                        st.stop()
                    df_original = mono_convertir_a_formato_mercantil(df_normalizado, banco)
                    # Guardar el saldo si es necesario
                    st.session_state.saldo_activo = obtener_saldo_final_banco_activo(df_raw)
                except Exception as e:
                    st.error(f"Error leyendo Banco Activo: {str(e)}")
                    st.stop()
            
            else:
                df_raw = leer_excel_con_encabezados(archivo)
                df_original = mono_convertir_a_formato_mercantil(df_raw, banco)
            
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

                # Ajustar automáticamente el rango si las fechas del archivo están fuera del rango seleccionado
                if not fechas_convertidas.empty:
                    min_file_date = fechas_convertidas.min()
                    max_file_date = fechas_convertidas.max()
                    if pd.notna(min_file_date) and pd.notna(max_file_date):
                        if fecha_inicio_dt > min_file_date or fecha_fin_dt < max_file_date:
                            fecha_inicio_dt = min_file_date
                            fecha_fin_dt = max_file_date
                            st.info(f"💡 **Rango de fechas ajustado automáticamente** al contenido del archivo: {min_file_date.strftime('%d/%m/%Y')} al {max_file_date.strftime('%d/%m/%Y')}")

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
            
                st.success(f"Filtro de fechas aplicado: {fecha_inicio_dt.strftime('%d/%m/%Y')} a {fecha_fin_dt.strftime('%d/%m/%Y')}")
            except Exception as e:
                st.warning(f"Error filtrando fechas: {e}")
            
            if df_original.empty or len(df_original) == 0:
                st.error("❌ No se encontraron movimientos válidos en el rango de fechas.")
                st.stop()

            with st.expander("👁️ Vista previa archivo original"):
                st.dataframe(df_original.head(20), use_container_width=True)

            # =========================================================
            # RESUMEN DE FECHAS PROCESADAS
            # =========================================================
            if st.session_state.info_fechas_por_banco:
                with st.expander("📅 Resumen de Fechas Procesadas", expanded=False):
                    data_fechas = []
                    for banco, info in st.session_state.info_fechas_por_banco.items():
                        fechas_excluidas = []
                        for fecha, count in info.get('detalle_fechas', {}).items():
                            if fecha != info.get('fecha', ''):
                                fechas_excluidas.append(f"{fecha} ({count})")
                        
                        data_fechas.append({
                            "Banco": banco,
                            "Fecha Procesada": info.get('fecha', 'No detectada'),
                            "Registros Procesados": info.get('registros', 0),
                            "Registros Excluidos": info.get('excluidos', 0),
                            "Total Original": info.get('total_original', 0),
                            "Fechas Excluidas": ", ".join(fechas_excluidas) or "Ninguna"
                        })
                    
                    df_fechas = pd.DataFrame(data_fechas)
                    st.dataframe(df_fechas, use_container_width=True)

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
                    # 🔥 USAR LA FUNCIÓN CORREGIDA QUE MANEJA BDV CORRECTAMENTE
                    ingresos, egresos, comisiones = mono_procesar_archivo(df_original, usar_api, banco=banco)

                df_ingresos = pd.DataFrame(ingresos)
                df_egresos = pd.DataFrame(egresos)
                df_comisiones = pd.DataFrame(comisiones)

                # =========================================================
                # 🔥 CRUCE CON IPAGO - VERSIÓN MEJORADA (CRUCE FLEXIBLE)
                # =========================================================
                if archivo_ipago and not df_egresos.empty:
                    # Enriquecer egresos con datos de iPago
                    df_egresos = mono_enriquecer_egresos_con_ipago(df_egresos, df_ipago)
                
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

                # =========================================================
                # 🔥 BARRIDO FINAL DE COMISIONES (EGRESOS → COMISIONES)
                # Mueve a la sección COMISIONES cualquier egreso cuya descripción
                # (original del banco o de iPago) o tipo de egreso iPago indique
                # una comisión bancaria. Evita que comisiones queden dentro de
                # EGRESOS al exportar el Excel.
                # =========================================================
                if not df_egresos.empty:
                    columnas_comisiones = [
                        "FECHA", "REFERENCIA", "DESCRIPCIÓN", "DESCRIPCION_ORIGINAL",
                        "MONTO BS", "TASA BCV", "MONTO USD",
                        "STATUS", "OBSERVACIÓN", "TIPO_PAGO", "PROVEEDOR_IPAGO"
                    ]
                    mascara_comision_final = pd.Series(False, index=df_egresos.index)
                    for idx in df_egresos.index:
                        fila = df_egresos.loc[idx]
                        desc_original = str(fila.get("DESCRIPCION_ORIGINAL", "") or "")
                        desc_actual = str(fila.get("DESCRIPCIÓN", "") or "")
                        tipo_egreso_ipago = str(fila.get("OBSERVACIÓN", "") or "")
                        desc_upper = (desc_original or desc_actual).upper()
                        if (
                            mono_es_comision(desc_original)
                            or mono_es_comision(desc_actual)
                            or any(x in desc_upper for x in [
                                "COMISION PAGO A PROVEEDORES",
                                "COMISION PAGO A PROVEEDOR",
                                "COM PAGO A PROVEEDORES",
                                "COM. PAGO A PROVEEDORES",
                                "COM PAGO PROVEEDORES"
                            ])
                            or "COMISION" in tipo_egreso_ipago.upper()
                            or "COMISIÓN" in tipo_egreso_ipago.upper()
                        ):
                            mascara_comision_final.at[idx] = True

                    if mascara_comision_final.any():
                        df_comisiones_final = df_egresos[mascara_comision_final].copy()
                        for col in columnas_comisiones:
                            if col not in df_comisiones_final.columns:
                                df_comisiones_final[col] = ""
                        df_comisiones_final = df_comisiones_final[columnas_comisiones]
                        if not df_comisiones.empty:
                            df_comisiones = pd.concat([df_comisiones, df_comisiones_final], ignore_index=True)
                        else:
                            df_comisiones = df_comisiones_final
                        df_egresos = df_egresos[~mascara_comision_final].copy()
                        st.success(f"💳 Barrido final: {len(df_comisiones_final)} comisiones movidas de EGRESOS a COMISIONES.")

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
                            try:
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
                            except Exception:
                                pass

                            fila_data += 1

                        total_cell = hoja.cell(row=fila_data, column=4)
                        total_cell.value = f"TOTAL {titulo}"
                        total_cell.font = Font(bold=True)

                        # 🔥 CORREGIDO: Convertir a numérico antes de sumar para que
                        # una celda en formato texto no produzca totales erróneos
                        total_bs_cell = hoja.cell(row=fila_data, column=5)
                        total_bs_cell.value = pd.to_numeric(dataframe["MONTO BS"], errors="coerce").sum() if not dataframe.empty else 0
                        total_bs_cell.number_format = '#,##0.00'
                        total_bs_cell.fill = color_total

                        monto_total = hoja.cell(row=fila_data, column=7)
                        monto_total.value = pd.to_numeric(dataframe["MONTO USD"], errors="coerce").sum() if not dataframe.empty else 0
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
        ✅ **Bancos soportados:** Mercantil, Banco de Venezuela, Banesco, Provincial, BNC, Tesoro, Bancamiga, BanPlus, Banco Activo.
        ✅ Clasifica automáticamente: Ingresos (NC, C, CREDITO, ABONO), Egresos (ND, D, DEBITO, DEBIT), Comisiones.
        ✅ **NUEVO:** Cruce inteligente con iPago usando REFERENCIA + DESCRIPCIÓN.
        ✅ **NUEVO:** Exporta con datos completos de iPago: Proveedor, Tipo de Egreso, Tipo de Pago.
        ✅ **NUEVO:** Conserva la descripción original del banco y la reemplaza con la de iPago cuando hay coincidencia.
        ✅ Calcula USD con tasa BCV real por fecha.
        ✅ Exporta reporte profesional con todas las columnas.
        """)
