import streamlit as st
import pandas as pd
from io import BytesIO

# Configuración de la página
st.set_page_config(
    page_title="Clasificador Bancario - Grupo Bodeguita Oriente",
    page_icon="🏦",
    layout="wide"
)

# Estilos CSS
st.markdown("""
    <style>
    .stApp { background-color: #ffffff; }
    .stButton > button {
        background-color: #1e3a5f;
        color: white;
        border-radius: 8px;
        padding: 10px 24px;
        font-weight: bold;
        border: none;
    }
    .stButton > button:hover { background-color: #2c5282; }
    h1, h2, h3 { color: #1e3a5f; }
    .stMetric { background-color: white; border-radius: 12px; padding: 15px; box-shadow: 0 2px 8px rgba(0,0,0,0.05); }
    .footer { text-align: center; color: #666; padding: 20px; font-size: 14px; }
    </style>
""", unsafe_allow_html=True)

# Título
col_logo, col_titulo = st.columns([1, 5])
with col_logo:
    try:
        st.image("LOGO.jpeg", width=80)
    except:
        pass
with col_titulo:
    st.title("Clasificador Bancario")
    st.markdown("### INGRESOS / EGRESOS / COMISIONES")
st.markdown("---")

with st.sidebar:
    st.image("https://raw.githubusercontent.com/pelobravo/clasificador-excel/main/LOGO.jpeg", width=100)
    st.markdown("---")
    archivo = st.file_uploader("📂 Cargar archivo Excel del banco", type=['xlsx', 'xls'])
    st.markdown("---")
    procesar = st.button("🚀 Procesar y Organizar", type="primary", use_container_width=True)

def detectar_columnas(df):
    """Detecta automáticamente las columnas importantes"""
    columnas = {}
    
    # Buscar columna de fecha
    for col in df.columns:
        col_lower = str(col).lower()
        if 'fecha' in col_lower or 'date' in col_lower:
            columnas['fecha'] = col
            break
    if 'fecha' not in columnas:
        columnas['fecha'] = None
    
    # Buscar columna de referencia
    for col in df.columns:
        col_lower = str(col).lower()
        if 'referencia' in col_lower or 'ref' in col_lower or 'refer' in col_lower:
            columnas['referencia'] = col
            break
    if 'referencia' not in columnas:
        columnas['referencia'] = None
    
    # Buscar columna de descripción
    for col in df.columns:
        col_lower = str(col).lower()
        if 'descripcion' in col_lower or 'concepto' in col_lower or 'detalle' in col_lower:
            columnas['descripcion'] = col
            break
    if 'descripcion' not in columnas:
        # Si no hay descripción, usar la primera columna de texto
        for col in df.columns:
            if df[col].dtype == 'object':
                columnas['descripcion'] = col
                break
    if 'descripcion' not in columnas:
        columnas['descripcion'] = None
    
    # Buscar columna de monto (prioridad a TOTAL, MONTO, SALDO, etc.)
    for col in df.columns:
        col_lower = str(col).lower()
        if 'total' in col_lower or 'monto' in col_lower or 'saldo' in col_lower:
            columnas['monto'] = col
            break
    if 'monto' not in columnas:
        # Buscar columna numérica
        for col in df.columns:
            if pd.api.types.is_numeric_dtype(df[col]):
                columnas['monto'] = col
                break
    if 'monto' not in columnas:
        columnas['monto'] = None
    
    return columnas

def procesar_archivo(df, columnas):
    """Procesa el archivo y separa INGRESOS, EGRESOS y COMISIONES"""
    ingresos = []
    egresos = []
    comisiones = []
    
    for _, fila in df.iterrows():
        # Obtener valores básicos
        fecha = fila[columnas['fecha']] if columnas['fecha'] and pd.notna(fila[columnas['fecha']]) else ''
        referencia = fila[columnas['referencia']] if columnas['referencia'] and pd.notna(fila[columnas['referencia']]) else ''
        descripcion = fila[columnas['descripcion']] if columnas['descripcion'] and pd.notna(fila[columnas['descripcion']]) else ''
        monto = fila[columnas['monto']] if columnas['monto'] and pd.notna(fila[columnas['monto']]) else 0
        
        # Convertir monto a número
        try:
            monto_num = float(str(monto).replace(',', '.').replace(' ', ''))
        except:
            monto_num = 0
        
        # Formatear fecha
        if fecha and not isinstance(fecha, str):
            fecha = str(fecha)
        
        registro = {
            'FECHA': fecha,
            'REFERENCIA': referencia,
            'DESCRIPCIÓN': descripcion,
            'MONTO': abs(monto_num),
            'TOTAL': abs(monto_num)
        }
        
        # Clasificar por tipo
        desc_lower = str(descripcion).lower()
        
        # Detectar comisiones
        if 'comision' in desc_lower:
            comisiones.append(registro)
        elif monto_num > 0:
            ingresos.append(registro)
        elif monto_num < 0:
            registro['CLASIFICACIÓN'] = 'EGRESO'
            egresos.append(registro)
    
    return ingresos, egresos, comisiones

# Área principal
if archivo:
    st.info(f"📄 Archivo: **{archivo.name}** - {archivo.size/1024:.1f} KB")
    
    try:
        df_original = pd.read_excel(archivo)
        
        with st.expander("👁️ Vista previa del archivo original (primeras 10 filas)"):
            st.dataframe(df_original.head(10), use_container_width=True)
            st.caption(f"Columnas detectadas: {list(df_original.columns)}")
        
        if procesar:
            with st.spinner('Analizando archivo...'):
                # Detectar columnas automáticamente
                columnas = detectar_columnas(df_original)
                
                st.write("**Columnas detectadas:**")
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.info(f"📅 Fecha: `{columnas['fecha']}`" if columnas['fecha'] else "⚠️ No se encontró columna de fecha")
                with col2:
                    st.info(f"🔢 Referencia: `{columnas['referencia']}`" if columnas['referencia'] else "⚠️ No se encontró referencia")
                with col3:
                    st.info(f"📝 Descripción: `{columnas['descripcion']}`" if columnas['descripcion'] else "⚠️ No se encontró descripción")
                with col4:
                    st.info(f"💰 Monto: `{columnas['monto']}`" if columnas['monto'] else "❌ CRÍTICO: No se encontró columna de monto")
                
                if not columnas['monto']:
                    st.error("❌ No se pudo identificar la columna de montos. Verifica que tu archivo tenga una columna con TOTAL, MONTO o SALDO")
                else:
                    ingresos, egresos, comisiones = procesar_archivo(df_original, columnas)
                    
                    st.success(f"✅ Procesado: {len(ingresos)} INGRESOS | {len(egresos)} EGRESOS | {len(comisiones)} COMISIONES")
                    
                    # Métricas
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("💰 INGRESOS", len(ingresos))
                    with col2:
                        st.metric("💸 EGRESOS", len(egresos))
                    with col3:
                        st.metric("💳 COMISIONES", len(comisiones))
                    
                    # Crear Excel
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        if ingresos:
                            pd.DataFrame(ingresos).to_excel(writer, sheet_name='INGRESOS', index=False)
                        if egresos:
                            pd.DataFrame(egresos).to_excel(writer, sheet_name='EGRESOS', index=False)
                        if comisiones:
                            pd.DataFrame(comisiones).to_excel(writer, sheet_name='COMISIONES', index=False)
                        
                        # Resumen
                        resumen = pd.DataFrame([
                            {'TIPO': 'INGRESOS', 'CANTIDAD': len(ingresos), 'MONTO TOTAL': sum(i['MONTO'] for i in ingresos)},
                            {'TIPO': 'EGRESOS', 'CANTIDAD': len(egresos), 'MONTO TOTAL': sum(e['MONTO'] for e in egresos)},
                            {'TIPO': 'COMISIONES', 'CANTIDAD': len(comisiones), 'MONTO TOTAL': sum(c['MONTO'] for c in comisiones)}
                        ])
                        resumen.to_excel(writer, sheet_name='RESUMEN', index=False)
                    
                    output.seek(0)
                    
                    # Mostrar pestañas
                    st.subheader("📊 Resultados")
                    tabs = st.tabs(["INGRESOS", "EGRESOS", "COMISIONES"])
                    
                    with tabs[0]:
                        if ingresos:
                            st.dataframe(pd.DataFrame(ingresos), use_container_width=True)
                        else:
                            st.info("No hay INGRESOS")
                    
                    with tabs[1]:
                        if egresos:
                            st.dataframe(pd.DataFrame(egresos), use_container_width=True)
                        else:
                            st.info("No hay EGRESOS")
                    
                    with tabs[2]:
                        if comisiones:
                            st.dataframe(pd.DataFrame(comisiones), use_container_width=True)
                        else:
                            st.info("No hay COMISIONES")
                    
                    # Descarga
                    st.download_button(
                        label="📥 Descargar Excel organizado",
                        data=output.getvalue(),
                        file_name=f"balance_{archivo.name}",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
    
    except Exception as e:
        st.error(f"❌ Error: {str(e)}")
        st.info("Verifica que el archivo sea un Excel válido")

else:
    st.markdown("""
    ### 👋 Clasificador Bancario
    
    **¿Cómo funciona?**
    
    1. 📂 Carga tu extracto bancario en Excel
    2. 🔍 El programa detecta automáticamente las columnas
    3. 📊 Exporta INGRESOS, EGRESOS y COMISIONES
    
    **Bancos compatibles:** Banesco, Mercantil, Provincial, Venezuela, etc.
    """)

st.markdown("---")
st.markdown('<div class="footer"><strong>Grupo Bodeguita Oriente</strong> - Clasificador Bancario v9.0</div>', unsafe_allow_html=True)
