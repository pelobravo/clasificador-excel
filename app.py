import streamlit as st
import pandas as pd
from io import BytesIO

# Configuración de la página
st.set_page_config(
    page_title="Clasificador de Excel - Grupo Bodeguita Oriente",
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
    .stMetric { background-color: white; border-radius: 12px; padding: 15px; }
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
    st.title("Clasificador de Excel - INGRESOS / EGRESOS")
    st.markdown("### Organiza transacciones bancarias por tipo de movimiento")

st.markdown("---")

# Sidebar
with st.sidebar:
    st.image("https://raw.githubusercontent.com/pelobravo/clasificador-excel/main/LOGO.jpeg", width=100)
    st.markdown("---")
    archivo = st.file_uploader("📂 Cargar archivo Excel del banco", type=['xlsx', 'xls'])
    st.markdown("---")
    
    # Configuración de columnas
    st.header("⚙️ Configuración de columnas")
    col_fecha = st.text_input("Columna de FECHA", value="FECHA")
    col_ref = st.text_input("Columna de REFERENCIA", value="REFERENCIA")
    col_desc = st.text_input("Columna de DESCRIPCIÓN", value="DESCRIPCIÓN")
    col_monto = st.text_input("Columna de MONTO", value="MONTO")
    col_tasa = st.text_input("Columna de TASA DEL DÍA", value="TASA")
    col_total = st.text_input("Columna de TOTAL", value="TOTAL")
    col_tipo = st.text_input("Columna de TIPO (NC/ND)", value="TIPO")
    
    st.markdown("---")
    procesar = st.button("🚀 Procesar y Organizar", type="primary", use_container_width=True)

def detectar_tipo_movimiento(row, col_tipo, col_monto):
    """
    Detecta si es INGRESO o EGRESO basado en:
    1. Si existe columna TIPO: NC=INGRESO, ND=EGRESO
    2. Si no, por el signo del monto: positivo=INGRESO, negativo=EGRESO
    """
    # Verificar si existe la columna TIPO y tiene valor
    if col_tipo and col_tipo in row.index and pd.notna(row[col_tipo]):
        tipo = str(row[col_tipo]).upper()
        if 'NC' in tipo:
            return 'INGRESO'
        elif 'ND' in tipo:
            return 'EGRESO'
    
    # Si no, usar el signo del monto
    monto = row[col_monto] if col_monto in row.index else 0
    try:
        monto_num = float(monto) if pd.notna(monto) else 0
        if monto_num > 0:
            return 'INGRESO'
        elif monto_num < 0:
            return 'EGRESO'
        else:
            return 'SIN CLASIFICAR'
    except:
        return 'ERROR'

def procesar_archivo(df, config):
    """
    Procesa el archivo y separa INGRESOS y EGRESOS con las columnas requeridas
    """
    # Columnas finales requeridas
    columnas_finales = ['FECHA', 'REFERENCIA', 'DESCRIPCIÓN', 'MONTO', 'TASA DEL DÍA', 'TOTAL', 'OBSERVACIÓN', 'CLASIFICACIÓN DE EGRESO']
    
    # Crear estructura de resultado
    ingresos = []
    egresos = []
    
    for _, row in df.iterrows():
        # Determinar tipo de movimiento
        tipo = detectar_tipo_movimiento(row, config['col_tipo'], config['col_monto'])
        
        # Construir fila con columnas finales
        nueva_fila = {
            'FECHA': row[config['col_fecha']] if config['col_fecha'] in row.index else '',
            'REFERENCIA': row[config['col_ref']] if config['col_ref'] in row.index else '',
            'DESCRIPCIÓN': row[config['col_desc']] if config['col_desc'] in row.index else '',
            'MONTO': abs(float(row[config['col_monto']])) if config['col_monto'] in row.index and pd.notna(row[config['col_monto']]) else '',
            'TASA DEL DÍA': row[config['col_tasa']] if config['col_tasa'] in row.index else '',
            'TOTAL': row[config['col_total']] if config['col_total'] in row.index else '',
            'OBSERVACIÓN': '',
            'CLASIFICACIÓN DE EGRESO': ''
        }
        
        if tipo == 'INGRESO':
            ingresos.append(nueva_fila)
        elif tipo == 'EGRESO':
            nueva_fila['CLASIFICACIÓN DE EGRESO'] = 'EGRESO'
            egresos.append(nueva_fila)
    return ingresos, egresos

# Área principal
if archivo:
    st.info(f"📄 Archivo: **{archivo.name}** - {archivo.size/1024:.1f} KB")
    
    try:
        df_original = pd.read_excel(archivo)
        
        with st.expander("👁️ Vista previa del archivo original (primeras 10 filas)"):
            st.dataframe(df_original.head(10), use_container_width=True)
            st.caption(f"Total filas: {len(df_original)} | Columnas: {df_original.columns.tolist()}")
        
        if procesar:
            config = {
                'col_fecha': col_fecha,
                'col_ref': col_ref,
                'col_desc': col_desc,
                'col_monto': col_monto,
                'col_tasa': col_tasa,
                'col_total': col_total,
                'col_tipo': col_tipo
            }
            
            with st.spinner('Procesando archivo...'):
                ingresos, egresos = procesar_archivo(df_original, config)
            
            st.success(f"✅ Procesado: {len(ingresos)} INGRESOS | {len(egresos)} EGRESOS")
            
            # Mostrar métricas
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("💰 Total INGRESOS", len(ingresos))
            with col2:
                st.metric("💸 Total EGRESOS", len(egresos))
            with col3:
                total = sum([f['MONTO'] for f in ingresos if f['MONTO']]) - sum([f['MONTO'] for f in egresos if f['MONTO']])
                st.metric("📊 Saldo", f"{total:,.2f}")
            
            # Crear Excel con formato
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                # Hoja de INGRESOS
                if ingresos:
                    df_ingresos = pd.DataFrame(ingresos)
                    df_ingresos.to_excel(writer, sheet_name='INGRESOS', index=False)
                
                # Hoja de EGRESOS
                if egresos:
                    df_egresos = pd.DataFrame(egresos)
                    df_egresos.to_excel(writer, sheet_name='EGRESOS', index=False)
                
                # Hoja de RESUMEN
                resumen = pd.DataFrame([
                    {'Concepto': 'Total INGRESOS', 'Cantidad': len(ingresos), 'Monto Total': sum([f['MONTO'] for f in ingresos if f['MONTO']])},
                    {'Concepto': 'Total EGRESOS', 'Cantidad': len(egresos), 'Monto Total': sum([f['MONTO'] for f in egresos if f['MONTO']])},
                    {'Concepto': 'SALDO NETO', 'Cantidad': '', 'Monto Total': sum([f['MONTO'] for f in ingresos if f['MONTO']]) - sum([f['MONTO'] for f in egresos if f['MONTO']])}
                ])
                resumen.to_excel(writer, sheet_name='RESUMEN', index=False)
            
            output.seek(0)
            
            st.subheader("📊 Resultados")
            
            # Mostrar previsualización
            tab1, tab2 = st.tabs(["📈 INGRESOS", "📉 EGRESOS"])
            with tab1:
                if ingresos:
                    st.dataframe(pd.DataFrame(ingresos), use_container_width=True)
                else:
                    st.info("No hay INGRESOS en este archivo")
            with tab2:
                if egresos:
                    st.dataframe(pd.DataFrame(egresos), use_container_width=True)
                else:
                    st.info("No hay EGRESOS en este archivo")
            
            # Botón de descarga
            st.download_button(
                label="📥 Descargar Excel organizado",
                data=output.getvalue(),
                file_name=f"balance_{archivo.name}",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
    
    except Exception as e:
        st.error(f"❌ Error: {str(e)}")
        st.info("Verifica que los nombres de las columnas coincidan con las configuradas")

else:
    st.markdown("""
    ### 👋 ¡Bienvenido al Clasificador de INGRESOS y EGRESOS!
    
    **¿Cómo funciona?**
    
    1. 📂 **Carga** un archivo Excel de un banco (Banesco, Mercantil, Venezuela, etc.)
    2. ⚙️ **Configura** los nombres de las columnas (FECHA, REFERENCIA, DESCRIPCIÓN, MONTO, etc.)
    3. 🚀 **Presiona** "Procesar y Organizar"
    
    ### 🎯 ¿Qué hace automáticamente?
    - **Diferencia INGRESOS de EGRESOS** (por NC/ND o por signo del monto)
    - **Organiza** en dos hojas separadas dentro del mismo Excel
    - **Agrega** una hoja de RESUMEN con totales
    
    ### 📋 Formato de salida
    | FECHA | REFERENCIA | DESCRIPCIÓN | MONTO | TASA DEL DÍA | TOTAL | OBSERVACIÓN | CLASIFICACIÓN DE EGRESO |
    |-------|------------|-------------|-------|--------------|-------|-------------|------------------------|
    
    ---
    **💡 Tip:** El programa es compatible con cualquier banco. Solo configura correctamente los nombres de las columnas.
    """)

# Footer
st.markdown("---")
st.markdown(
    """
    <div class='footer'>
        <strong>Grupo Bodeguita Oriente</strong> - Clasificador de INGRESOS/EGRESOS v7.0
    </div>
    """,
    unsafe_allow_html=True
)
