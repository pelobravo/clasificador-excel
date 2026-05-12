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
    st.title("Clasificador de Excel - INGRESOS / EGRESOS / COMISIONES")
    st.markdown("### Clasifica automáticamente transacciones bancarias")

st.markdown("---")

# Sidebar
with st.sidebar:
    st.image("https://raw.githubusercontent.com/pelobravo/clasificador-excel/main/LOGO.jpeg", width=100)
    st.markdown("---")
    archivo = st.file_uploader("📂 Cargar archivo Excel del banco", type=['xlsx', 'xls'])
    st.markdown("---")
    
    procesar = st.button("🚀 Procesar y Organizar", type="primary", use_container_width=True)

def detectar_tipo_movimiento(fila, monto):
    """
    Detecta si es INGRESO, EGRESO o COMISION basado en:
    - COMISION: si la descripción contiene "comision"
    - INGRESO/EGRESO: por signo del monto
    """
    if monto > 0:
        return 'INGRESO'
    elif monto < 0:
        return 'EGRESO'
    else:
        return 'SIN CLASIFICAR'

def clasificar_por_concepto(df, lista_conceptos):
    """
    Clasifica las filas por concepto (ej: comision, pago movil)
    """
    resultados = {concepto: [] for concepto in lista_conceptos}
    
    for _, fila in df.iterrows():
        # Buscar en todas las columnas
        texto_fila = " ".join([str(v).lower() for v in fila.values if pd.notna(v)])
        
        for concepto in lista_conceptos:
            if concepto in texto_fila:
                # Encontrar la descripción completa
                descripcion = ""
                for col in df.columns:
                    valor = fila[col]
                    if pd.notna(valor) and concepto in str(valor).lower():
                        descripcion = str(valor)
                        break
                
                # Buscar fecha, referencia, monto
                fecha = ''
                referencia = ''
                monto = None
                total = ''
                
                for col in df.columns:
                    col_lower = str(col).lower()
                    valor = fila[col]
                    if pd.notna(valor):
                        if 'fecha' in col_lower:
                            fecha = valor
                        elif 'referencia' in col_lower or 'ref' in col_lower:
                            referencia = valor
                        elif 'monto' in col_lower or 'monto' in col_lower:
                            try:
                                monto = float(valor)
                            except:
                                pass
                        elif 'total' in col_lower:
                            total = valor
                
                resultados[concepto].append({
                    'FECHA': fecha,
                    'REFERENCIA': referencia,
                    'DESCRIPCIÓN': descripcion,
                    'MONTO': abs(monto) if monto else 0,
                    'TOTAL': total if total else (abs(monto) if monto else 0)
                })
    
    return resultados

def separar_ingresos_egresos(df):
    """
    Separa ingresos y egresos por el signo del monto
    """
    ingresos = []
    egresos = []
    
    # Buscar la columna que contiene los montos
    columna_monto = None
    for col in df.columns:
        col_lower = str(col).lower()
        if 'monto' in col_lower or 'total' in col_lower:
            columna_monto = col
            break
    
    if columna_monto is None:
        st.error("No se encontró una columna de montos (MONTO o TOTAL)")
        return [], []
    
    for _, fila in df.iterrows():
        monto = fila[columna_monto]
        try:
            monto_num = float(monto) if pd.notna(monto) else 0
        except:
            continue
        
        # Obtener datos básicos
        fecha = ''
        referencia = ''
        descripcion = ''
        
        for col in df.columns:
            col_lower = str(col).lower()
            valor = fila[col]
            if pd.notna(valor):
                if 'fecha' in col_lower:
                    fecha = valor
                elif 'referencia' in col_lower or 'ref' in col_lower:
                    referencia = valor
                elif 'descripcion' in col_lower or 'concepto' in col_lower:
                    descripcion = str(valor)
        
        registro = {
            'FECHA': fecha,
            'REFERENCIA': referencia,
            'DESCRIPCIÓN': descripcion,
            'MONTO': abs(monto_num),
            'TOTAL': abs(monto_num)
        }
        
        if monto_num > 0:
            ingresos.append(registro)
        elif monto_num < 0:
            registro['CLASIFICACIÓN DE EGRESO'] = 'EGRESO'
            egresos.append(registro)
    
    return ingresos, egresos

# Área principal
if archivo:
    st.info(f"📄 Archivo: **{archivo.name}** - {archivo.size/1024:.1f} KB")
    
    try:
        df_original = pd.read_excel(archivo)
        
        with st.expander("👁️ Vista previa del archivo original (primeras 10 filas)"):
            st.dataframe(df_original.head(10), use_container_width=True)
            st.caption(f"Total filas: {len(df_original)} | Columnas: {list(df_original.columns)}")
        
        if procesar:
            with st.spinner('Analizando archivo...'):
                # 1. Separar ingresos y egresos
                ingresos, egresos = separar_ingresos_egresos(df_original)
                
                # 2. Clasificar por conceptos (comisiones en este caso)
                conceptos_buscar = ['comision', 'pago movil']
                clasificacion = clasificar_por_concepto(df_original, conceptos_buscar)
            
            st.success(f"✅ Procesado: {len(ingresos)} INGRESOS | {len(egresos)} EGRESOS | {len(clasificacion.get('comision', []))} COMISIONES")
            
            # Métricas
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("💰 INGRESOS", len(ingresos))
            with col2:
                st.metric("💸 EGRESOS", len(egresos))
            with col3:
                st.metric("📋 COMISIONES", len(clasificacion.get('comision', [])))
            
            # Crear Excel con múltiples hojas
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                # Hoja INGRESOS
                if ingresos:
                    pd.DataFrame(ingresos).to_excel(writer, sheet_name='INGRESOS', index=False)
                else:
                    st.warning("No se encontraron INGRESOS")
                
                # Hoja EGRESOS
                if egresos:
                    pd.DataFrame(egresos).to_excel(writer, sheet_name='EGRESOS', index=False)
                else:
                    st.warning("No se encontraron EGRESOS")
                
                # Hoja COMISIONES
                if clasificacion.get('comision'):
                    pd.DataFrame(clasificacion['comision']).to_excel(writer, sheet_name='COMISIONES', index=False)
                else:
                    st.warning("No se encontraron COMISIONES")
                
                # Hoja RESUMEN
                resumen = {
                    'TIPO': ['INGRESOS', 'EGRESOS', 'COMISIONES'],
                    'CANTIDAD': [len(ingresos), len(egresos), len(clasificacion.get('comision', []))],
                    'MONTO TOTAL': [
                        sum([i['MONTO'] for i in ingresos]),
                        sum([e['MONTO'] for e in egresos]),
                        sum([c['MONTO'] for c in clasificacion.get('comision', [])])
                    ]
                }
                pd.DataFrame(resumen).to_excel(writer, sheet_name='RESUMEN', index=False)
            
            output.seek(0)
            
            # Mostrar previsualización en pestañas
            st.subheader("📊 Resultados")
            tabs = st.tabs(["📈 INGRESOS", "📉 EGRESOS", "💳 COMISIONES"])
            
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
                if clasificacion.get('comision'):
                    st.dataframe(pd.DataFrame(clasificacion['comision']), use_container_width=True)
                else:
                    st.info("No hay COMISIONES")
            
            # Botón de descarga
            st.download_button(
                label="📥 Descargar Excel organizado (3 tablas + resumen)",
                data=output.getvalue(),
                file_name=f"balance_{archivo.name}",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
    
    except Exception as e:
        st.error(f"❌ Error: {str(e)}")

else:
    st.markdown("""
    ### 👋 ¡Bienvenido al Clasificador Bancario!
    
    **¿Qué hace este programa?**
    
    1. 📂 **Carga** un extracto bancario en Excel
    2. 🔍 **Detecta automáticamente** INGRESOS, EGRESOS y COMISIONES
    3. 📊 **Exporta** un Excel con 4 hojas separadas
    
    ### 📋 Hojas de salida
    | Hoja | Contenido |
    |------|-----------|
    | **INGRESOS** | Todos los depósitos, créditos y pagos recibidos |
    | **EGRESOS** | Todos los pagos, débitos y transferencias enviadas |
    | **COMISIONES** | Todas las transacciones que contengan "comision" |
    | **RESUMEN** | Totales consolidados |
    
    ---
    **💡 Tip:** El programa funciona con Banesco, Mercantil, Provincial, Venezuela y cualquier otro banco.
    """)

# Footer
st.markdown("---")
st.markdown(
    """
    <div class='footer'>
        <strong>Grupo Bodeguita Oriente</strong> - Clasificador Bancario v8.0
    </div>
    """,
    unsafe_allow_html=True
)
