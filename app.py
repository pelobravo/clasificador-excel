import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="Clasificador Bancario - Grupo Bodeguita Oriente", page_icon="🏦", layout="wide")

st.markdown("""
    <style>
    .stApp { background-color: #ffffff; }
    .stButton > button { background-color: #1e3a5f; color: white; border-radius: 8px; padding: 10px 24px; font-weight: bold; border: none; }
    .stButton > button:hover { background-color: #2c5282; }
    h1, h2, h3 { color: #1e3a5f; }
    .footer { text-align: center; color: #666; padding: 20px; font-size: 14px; }
    </style>
""", unsafe_allow_html=True)

# Título con logo
col_logo, col_titulo = st.columns([1, 5])
with col_logo:
    try:
        st.image("LOGO.jpeg", width=80)
    except:
        st.image("https://raw.githubusercontent.com/pelobravo/clasificador-excel/main/LOGO.jpeg", width=80)
with col_titulo:
    st.title("Clasificador Bancario")
    st.markdown("### Grupo Bodeguita Oriente")
st.markdown("---")

with st.sidebar:
    st.image("https://raw.githubusercontent.com/pelobravo/clasificador-excel/main/LOGO.jpeg", width=100)
    st.markdown("---")
    archivo = st.file_uploader("📂 Cargar archivo Excel", type=['xlsx', 'xls'])
    st.markdown("---")
    procesar = st.button("🚀 Procesar", type="primary", use_container_width=True)

def procesar_completo(df):
    """
    Procesa el archivo y extrae: INGRESOS, EGRESOS y COMISIONES
    con descripción real y montos correctos
    """
    # Convertir todo a string para análisis
    df_str = df.astype(str)
    
    ingresos = []
    egresos = []
    comisiones = []
    
    # Buscar la columna que contiene descripciones largas (texto real)
    columna_descripcion = None
    for col in df.columns:
        # Buscar columna con textos largos (descripciones reales)
        textos_largos = 0
        for valor in df[col].head(20):
            if pd.notna(valor) and isinstance(valor, str) and len(str(valor)) > 15:
                textos_largos += 1
        if textos_largos > 5:  # Si tiene muchos textos largos, es la columna de descripción
            columna_descripcion = col
            break
    
    # Si no se encontró, buscar cualquier columna con texto
    if columna_descripcion is None:
        for col in df.columns:
            if df[col].dtype == 'object':
                columna_descripcion = col
                break
    
    st.info(f"📝 Columna de descripción detectada: {columna_descripcion}")
    
    # Buscar la columna de montos (números grandes)
    columna_monto = None
    for col in df.columns:
        valores_numericos = 0
        for valor in df[col].head(50):
            if pd.notna(valor):
                try:
                    num = float(str(valor).replace(',', '').replace('.', ''))
                    if 100 < num < 999999999:  # Rango típico de montos
                        valores_numericos += 1
                except:
                    pass
        if valores_numericos > 10:
            columna_monto = col
            break
    
    st.info(f"💰 Columna de montos detectada: {columna_monto}")
    
    # Buscar columna de fecha
    columna_fecha = None
    for col in df.columns:
        for valor in df[col].head(20):
            if pd.notna(valor):
                valor_str = str(valor)
                if re.search(r'\d{2}/\d{2}/\d{4}|\d{2}-\d{2}-\d{4}|\d{6}', valor_str):
                    columna_fecha = col
                    break
        if columna_fecha:
            break
    
    # Procesar cada fila
    for idx, fila in df.iterrows():
        # Obtener descripción REAL (texto largo)
        descripcion = ""
        if columna_descripcion and pd.notna(fila[columna_descripcion]):
            descripcion = str(fila[columna_descripcion])
        
        # Si la descripción es muy corta o es un número, buscar en otras columnas
        if len(descripcion) < 10 or descripcion.isdigit():
            for col in df.columns:
                valor = fila[col]
                if pd.notna(valor) and isinstance(valor, str) and len(str(valor)) > 10:
                    descripcion = str(valor)
                    break
        
        # Obtener monto
        monto = 0
        if columna_monto and pd.notna(fila[columna_monto]):
            try:
                monto_str = str(fila[columna_monto]).replace(',', '').replace(' ', '')
                monto = float(monto_str)
            except:
                monto = 0
        
        # Obtener fecha
        fecha = ""
        if columna_fecha and pd.notna(fila[columna_fecha]):
            fecha = str(fila[columna_fecha])
        
        # Determinar tipo de movimiento
        texto_completo = " ".join([str(v) for v in fila.values if pd.notna(v)]).upper()
        
        # Buscar NC (Nota de Crédito = Ingreso), ND (Nota de Débito = Egreso)
        es_ingreso = False
        es_egreso = False
        es_comision = False
        
        if 'COMISION' in texto_completo or 'COMISIÓN' in texto_completo:
            es_comision = True
        elif 'NC' in texto_completo or 'INGRESO' in texto_completo or monto > 0:
            es_ingreso = True
        elif 'ND' in texto_completo or 'EGRESO' in texto_completo or monto < 0:
            es_egreso = True
        elif monto > 0:
            es_ingreso = True
        elif monto < 0:
            es_egreso = True
        
        # Crear registro
        registro = {
            'FECHA': fecha,
            'DESCRIPCIÓN': descripcion,
            'MONTO': abs(monto)
        }
        
        # Clasificar
        if es_comision:
            comisiones.append(registro)
        elif es_ingreso:
            ingresos.append(registro)
        elif es_egreso:
            registro['TIPO'] = 'EGRESO'
            egresos.append(registro)
    
    return ingresos, egresos, comisiones

if archivo:
    st.info(f"📄 Archivo: **{archivo.name}** - {archivo.size/1024:.1f} KB")
    
    try:
        # Cargar el Excel
        df_original = pd.read_excel(archivo)
        
        with st.expander("👁️ Vista previa del archivo (primeras 15 filas)"):
            st.dataframe(df_original.head(15), use_container_width=True)
        
        if procesar:
            with st.spinner('Procesando archivo...'):
                ingresos, egresos, comisiones = procesar_completo(df_original)
            
            st.success(f"✅ Procesado: {len(ingresos)} INGRESOS | {len(egresos)} EGRESOS | {len(comisiones)} COMISIONES")
            
            # Métricas
            col1, col2, col3 = st.columns(3)
            col1.metric("💰 INGRESOS", len(ingresos))
            col2.metric("💸 EGRESOS", len(egresos))
            col3.metric("💳 COMISIONES", len(comisiones))
            
            # Crear Excel con 3 hojas
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                if ingresos:
                    df_ingresos = pd.DataFrame(ingresos)
                    df_ingresos.to_excel(writer, sheet_name='INGRESOS', index=False)
                
                if egresos:
                    df_egresos = pd.DataFrame(egresos)
                    df_egresos.to_excel(writer, sheet_name='EGRESOS', index=False)
                
                if comisiones:
                    df_comisiones = pd.DataFrame(comisiones)
                    df_comisiones.to_excel(writer, sheet_name='COMISIONES', index=False)
                
                # Hoja de resumen
                resumen = pd.DataFrame([
                    {'TIPO': 'INGRESOS', 'CANTIDAD': len(ingresos), 'MONTO_TOTAL': sum(i.get('MONTO', 0) for i in ingresos)},
                    {'TIPO': 'EGRESOS', 'CANTIDAD': len(egresos), 'MONTO_TOTAL': sum(e.get('MONTO', 0) for e in egresos)},
                    {'TIPO': 'COMISIONES', 'CANTIDAD': len(comisiones), 'MONTO_TOTAL': sum(c.get('MONTO', 0) for c in comisiones)}
                ])
                resumen.to_excel(writer, sheet_name='RESUMEN', index=False)
            
            output.seek(0)
            
            # Mostrar previsualización en pestañas
            st.subheader("📊 Resultados")
            tab1, tab2, tab3 = st.tabs(["📈 INGRESOS", "📉 EGRESOS", "💳 COMISIONES"])
            
            with tab1:
                if ingresos:
                    st.dataframe(pd.DataFrame(ingresos), use_container_width=True)
                else:
                    st.info("No se encontraron INGRESOS")
            
            with tab2:
                if egresos:
                    st.dataframe(pd.DataFrame(egresos), use_container_width=True)
                else:
                    st.info("No se encontraron EGRESOS")
            
            with tab3:
                if comisiones:
                    st.dataframe(pd.DataFrame(comisiones), use_container_width=True)
                else:
                    st.info("No se encontraron COMISIONES")
            
            # Botón de descarga
            st.download_button(
                label="📥 Descargar Excel (INGRESOS, EGRESOS, COMISIONES, RESUMEN)",
                data=output.getvalue(),
                file_name=f"balance_{archivo.name}",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
    
    except Exception as e:
        st.error(f"❌ Error: {str(e)}")

else:
    st.markdown("""
    ### 👋 Clasificador Bancario
    
    **Carga el archivo y presiona Procesar**
    
    El programa extrae automáticamente:
    - ✅ **INGRESOS** (Notas de Crédito, montos positivos)
    - ✅ **EGRESOS** (Notas de Débito, montos negativos)
    - ✅ **COMISIONES** (transacciones que contengan "comision")
    
    **Luego puedes descargar un EXCEL con 3 hojas separadas**
    """)

st.markdown("---")
st.markdown('<div class="footer"><strong>Grupo Bodeguita Oriente</strong> - Clasificador v11.0</div>', unsafe_allow_html=True)
