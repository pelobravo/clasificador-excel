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
    st.markdown("### Grupo Bodeguita Oriente - Detección Automática")
st.markdown("---")

with st.sidebar:
    st.image("https://raw.githubusercontent.com/pelobravo/clasificador-excel/main/LOGO.jpeg", width=100)
    st.markdown("---")
    archivo = st.file_uploader("📂 Cargar archivo Excel", type=['xlsx', 'xls'])
    st.markdown("---")
    st.caption("El programa detecta automáticamente: Banco, Fechas, Montos y tipo de movimiento")
    procesar = st.button("🚀 Procesar Automáticamente", type="primary", use_container_width=True)

def detectar_banco(df):
    """Detecta el banco basado en el contenido del archivo"""
    texto_completo = " ".join([str(v).lower() for v in df.values.flatten() if pd.notna(v)])
    
    if 'mercantil' in texto_completo:
        return 'MERCANTIL'
    elif 'banesco' in texto_completo:
        return 'BANESCO'
    elif 'provincial' in texto_completo:
        return 'PROVINCIAL'
    elif 'venezuela' in texto_completo or 'bdv' in texto_completo:
        return 'VENEZUELA'
    else:
        return 'GENERICO'

def encontrar_columna_datos(df):
    """
    Encuentra la fila donde comienzan los datos reales
    Busca patrones como: NC, ND, fechas numéricas, montos
    """
    for idx, fila in df.iterrows():
        texto_fila = " ".join([str(v) for v in fila.values if pd.notna(v)]).upper()
        if 'NC' in texto_fila or 'ND' in texto_fila:
            return idx
        # También buscar si hay números que parecen fechas (DDMMYYYY o D/M/YYYY)
        match = re.search(r'\d{2}\d{2}\d{4}|\d{1,2}/\d{1,2}/\d{2,4}', texto_fila)
        if match and len(texto_fila) > 20:
            return idx
    return 0

def clasificar_automaticamente(df, fila_inicio):
    """
    Clasifica automáticamente ingresos, egresos y comisiones
    """
    # Crear nuevo DataFrame desde la fila de inicio
    df_datos = df.iloc[fila_inicio:].reset_index(drop=True)
    
    ingresos = []
    egresos = []
    comisiones = []
    
    # Buscar columnas que parecen contener fechas, descripciones y montos
    col_fecha = None
    col_desc = None
    col_monto = None
    
    # Analizar primera fila de datos para identificar columnas
    primera_fila = df_datos.iloc[0] if len(df_datos) > 0 else None
    
    if primera_fila is not None:
        for i, valor in enumerate(primera_fila):
            valor_str = str(valor).upper()
            
            # Detectar columna de NC/ND (tipo de movimiento)
            if 'NC' in valor_str or 'ND' in valor_str:
                col_tipo = i
            
            # Detectar columna de fecha (patrón de números: DDMMYYYY o DD/MM/YYYY)
            if re.search(r'\d{2}\d{2}\d{4}|\d{1,2}/\d{1,2}/\d{2,4}', valor_str):
                if col_fecha is None:
                    col_fecha = i
            
            # Detectar columna de descripción (texto largo, mayor a 10 caracteres)
            if len(valor_str) > 15 and not re.search(r'\d', valor_str.replace('/', '').replace('-', '')):
                if col_desc is None:
                    col_desc = i
    
    # Si no se encontró columna de fecha, buscar en todas las filas
    if col_fecha is None:
        for fila in df_datos.iterrows():
            for i, valor in enumerate(fila[1]):
                valor_str = str(valor)
                if re.search(r'\d{2}\d{2}\d{4}', valor_str) and len(valor_str) == 8:
                    col_fecha = i
                    break
            if col_fecha is not None:
                break
    
    # Detectar columna de monto (la que tenga números grandes o decimales)
    for i in range(len(df_datos.columns)):
        valores_numericos = 0
        for fila in df_datos.iterrows():
            valor = fila[1][i]
            if pd.notna(valor):
                try:
                    num = float(str(valor).replace(',', '.'))
                    if num > 0 and num < 999999999:
                        valores_numericos += 1
                except:
                    pass
        if valores_numericos > len(df_datos) * 0.5:  # Más del 50% son números
            col_monto = i
            break
    
    st.info(f"🔍 Detección automática: Fecha columna {col_fecha}, Monto columna {col_monto}, Descripción columna {col_desc}")
    
    # Procesar cada fila
    for _, fila in df_datos.iterrows():
        # Obtener valores
        fecha = fila[col_fecha] if col_fecha is not None and col_fecha < len(fila) else ''
        descripcion = fila[col_desc] if col_desc is not None and col_desc < len(fila) else ''
        monto_valor = fila[col_monto] if col_monto is not None and col_monto < len(fila) else 0
        
        # Convertir monto a número
        try:
            monto_num = float(str(monto_valor).replace(',', '.').replace(' ', ''))
        except:
            monto_num = 0
        
        # Detectar tipo por NC/ND
        tipo_fila = " ".join([str(v).upper() for v in fila.values if pd.notna(v)])
        
        registro = {
            'FECHA': fecha,
            'DESCRIPCIÓN': descripcion,
            'MONTO': abs(monto_num)
        }
        
        # Clasificar
        if 'COMISION' in tipo_fila:
            comisiones.append(registro)
        elif 'NC' in tipo_fila or monto_num > 0:
            ingresos.append(registro)
        elif 'ND' in tipo_fila or monto_num < 0:
            registro['CLASIFICACIÓN'] = 'EGRESO'
            egresos.append(registro)
    
    return ingresos, egresos, comisiones

# Área principal
if archivo:
    st.info(f"📄 Archivo: **{archivo.name}** - {archivo.size/1024:.1f} KB")
    
    try:
        df_original = pd.read_excel(archivo, header=None)  # Leer sin encabezados
        
        with st.expander("👁️ Vista previa del archivo (primeras 20 filas)"):
            st.dataframe(df_original.head(20), use_container_width=True)
        
        if procesar:
            with st.spinner('Analizando archivo automáticamente...'):
                # Detectar banco
                banco = detectar_banco(df_original)
                st.info(f"🏦 Banco detectado: **{banco}**")
                
                # Encontrar fila donde comienzan los datos
                fila_inicio = encontrar_columna_datos(df_original)
                st.info(f"📌 Datos comienzan en fila: {fila_inicio + 1}")
                
                # Clasificar
                ingresos, egresos, comisiones = clasificar_automaticamente(df_original, fila_inicio)
            
            st.success(f"✅ Procesado: {len(ingresos)} INGRESOS | {len(egresos)} EGRESOS | {len(comisiones)} COMISIONES")
            
            # Métricas
            col1, col2, col3 = st.columns(3)
            col1.metric("💰 INGRESOS", len(ingresos))
            col2.metric("💸 EGRESOS", len(egresos))
            col3.metric("💳 COMISIONES", len(comisiones))
            
            # Crear Excel de salida
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
                    {'TIPO': 'INGRESOS', 'CANTIDAD': len(ingresos), 'MONTO TOTAL': sum(i.get('MONTO', 0) for i in ingresos)},
                    {'TIPO': 'EGRESOS', 'CANTIDAD': len(egresos), 'MONTO TOTAL': sum(e.get('MONTO', 0) for e in egresos)},
                    {'TIPO': 'COMISIONES', 'CANTIDAD': len(comisiones), 'MONTO TOTAL': sum(c.get('MONTO', 0) for c in comisiones)}
                ])
                resumen.to_excel(writer, sheet_name='RESUMEN', index=False)
            
            output.seek(0)
            
            # Mostrar previsualización
            tabs = st.tabs(["INGRESOS", "EGRESOS", "COMISIONES"])
            with tabs[0]:
                if ingresos:
                    st.dataframe(pd.DataFrame(ingresos), use_container_width=True)
                else:
                    st.info("No se encontraron INGRESOS")
            with tabs[1]:
                if egresos:
                    st.dataframe(pd.DataFrame(egresos), use_container_width=True)
                else:
                    st.info("No se encontraron EGRESOS")
            with tabs[2]:
                if comisiones:
                    st.dataframe(pd.DataFrame(comisiones), use_container_width=True)
                else:
                    st.info("No se encontraron COMISIONES")
            
            st.download_button(
                label="📥 Descargar Excel organizado",
                data=output.getvalue(),
                file_name=f"balance_{banco}_{archivo.name}",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
    
    except Exception as e:
        st.error(f"❌ Error: {str(e)}")

else:
    st.markdown("""
    ### 👋 Clasificador Bancario Automático
    
    **Simplemente:**
    1. 📂 Carga tu extracto bancario
    2. 🚀 Presiona "Procesar Automáticamente"
    3. 📊 Obtén INGRESOS, EGRESOS y COMISIONES
    
    **El programa detecta todo automáticamente:**
    - 🔍 El banco (Banesco, Mercantil, Provincial, etc.)
    - 📅 Las fechas y montos
    - 💳 Las comisiones
    - ⬆️⬇️ Los ingresos y egresos por NC/ND
    
    **Sin configurar nada. Sin seleccionar columnas.**
    """)

st.markdown("---")
st.markdown('<div class="footer"><strong>Grupo Bodeguita Oriente</strong> - Clasificador Automático v10.0</div>', unsafe_allow_html=True)
