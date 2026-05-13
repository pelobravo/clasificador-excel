import streamlit as st
import pandas as pd
import re
from io import BytesIO
from datetime import datetime

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

def es_fecha_valida(valor):
    """Determina si un valor parece una fecha"""
    if pd.isna(valor):
        return False
    valor_str = str(valor)
    # Patrones de fecha comunes
    patrones = [
        r'\d{2}/\d{2}/\d{4}',  # DD/MM/YYYY
        r'\d{2}-\d{2}-\d{4}',  # DD-MM-YYYY
        r'\d{4}/\d{2}/\d{2}',  # YYYY/MM/DD
        r'\d{2}\.\d{2}\.\d{4}', # DD.MM.YYYY
    ]
    for patron in patrones:
        if re.search(patron, valor_str):
            return True
    return False

def es_monto_valido(valor):
    """Determina si un valor parece un monto bancario"""
    if pd.isna(valor):
        return False
    try:
        # Intentar convertir a número
        if isinstance(valor, (int, float)):
            num = float(valor)
            # Los montos suelen ser mayores a 100 y menores a 1 billón
            return 100 < num < 999999999
        # Si es string, limpiar y convertir
        valor_str = str(valor).replace(',', '').replace('.', '').strip()
        if valor_str.isdigit():
            num = float(valor_str)
            return 100 < num < 999999999
    except:
        pass
    return False

def es_descripcion_valida(valor):
    """Determina si un valor parece una descripción de transacción"""
    if pd.isna(valor):
        return False
    valor_str = str(valor)
    # Descripciones suelen tener más de 10 caracteres y contener letras
    if len(valor_str) > 10 and re.search(r'[A-Za-zÁÉÍÓÚáéíóúÑñ]', valor_str):
        return True
    return False

def procesar_archivo(df):
    """Procesa el archivo buscando datos en toda la tabla"""
    
    ingresos = []
    egresos = []
    comisiones = []
    
    # Recorrer cada fila y cada columna buscando datos
    for idx, fila in df.iterrows():
        # Variables para esta fila
        fecha = ""
        descripcion = ""
        monto = 0
        tipo_movimiento = ""
        
        # Buscar en todas las columnas de esta fila
        for col in df.columns:
            valor = fila[col]
            if pd.isna(valor):
                continue
            
            valor_str = str(valor)
            valor_lower = valor_str.lower()
            
            # 1. Detectar tipo de movimiento (NC/ND)
            if 'nc' in valor_lower or 'ingreso' in valor_lower:
                tipo_movimiento = 'INGRESO'
            elif 'nd' in valor_lower or 'egreso' in valor_lower:
                tipo_movimiento = 'EGRESO'
            
            # 2. Detectar comisiones
            if 'comision' in valor_lower or 'comisión' in valor_lower:
                if not descripcion:
                    descripcion = valor_str
            
            # 3. Detectar fecha
            if es_fecha_valida(valor) and not fecha:
                fecha = valor_str
            
            # 4. Detectar monto (prioridad a números grandes)
            if es_monto_valido(valor):
                try:
                    # Limpiar el valor para obtener el número
                    if isinstance(valor, (int, float)):
                        monto_candidato = abs(float(valor))
                    else:
                        valor_limpio = str(valor).replace(',', '').replace('.', '').strip()
                        if valor_limpio.isdigit():
                            monto_candidato = float(valor_limpio)
                        else:
                            continue
                    
                    # Tomar el monto más grande encontrado (suele ser el correcto)
                    if monto_candidato > monto:
                        monto = monto_candidato
                except:
                    pass
            
            # 5. Detectar descripción (texto largo)
            if es_descripcion_valida(valor) and not descripcion:
                descripcion = valor_str
        
        # Si no encontramos descripción pero hay tipo de movimiento, buscar en toda la fila
        if not descripcion and (tipo_movimiento or 'comision' in str(fila.values).lower()):
            for col in df.columns:
                valor = fila[col]
                if pd.notna(valor) and len(str(valor)) > 10:
                    descripcion = str(valor)
                    break
        
        # Si tenemos datos válidos, guardar
        if monto > 0 or (tipo_movimiento and descripcion):
            registro = {
                'FECHA': fecha if fecha else '',
                'DESCRIPCIÓN': descripcion if descripcion else 'Sin descripción',
                'MONTO': monto
            }
            
            # Clasificar
            if 'comision' in descripcion.lower() or 'comisión' in descripcion.lower():
                comisiones.append(registro)
            elif tipo_movimiento == 'INGRESO':
                ingresos.append(registro)
            elif tipo_movimiento == 'EGRESO':
                registro['TIPO'] = 'EGRESO'
                egresos.append(registro)
            elif monto > 0:
                ingresos.append(registro)
            elif monto < 0:
                registro['TIPO'] = 'EGRESO'
                egresos.append(registro)
    
    return ingresos, egresos, comisiones

if archivo:
    st.info(f"📄 Archivo: **{archivo.name}** - {archivo.size/1024:.1f} KB")
    
    try:
        # Cargar el Excel sin encabezados para tener control total
        df_original = pd.read_excel(archivo, header=None)
        
        with st.expander("👁️ Vista previa del archivo (primeras 20 filas)"):
            st.dataframe(df_original.head(20), use_container_width=True)
        
        if procesar:
            with st.spinner('Procesando archivo...'):
                ingresos, egresos, comisiones = procesar_archivo(df_original)
            
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
                    pd.DataFrame(ingresos).to_excel(writer, sheet_name='INGRESOS', index=False)
                if egresos:
                    pd.DataFrame(egresos).to_excel(writer, sheet_name='EGRESOS', index=False)
                if comisiones:
                    pd.DataFrame(comisiones).to_excel(writer, sheet_name='COMISIONES', index=False)
                
                # Hoja de resumen
                resumen = pd.DataFrame([
                    {'TIPO': 'INGRESOS', 'CANTIDAD': len(ingresos), 'MONTO_TOTAL': sum(i.get('MONTO', 0) for i in ingresos)},
                    {'TIPO': 'EGRESOS', 'CANTIDAD': len(egresos), 'MONTO_TOTAL': sum(e.get('MONTO', 0) for e in egresos)},
                    {'TIPO': 'COMISIONES', 'CANTIDAD': len(comisiones), 'MONTO_TOTAL': sum(c.get('MONTO', 0) for c in comisiones)}
                ])
                resumen.to_excel(writer, sheet_name='RESUMEN', index=False)
            
            output.seek(0)
            
            # Mostrar previsualización
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
                label="📥 Descargar Excel (INGRESOS | EGRESOS | COMISIONES | RESUMEN)",
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
    
    **Instrucciones:**
    1. Carga el archivo Excel del banco
    2. Presiona "Procesar"
    3. Descarga el Excel con 3 hojas separadas
    
    **El programa busca automáticamente:**
    - 📅 **Fechas** (DD/MM/YYYY, DD-MM-YYYY, etc.)
    - 📝 **Descripciones** (textos largos con letras)
    - 💰 **Montos** (números entre 100 y 999M)
    - 🔄 **NC/ND** (para diferenciar Ingresos/Egresos)
    - 💳 **Comisiones** (palabra "comision" en la descripción)
    """)

st.markdown("---")
st.markdown('<div class="footer"><strong>Grupo Bodeguita Oriente</strong> - Clasificador v12.0</div>', unsafe_allow_html=True)
