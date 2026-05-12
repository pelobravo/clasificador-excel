import streamlit as st
import pandas as pd
import re
from io import BytesIO

# Configuración de la página
st.set_page_config(
    page_title="Clasificador de Excel - Grupo Bodeguita Oriente",
    page_icon="🏦",
    layout="wide"
)

# Personalización de colores y estilos (CSS)
st.markdown("""
    <style>
    /* Color de fondo principal */
    .stApp {
        background-color: #ffffff;
    }
    
    /* Botón principal */
    .stButton > button {
        background-color: #1e3a5f;
        color: white;
        border-radius: 8px;
        padding: 10px 24px;
        font-weight: bold;
        border: none;
        transition: all 0.3s ease;
    }
    .stButton > button:hover {
        background-color: #2c5282;
        transform: translateY(-2px);
        box-shadow: 0 4px 12px rgba(0,0,0,0.15);
    }
    
    /* Títulos */
    h1, h2, h3 {
        color: #1e3a5f;
    }
    
    /* Sidebar */
    .css-1d391kg {
        background-color: #ffffff;
    }
    
    /* Tarjetas de métricas */
    .stMetric {
        background-color: white;
        border-radius: 12px;
        padding: 15px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.05);
    }
    
    /* Expander */
    .streamlit-expanderHeader {
        background-color: #ffffff;
        border-radius: 8px;
    }
    
    /* Dataframe */
    .stDataFrame {
        border-radius: 12px;
        overflow: hidden;
    }
    
    /* Footer */
    .footer {
        text-align: center;
        color: #666;
        padding: 20px;
        font-size: 14px;
    }
    </style>
""", unsafe_allow_html=True)

# Título principal con logo
col_logo, col_titulo = st.columns([1, 5])
with col_logo:
    try:
        st.image("LOGO.jpeg", width=80)
    except:
        st.image("https://via.placeholder.com/80?text=GBO", width=80)
with col_titulo:
    st.title("Clasificador de Excel - Bancos y Conceptos")
    st.markdown("### Clasifica automáticamente pagos móviles, transferencias y bancos")
st.markdown("---")

# Inicializar estado de sesión
if 'df_procesado' not in st.session_state:
    st.session_state.df_procesado = None

# Sidebar para configuración
with st.sidebar:
    st.markdown(
        """
        <div style="display: flex; justify-content: center; margin: 10px 0 10px 0;">
            <img src="https://raw.githubusercontent.com/pelobravo/clasificador-excel/main/LOGO.jpeg" width="100">
        </div>
        """,
        unsafe_allow_html=True
    )
    
    st.markdown(
        """
        <div style="text-align: center; margin-bottom: 20px;">
            <strong style="font-size: 16px; color: #1e3a5f;">Grupo Bodeguita Oriente</strong>
        </div>
        """,
        unsafe_allow_html=True
    )
    
    st.markdown("---")
    
    st.header("📂 Cargar Archivo")
    archivo = st.file_uploader(
        "Selecciona un archivo Excel",
        type=['xlsx', 'xls'],
        help="Formatos soportados: .xlsx, .xls"
    )
    
    st.markdown("---")
    
    st.header("🔍 Configuración de Búsqueda")
    
    conceptos_default = "pago movil, comision, transferencia, pago celular"
    conceptos = st.text_area(
        "📝 Conceptos a buscar (separados por coma)",
        conceptos_default,
        help="Ejemplo: pago movil, comision, transferencia. Cada concepto será una columna independiente."
    )
    
    bancos_default = "venezuela, mercantil, provincial, banesco, bicentenario, banco de venezuela, banco caribbean, banco exterior"
    bancos = st.text_area(
        "🏦 Bancos a buscar (separados por coma)",
        bancos_default,
        help="Ejemplo: venezuela, mercantil, provincial"
    )
    
    st.markdown("---")
    
    procesar = st.button("🚀 Procesar y Clasificar", type="primary", use_container_width=True)

# FUNCIÓN CORREGIDA - Sin errores de startswith
def buscar_coincidencias_columnas(df, conceptos_list, bancos_list):
    """
    Busca conceptos y bancos en todas las columnas del dataframe.
    Crea una columna independiente para CADA concepto buscado.
    """
    resultados = []
    
    for idx, row in df.iterrows():
        # Crear una lista segura con todas las celdas convertidas a string
        partes = []
        for valor in row.values:
            if pd.notna(valor):
                try:
                    partes.append(str(valor).lower())
                except:
                    partes.append("")
        texto_completo = " ".join(partes)
        
        # --- DETECTAR CONCEPTOS (una columna por cada concepto) ---
        conceptos_detectados = {}
        for concepto in conceptos_list:
            if concepto in texto_completo:
                conceptos_detectados[f'📌 {concepto.upper()}'] = 'Sí'
            else:
                conceptos_detectados[f'📌 {concepto.upper()}'] = 'No'
        
        # --- DETECTAR BANCOS ---
        banco_encontrado = None
        for banco in bancos_list:
            if banco in texto_completo:
                banco_encontrado = banco.title()
                break
        
        # Búsqueda más precisa en columnas específicas (evitando errores de tipo)
        if not banco_encontrado:
            for col in df.columns:
                try:
                    valor = row[col]
                    if pd.notna(valor):
                        valor_str = str(valor).lower()
                        for banco in bancos_list:
                            if banco in valor_str:
                                banco_encontrado = banco.title()
                                break
                        if banco_encontrado:
                            break
                except:
                    continue
        
        # --- DETECTAR MONTOS ---
        monto = None
        for col in df.columns:
            try:
                valor = row[col]
                if pd.notna(valor) and isinstance(valor, (int, float)) and valor > 0:
                    # Convertir el nombre de la columna a string SIEMPRE
                    nombre_col = str(col).lower()
                    if 'monto' in nombre_col:
                        monto = valor
                        break
                    elif monto is None:
                        monto = valor
            except:
                continue
        
        # --- CONSTRUIR REGISTRO ---
        registro = {**row.to_dict()}
        registro.update(conceptos_detectados)
        registro['🏦 Banco Detectado'] = banco_encontrado if banco_encontrado else 'No detectado'
        registro['💰 Monto Detectado'] = monto if monto else 'No aplica'
        
        # Solo guardar si tiene al menos un concepto detectado como 'Sí' o un banco detectado
        tiene_concepto = any(conceptos_detectados.values())
        if tiene_concepto or banco_encontrado:
            resultados.append(registro)
    
    if len(resultados) == 0:
        return pd.DataFrame()
    else:
        return pd.DataFrame(resultados)

# Área principal - Mostrar resultados
if archivo:
    st.info(f"📄 Archivo cargado: **{archivo.name}** - Tamaño: {archivo.size/1024:.1f} KB")
    
    try:
        df_original = pd.read_excel(archivo)
        
        with st.expander("👁️ Vista previa del archivo original (primeras 10 filas)"):
            st.dataframe(df_original.head(10), use_container_width=True)
            st.caption(f"Total de filas: {len(df_original)} | Columnas: {len(df_original.columns)}")
        
        if procesar:
            if not conceptos.strip() and not bancos.strip():
                st.error("❌ Debes ingresar al menos un concepto o banco para buscar")
            else:
                conceptos_list = [c.strip().lower() for c in conceptos.split(',') if c.strip()]
                bancos_list = [b.strip().lower() for b in bancos.split(',') if b.strip()]
                
                st.write("**Buscando:**")
                col1, col2 = st.columns(2)
                with col1:
                    st.write(f"📝 Conceptos ({len(conceptos_list)}): {', '.join(conceptos_list)}")
                    st.info(f"📌 Se crearán {len(conceptos_list)} columnas independientes, una por cada concepto")
                with col2:
                    st.write(f"🏦 Bancos: {', '.join(bancos_list)}")
                
                with st.spinner('🔄 Clasificando datos... Espere un momento'):
                    df_resultados = buscar_coincidencias_columnas(df_original, conceptos_list, bancos_list)
                    st.session_state.df_procesado = df_resultados
                
                if len(df_resultados) > 0:
                    st.success(f"✅ **¡Éxito!** Se procesaron {len(df_resultados)} registros")
                    
                    # Contar coincidencias de conceptos
                    columnas_conceptos = []
                    for col in df_resultados.columns:
                        if col.startswith('📌'):
                            columnas_conceptos.append(col)
                    
                    if columnas_conceptos:
                        total_coincidencias = 0
                        for col in columnas_conceptos:
                            total_coincidencias += (df_resultados[col] == 'Sí').sum()
                        st.metric("📊 Total de coincidencias entre conceptos", total_coincidencias)
                    
                    col1, col2, col3, col4 = st.columns(4)
                    with col1:
                        st.metric("📊 Total Registros", len(df_original))
                    with col2:
                        st.metric("✅ Registros con coincidencias", len(df_resultados))
                    with col3:
                        porcentaje = (len(df_resultados)/len(df_original)*100) if len(df_original) > 0 else 0
                        st.metric("📈 Porcentaje", f"{porcentaje:.1f}%")
                    with col4:
                        st.metric("📌 Conceptos buscados", len(conceptos_list))
                    
                    if '🏦 Banco Detectado' in df_resultados.columns:
                        st.subheader("📊 Distribución por Banco")
                        resumen_banco = df_resultados['🏦 Banco Detectado'].value_counts()
                        st.bar_chart(resumen_banco)
                    
                    st.subheader("📋 Detalle de registros clasificados")
                    st.dataframe(df_resultados, use_container_width=True, height=400)
                    
                    col1, col2 = st.columns(2)
                    csv = df_resultados.to_csv(index=False).encode('utf-8')
                    with col1:
                        st.download_button(
                            label="📥 Descargar como CSV",
                            data=csv,
                            file_name=f"resultados_{archivo.name.replace('.xlsx', '').replace('.xls', '')}.csv",
                            mime="text/csv",
                            use_container_width=True
                        )
                    
                    with col2:
                        output = BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            df_resultados.to_excel(writer, sheet_name='Resultados', index=False)
                        output.seek(0)
                        st.download_button(
                            label="📊 Descargar como Excel",
                            data=output,
                            file_name=f"resultados_{archivo.name}",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                else:
                    st.warning("⚠️ No se encontraron coincidencias con los criterios de búsqueda")
                    st.info("💡 Sugerencia: Revisa que los conceptos y bancos escritos coincidan con la ortografía de tu archivo")
    
    except Exception as e:
        st.error(f"❌ Error al procesar el archivo: {str(e)}")
        st.info("Asegúrate de que el archivo sea un Excel válido (.xlsx o .xls)")

else:
    st.markdown("""
    ### 👋 ¡Bienvenido a Grupo Bodeguita Oriente!
    
    **¿Cómo funciona este clasificador?**
    
    1. 📂 **Carga** un archivo Excel desde la barra lateral izquierda
    2. 🔍 **Configura** qué conceptos y bancos quieres buscar
    3. 🚀 **Presiona** "Procesar y Clasificar"
    
    ### 🆕 Novedad - Columnas por concepto!
    Ahora cada concepto que buscas se convierte en una columna independiente.
    
    ### Ejemplos de búsqueda:
    - **Conceptos:** pago movil, comision, transferencia
    - **Bancos:** venezuela, mercantil, provincial, banesco
    """)

# Footer
st.markdown("---")
st.markdown(
    """
    <div class="footer">
        <strong>Grupo Bodeguita Oriente</strong> - Clasificador de Excel v3.2<br>
        Sistema de clasificación de pagos y transferencias bancarias
    </div>
    """,
    unsafe_allow_html=True
)
