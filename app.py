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
    
    /* Centrar contenido en sidebar */
    .sidebar-center {
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        text-align: center;
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

# Sidebar para configuración con logo centrado
with st.sidebar:
    # Logo centrado usando HTML/CSS
    st.markdown(
        """
        <div style="display: flex; justify-content: center; margin: 10px 0 10px 0;">
            <img src="https://raw.githubusercontent.com/pelobravo/clasificador-excel/main/LOGO.jpeg" width="100">
        </div>
        """,
        unsafe_allow_html=True
    )
    
    # Texto del grupo centrado
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
    
    # Palabras clave para conceptos
    conceptos_default = "pago movil, comision, transferencia, pago celular"
    conceptos = st.text_area(
        "📝 Conceptos a buscar (separados por coma)",
        conceptos_default,
        help="Ejemplo: pago movil, comision, transferencia. Cada concepto será una columna independiente."
    )
    
    # Palabras clave para bancos
    bancos_default = "venezuela, mercantil, provincial, banesco, bicentenario, banco de venezuela, banco caribbean, banco exterior"
    bancos = st.text_area(
        "🏦 Bancos a buscar (separados por coma)",
        bancos_default,
        help="Ejemplo: venezuela, mercantil, provincial"
    )
    
    st.markdown("---")
    
    # Botón de procesamiento
    procesar = st.button("🚀 Procesar y Clasificar", type="primary", use_container_width=True)

# NUEVA FUNCIÓN: Crea una columna por cada concepto
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
                partes.append(str(valor).lower())
        texto_completo = " ".join(partes)
        
        # --- DETECTAR CONCEPTOS (una columna por cada concepto) ---
        conceptos_detectados = {}
        for concepto in conceptos_list:
            # Buscar el concepto en el texto completo
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
        
        # Búsqueda más precisa en columnas específicas
        if not banco_encontrado:
            for col in df.columns:
                col_str = str(col).lower()
                if pd.notna(row[col]):
                    valor_str = str(row[col]).lower()
                    for banco in bancos_list:
                        if banco in valor_str:
                            banco_encontrado = banco.title()
                            break
                    if banco_encontrado:
                        break
        
        # --- DETECTAR MONTOS ---
        monto = None
        for col in df.columns:
            valor = row[col]
            if pd.notna(valor) and isinstance(valor, (int, float)) and valor > 0:
                if 'monto' in str(col).lower():
                    monto = valor
                    break
                elif 'monto' in str(col).lower():
                    monto = valor
                    break
                elif monto is None:
                    monto = valor
        
        # --- CONSTRUIR REGISTRO ---
        # Comenzar con los datos originales
        registro = {**row.to_dict()}
        
        # Agregar columnas de conceptos (una por cada concepto buscado)
        registro.update(conceptos_detectados)
        
        # Agregar banco detectado y monto
        registro['🏦 Banco Detectado'] = banco_encontrado if banco_encontrado else 'No detectado'
        registro['💰 Monto Detectado'] = monto if monto else 'No aplica'
        
        resultados.append(registro)
    
    return pd.DataFrame(resultados)

# Área principal - Mostrar resultados
if archivo:
    # Mostrar información del archivo cargado
    st.info(f"📄 Archivo cargado: **{archivo.name}** - Tamaño: {archivo.size/1024:.1f} KB")
    
    # Cargar el Excel
    try:
        df_original = pd.read_excel(archivo)
        
        # Mostrar vista previa
        with st.expander("👁️ Vista previa del archivo original (primeras 10 filas)"):
            st.dataframe(df_original.head(10), use_container_width=True)
            st.caption(f"Total de filas: {len(df_original)} | Columnas: {len(df_original.columns)}")
        
        # Procesamiento
        if procesar:
            # Verificar que hay palabras clave
            if not conceptos.strip() and not bancos.strip():
                st.error("❌ Debes ingresar al menos un concepto o banco para buscar")
            else:
                # Procesar listas
                conceptos_list = [c.strip().lower() for c in conceptos.split(',') if c.strip()]
                bancos_list = [b.strip().lower() for b in bancos.split(',') if b.strip()]
                
                # Mostrar qué se está buscando
                st.write("**Buscando:**")
                col1, col2 = st.columns(2)
                with col1:
                    st.write(f"📝 Conceptos ({len(conceptos_list)}): {', '.join(conceptos_list)}")
                    st.info(f"📌 Se crearán {len(conceptos_list)} columnas independientes, una por cada concepto")
                with col2:
                    st.write(f"🏦 Bancos: {', '.join(bancos_list)}")
                
                # Realizar clasificación con la NUEVA función
                with st.spinner('🔄 Clasificando datos... Espere un momento'):
                    df_resultados = buscar_coincidencias_columnas(df_original, conceptos_list, bancos_list)
                    st.session_state.df_procesado = df_resultados
                
                # Mostrar resultados
                if len(df_resultados) > 0:
                    st.success(f"✅ **¡Éxito!** Se procesaron {len(df_resultados)} registros")
                    
                    # Contar filas que tienen al menos un concepto marcado como 'Sí'
                    columnas_conceptos = [col for col in df_resultados.columns if col.startswith('📌')]
                    if columnas_conceptos:
                        total_coincidencias = 0
                        for col in columnas_conceptos:
                            total_coincidencias += (df_resultados[col] == 'Sí').sum()
                        st.metric("📊 Total de coincidencias entre conceptos", total_coincidencias)
                    
                    # Métricas adicionales
                    col1, col2, col3, col4 = st.columns(4)
                    with col1:
                        st.metric("📊 Total Registros", len(df_original))
                    with col2:
                        st.metric("✅ Registros con coincidencias", len(df_resultados))
                    with col3:
                        st.metric("📈 Porcentaje", f"{(len(df_resultados)/len(df_original)*100):.1f}%")
                    with col4:
                        st.metric("📌 Conceptos buscados", len(conceptos_list))
                    
                    # Gráfico de bancos encontrados
                    if '🏦 Banco Detectado' in df_resultados.columns:
                        st.subheader("📊 Distribución por Banco")
                        resumen_banco = df_resultados['🏦 Banco Detectado'].value_counts()
                        st.bar_chart(resumen_banco)
                    
                    # Tabla de resultados
                    st.subheader("📋 Detalle de registros clasificados")
                    st.dataframe(df_resultados, use_container_width=True, height=400)
                    
                    # Botones de descarga
                    col1, col2 = st.columns(2)
                    
                    # Descargar CSV
                    csv = df_resultados.to_csv(index=False).encode('utf-8')
                    with col1:
                        st.download_button(
                            label="📥 Descargar como CSV",
                            data=csv,
                            file_name=f"resultados_{archivo.name.replace('.xlsx', '').replace('.xls', '')}.csv",
                            mime="text/csv",
                            use_container_width=True
                        )
                    
                    # Descargar Excel
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
        st.error(f"❌ Error al leer el archivo: {str(e)}")
        st.info("Asegúrate de que el archivo sea un Excel válido (.xlsx o .xls)")

else:
    # Mensaje inicial
    st.markdown("""
    ### 👋 ¡Bienvenido a Grupo Bodeguita Oriente!
    
    **¿Cómo funciona este clasificador?**
    
    1. 📂 **Carga** un archivo Excel desde la barra lateral izquierda
    2. 🔍 **Configura** qué conceptos y bancos quieres buscar
    3. 🚀 **Presiona** "Procesar y Clasificar"
    4. 📊 **Obtén** resultados filtrados y estadísticas
    
    ### 🆕 Novedad - Columnas por concepto!
    Ahora cada concepto que buscas se convierte en una columna independiente.
    - Si buscas **"pago movil"** y **"comision"**, tendrás dos columnas separadas:
      - 📌 PAGO MOVIL (con valores "Sí" o "No")
      - 📌 COMISION (con valores "Sí" o "No")
    
    ### Ejemplos de búsqueda:
    - **Conceptos:** pago movil, comision, transferencia
    - **Bancos:** venezuela, mercantil, provincial, banesco
    
    ---
    **💡 Tip:** El programa buscará estas palabras en TODAS las columnas de tu archivo Excel
    """)

# Footer personalizado
st.markdown("---")
st.markdown(
    """
    <div class="footer">
        <strong>Grupo Bodeguita Oriente</strong> - Clasificador de Excel v3.0<br>
        Sistema de clasificación de pagos y transferencias bancarias<br>
        Desarrollado con Streamlit
    </div>
    """,
    unsafe_allow_html=True
)
