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
        transition: all 0.3s ease;
    }
    .stButton > button:hover {
        background-color: #2c5282;
        transform: translateY(-2px);
        box-shadow: 0 4px 12px rgba(0,0,0,0.15);
    }
    h1, h2, h3 { color: #1e3a5f; }
    .stMetric {
        background-color: white;
        border-radius: 12px;
        padding: 15px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.05);
    }
    .stDataFrame { border-radius: 12px; overflow: hidden; }
    .footer { text-align: center; color: #666; padding: 20px; font-size: 14px; }
    </style>
""", unsafe_allow_html=True)

# Título
col_logo, col_titulo = st.columns([1, 5])
with col_logo:
    try:
        st.image("LOGO.jpeg", width=80)
    except:
        st.image("https://via.placeholder.com/80?text=GBO", width=80)
with col_titulo:
    st.title("Clasificador de Excel - Conceptos")
    st.markdown("### Filtra y clasifica pagos según los conceptos que elijas")
st.markdown("---")

# Sidebar
with st.sidebar:
    st.markdown(
        """
        <div style="display: flex; justify-content: center; margin: 10px 0;">
            <img src="https://raw.githubusercontent.com/pelobravo/clasificador-excel/main/LOGO.jpeg" width="100">
        </div>
        <div style="text-align: center; margin-bottom: 20px;">
            <strong style="font-size: 16px; color: #1e3a5f;">Grupo Bodeguita Oriente</strong>
        </div>
        """,
        unsafe_allow_html=True
    )
    st.markdown("---")
    
    archivo = st.file_uploader("📂 Cargar Archivo Excel", type=['xlsx', 'xls'])
    st.markdown("---")
    
    conceptos = st.text_area(
        "📝 Conceptos a buscar (separados por coma)",
        value="pago movil, comision",
        help="Ejemplo: pago movil, comision. Solo se mostrarán las filas que contengan estos conceptos."
    )
    st.markdown("---")
    
    procesar = st.button("🚀 Procesar y Clasificar", type="primary", use_container_width=True)

# Función principal - SOLO guarda filas con coincidencias
def procesar_archivo(df, lista_conceptos):
    """
    Procesa el archivo y SOLO guarda las filas que tienen al menos un concepto detectado
    """
    resultados = []
    
    for _, row in df.iterrows():
        # Unir todo el texto de la fila
        texto_fila = " ".join([str(v).lower() for v in row.values if pd.notna(v)])
        
        # Detectar conceptos
        conceptos_detectados = {}
        for concepto in lista_conceptos:
            if concepto in texto_fila:
                conceptos_detectados[f'{concepto.upper()}'] = concepto
            else:
                conceptos_detectados[f'{concepto.upper()}'] = 'No detectado'
        
        # Verificar si tiene al menos un concepto detectado (no todos "No detectado")
        tiene_algun_concepto = any(valor != 'No detectado' for valor in conceptos_detectados.values())
        
        # SOLO guardar si tiene al menos un concepto detectado
        if tiene_algun_concepto:
            # Detectar monto
            monto = None
            for col in df.columns:
                valor = row[col]
                if pd.notna(valor) and isinstance(valor, (int, float)) and valor > 0:
                    if monto is None:
                        monto = valor
                    if 'monto' in str(col).lower():
                        monto = valor
                        break
            
            # Armar registro
            registro = {**row.to_dict()}
            registro.update(conceptos_detectados)
            registro['💰 Monto'] = monto if monto else 'No aplica'
            resultados.append(registro)
    
    if len(resultados) == 0:
        return pd.DataFrame()
    else:
        return pd.DataFrame(resultados)

# Área principal
if archivo:
    st.info(f"📄 Archivo: **{archivo.name}** - {archivo.size/1024:.1f} KB")
    
    try:
        df_original = pd.read_excel(archivo)
        
        with st.expander("👁️ Vista previa del archivo original (primeras 10 filas)"):
            st.dataframe(df_original.head(10), use_container_width=True)
            st.caption(f"Total filas: {len(df_original)} | Columnas: {len(df_original.columns)}")
        
        if procesar:
            if not conceptos.strip():
                st.error("❌ Escribe al menos un concepto para buscar")
            else:
                # Limpiar lista de conceptos
                lista_conceptos = [c.strip().lower() for c in conceptos.split(',') if c.strip()]
                
                st.write("**Buscando:**")
                st.write(f"📝 Conceptos: {', '.join(lista_conceptos)}")
                st.info(f"🔍 Se buscarán filas que contengan: {', '.join(lista_conceptos)}")
                
                with st.spinner('Procesando...'):
                    df_resultado = procesar_archivo(df_original, lista_conceptos)
                
                if len(df_resultado) > 0:
                    st.success(f"✅ **Éxito!** Se encontraron {len(df_resultado)} filas que contienen los conceptos buscados (de {len(df_original)} totales)")
                    
                    # Resumen por concepto
                    st.subheader("📊 Resumen de coincidencias")
                    cols = st.columns(len(lista_conceptos))
                    for i, concepto in enumerate(lista_conceptos):
                        cantidad = (df_resultado[concepto.upper()] != 'No detectado').sum()
                        cols[i].metric(f"{concepto.upper()}", f"{cantidad} / {len(df_resultado)}")
                    
                    # Mostrar resultados
                    st.subheader("📋 Resultados filtrados (solo filas con coincidencias)")
                    st.dataframe(df_resultado, use_container_width=True, height=400)
                    
                    # Botones descarga
                    col1, col2 = st.columns(2)
                    csv = df_resultado.to_csv(index=False).encode('utf-8')
                    col1.download_button(
                        label="📥 Descargar CSV",
                        data=csv,
                        file_name=f"filtrado_{archivo.name.replace('.xlsx', '').replace('.xls', '')}.csv",
                        mime="text/csv",
                        use_container_width=True
                    )
                    
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_resultado.to_excel(writer, sheet_name='Filtrado', index=False)
                    output.seek(0)
                    col2.download_button(
                        label="📊 Descargar Excel",
                        data=output,
                        file_name=f"filtrado_{archivo.name}",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                else:
                    st.warning(f"⚠️ No se encontraron filas que contengan: {', '.join(lista_conceptos)}")
                    st.info("💡 Sugerencia: Revisa que los conceptos escritos coincidan exactamente con la ortografía de tu archivo")
    
    except Exception as e:
        st.error(f"❌ Error: {str(e)}")

else:
    st.markdown("""
    ### 👋 ¡Bienvenido a Grupo Bodeguita Oriente!
    
    **¿Cómo funciona?**
    
    1. 📂 Carga un archivo Excel
    2. ✏️ Escribe los conceptos que quieres buscar (separados por coma)
    3. 🚀 Presiona "Procesar y Clasificar"
    
    ### 🎯 ¿Qué hace?
    - Busca SOLO las filas que contengan los conceptos que escribiste
    - Crea columnas independientes para cada concepto
    - Filtra el resultado mostrando ÚNICAMENTE las filas con coincidencias
    
    ### Ejemplo:
    Si buscas `pago movil, comision`:
    - Solo verás las filas que tengan "pago movil" o "comision"
    - Las filas con "transferencia" o "crédito" NO aparecerán
    
    ---
    **💡 Tip:** El programa filtrará tu archivo y solo mostrará lo que te interesa
    """)

# Footer
st.markdown("---")
st.markdown(
    """
    <div class="footer">
        <strong>Grupo Bodeguita Oriente</strong> - Clasificador de Excel v6.0<br>
        Filtrado inteligente por conceptos personalizados
    </div>
    """,
    unsafe_allow_html=True
)
