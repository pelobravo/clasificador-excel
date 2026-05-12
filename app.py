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
    st.markdown("### Extrae la descripción completa donde aparece cada concepto")
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
        help="Ejemplo: pago movil, comision"
    )
    st.markdown("---")
    
    procesar = st.button("🚀 Procesar y Clasificar", type="primary", use_container_width=True)

def obtener_valor_concepto(row, concepto_buscar):
    """
    Busca el concepto en cada celda de la fila.
    Si lo encuentra, devuelve el valor COMPLETO de esa celda (ej: "PAGO MOVIL COMERCIAL INTERBANCARIO")
    """
    for columna in row.index:
        valor = row[columna]
        if pd.notna(valor):
            try:
                texto_valor = str(valor).lower()
                if concepto_buscar in texto_valor:
                    # Devolvemos el valor ORIGINAL, no el minúsculas
                    return str(valor)
            except:
                continue
    return None

def procesar_archivo(df, lista_conceptos):
    """
    Procesa el archivo y SOLO guarda las filas que tienen al menos un concepto.
    Para cada concepto, guarda el texto COMPLETO de la celda donde apareció.
    """
    resultados = []
    
    for _, fila in df.iterrows():
        # Verificar qué conceptos aparecen en esta fila
        fila_conceptos = {}
        tiene_concepto = False
        
        for concepto in lista_conceptos:
            valor_encontrado = obtener_valor_concepto(fila, concepto)
            if valor_encontrado:
                fila_conceptos[concepto.upper()] = valor_encontrado
                tiene_concepto = True
            else:
                fila_conceptos[concepto.upper()] = 'No detectado'
        
        # Si no tiene ningún concepto, saltamos esta fila
        if not tiene_concepto:
            continue
        
        # Detectar monto
        monto = None
        for col in df.columns:
            valor = fila[col]
            if pd.notna(valor) and isinstance(valor, (int, float)) and valor > 0:
                if monto is None:
                    monto = valor
                if 'monto' in str(col).lower():
                    monto = valor
                    break
        
        # Construir fila resultado
        nueva_fila = {**fila.to_dict()}
        nueva_fila.update(fila_conceptos)
        nueva_fila['💰 Monto'] = monto if monto else 'No aplica'
        resultados.append(nueva_fila)
    
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
                st.error("❌ Escribe al menos un concepto")
            else:
                lista_conceptos = [c.strip().lower() for c in conceptos.split(',') if c.strip()]
                
                st.write("**Conceptos a buscar:**")
                st.write(f"📝 {', '.join([c.upper() for c in lista_conceptos])}")
                
                with st.spinner('Procesando...'):
                    df_resultado = procesar_archivo(df_original, lista_conceptos)
                
                if len(df_resultado) > 0:
                    st.success(f"✅ Encontradas {len(df_resultado)} filas con los conceptos solicitados")
                    
                    # Resumen
                    st.subheader("📊 Resumen por concepto")
                    cols = st.columns(len(lista_conceptos))
                    for i, concepto in enumerate(lista_conceptos):
                        cant = (df_resultado[concepto.upper()] != 'No detectado').sum()
                        cols[i].metric(concepto.upper(), f"{cant} / {len(df_resultado)}")
                    
                    # Resultados
                    st.subheader("📋 Filas filtradas")
                    st.dataframe(df_resultado, use_container_width=True, height=400)
                    
                    # Descargas
                    col1, col2 = st.columns(2)
                    csv = df_resultado.to_csv(index=False).encode('utf-8')
                    col1.download_button("📥 CSV", csv, f"resultados_{archivo.name.replace('.xlsx', '.csv')}", "text/csv", use_container_width=True)
                    
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_resultado.to_excel(writer, index=False)
                    col2.download_button("📊 Excel", output.getvalue(), f"resultados_{archivo.name}", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                else:
                    st.warning("No se encontraron filas con esos conceptos")
    
    except Exception as e:
        st.error(f"Error: {str(e)}")

else:
    st.markdown("""
    ### 👋 ¡Bienvenido!
    
    1. Carga un archivo Excel
    2. Escribe los conceptos (ej: pago movil, comision)
    3. Procesar
    
    **El programa extraerá la DESCRIPCIÓN COMPLETA de cada concepto encontrado.**
    """)

st.markdown("---")
st.markdown('<div class="footer"><strong>Grupo Bodeguita Oriente</strong> - Clasificador v6.2</div>', unsafe_allow_html=True)
