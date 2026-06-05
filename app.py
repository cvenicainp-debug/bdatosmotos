import streamlit as st
import pandas as pd
import io

# Configuración de la página
st.set_page_config(page_title="Catálogo de Motos", layout="wide")

# URL de exportación directa a CSV de tu Google Sheets
GSHEETS_URL = "https://docs.google.com/spreadsheets/d/1A7mxMAoqpLn_4AzNx7yB7-aw8r41MdTlCyuvd0IckeE/export?format=csv"

@st.cache_data(ttl=600)
def cargar_datos(url):
    try:
        # 1. Leer el archivo ignorando temporalmente tipos de datos para limpiar cabeceras vacías
        df = pd.read_csv(url, dtype=str)
        
        # Limpieza por si Google Sheets exporta filas vacías arriba (común si hay títulos colgados)
        if 'MARCA' not in df.columns:
            # Buscar en qué fila real están los encabezados correctos
            for i in range(len(df)):
                fila_valores = df.iloc[i].astype(str).str.upper().values
                if 'MARCA' in fila_valores or 'ID MOTO' in fila_valores:
                    df.columns = df.iloc[i]
                    df = df.iloc[i+1:].reset_index(drop=True)
                    break
        
        # Limpiar espacios en blanco en los nombres de las columnas
        df.columns = df.columns.str.strip()
        
        # 2. Formatear ID MOTO y SKU respetando estrictamente los ceros adelante
        if 'ID MOTO' in df.columns:
            # Quitamos decimales si se colaron (.0) y rellenamos con ceros a la izquierda (ej: "00108")
            df['ID MOTO'] = df['ID MOTO'].fillna('').astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
            # Si tus IDs tienen longitudes variables pero quieres asegurar al menos "00" adelante:
            df['ID MOTO'] = df['ID MOTO'].apply(lambda x: f"00{x}" if x and not x.startswith('00') else x)
            
        if 'SKU' in df.columns:
            df['SKU'] = df['SKU'].fillna('').astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
            df['SKU'] = df['SKU'].apply(lambda x: f"00{x}" if x and not x.startswith('00') else x)

        return df
    except Exception as e:
        st.error(f"Error al procesar la planilla de Google Sheets: {e}")
        return None

# Título de la App
st.title("🏍️ Base de Datos - Catálogo de Motos")

# Manejo del estado de los datos
if 'datos_motos' not in st.session_state:
    st.session_state['datos_motos'] = cargar_datos(GSHEETS_URL)

df_motos = st.session_state['datos_motos']

if df_motos is not None and not df_motos.empty:
    
    # Botón superior de actualización fija
    if st.button("🔄 Actualizar Base de Datos desde Google Sheets", use_container_width=True):
        st.cache_data.clear()
        st.session_state['datos_motos'] = cargar_datos(GSHEETS_URL)
        st.rerun()

    # Creación de Solapas
    tab_consulta, tab_descarga = st.tabs(["🔍 Consultar información", "📥 Descargar y Actualizar (.xlsx)"])

    # ----------------------------------------------------
    # SOLAPA 1: CONSULTAS
    # ----------------------------------------------------
    with tab_consulta:
        # Validar columnas requeridas
        columnas_necesarias = ['MARCA', 'MODELO', 'VERSION', 'CIL.', 'ID MOTO', 'DESDE', 'HASTA', 'SKU', 'PRODUCTO']
        columnas_faltantes = [col for col in columnas_necesarias if col not in df_motos.columns]
        
        if columnas_faltantes:
            st.warning(f"Faltan o están mal escritas estas columnas en el origen: {columnas_faltantes}")
            st.write("Columnas detectadas actualmente:", list(df_motos.columns))
        else:
            col_sel1, col_sel2 = st.columns(2)
            
            with col_sel1:
                marcas = sorted(df_motos['MARCA'].dropna().unique())
                marca_sel = st.selectbox("Selecciona la MARCA:", marcas)
                df_filtrado_marca = df_motos[df_motos['MARCA'] == marca_sel]
                
            with col_sel2:
                modelos = sorted(df_filtrado_marca['MODELO'].dropna().unique())
                modelo_sel = st.selectbox("Selecciona el MODELO:", modelos)
                df_final = df_filtrado_marca[df_filtrado_marca['MODELO'] == modelo_sel]

            if not df_final.empty:
                reg = df_final.iloc[0]
                
                st.markdown("---")
                c1, c2 = st.columns(2)
                
                with c1:
                    st.markdown("### 📋 Datos Técnicos")
                    st.info(f"""
                    **MARCA:** {reg.get('MARCA', '')}  
                    **MODELO:** {reg.get('MODELO', '')}  
                    **VERSIÓN:** {reg.get('VERSION', '')}  
                    **CILINDRADA (CIL):** {reg.get('CIL.', '')}
                    """)
                    
                with c2:
                    st.markdown("### 🆔 Identificación y Período")
                    st.success(f"""
                    **ID MOTO:** {reg.get('ID MOTO', '')}  
                    **AÑO DESDE:** {reg.get('DESDE', '')}  
                    **AÑO HASTA:** {reg.get('HASTA', '')}
                    """)
                    
                st.markdown("---")
                st.markdown("### 📦 Productos Relacionados")
                
                # Mostrar listado limpio de SKUs y Productos
                df_prod_mostrar = df_final[['SKU', 'PRODUCTO']].drop_duplicates().reset_index(drop=True)
                st.dataframe(df_prod_mostrar, use_container_width=True)

    # ----------------------------------------------------
    # SOLAPA 2: DESCARGAS
    # ----------------------------------------------------
    with tab_descarga:
        st.subheader("Descargar Base de Datos Completa")
        
        # Buffer en memoria para compilar el Excel sin guardarlo en disco de forma local rígida
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_motos.to_excel(writer, index=False, sheet_name='Base_Motos')
            workbook  = writer.book
            worksheet = writer.sheets['Base_Motos']
            # Forzar formato texto en las columnas clave del Excel descargado
            format_texto = workbook.add_format({'num_format': '@'})
            worksheet.set_column('A:A', None, format_texto)
            worksheet.set_column('I:I', None, format_texto)
            
        data_excel = output.getvalue()
        
        st.download_button(
            label="📥 Descargar planilla completa en formato .xlsx",
            data=data_excel,
            file_name="Catalogo_Motos.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        
        st.markdown("---")
        st.subheader("Cargar Base Externa Temporal")
        archivo_manual = st.file_uploader("Sube un archivo .xlsx nuevo si quieres anular temporalmente el link fijo:", type=["xlsx"])
        if archivo_manual is not None:
            df_motos = pd.read_excel(archivo_manual, dtype=str)
            st.session_state['datos_motos'] = df_motos
            st.success("¡Datos reemplazados en pantalla! Ya puedes usar el buscador con tu archivo local.")
else:
    st.info("Esperando conexión o datos válidos desde la planilla de Google Sheets...")