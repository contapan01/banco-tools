import streamlit as st
import pandas as pd
from io import BytesIO
import datetime

# Configuración de página
st.set_page_config(
    page_title="Portal de Herramientas Bancarias",
    page_icon="🏦",
    layout="wide"
)

# Estilo personalizado para un look "premium"
st.markdown("""
<style>
    .main {
        background-color: #f8f9fa;
    }
    .stMetric {
        background-color: #ffffff;
        padding: 15px;
        border-radius: 10px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }
    .stButton>button {
        width: 100%;
        border-radius: 5px;
        height: 3em;
        background-color: #007bff;
        color: white;
        font-weight: bold;
    }
    .stDownloadButton>button {
        width: 100%;
        border-radius: 5px;
        height: 3em;
        background-color: #28a745;
        color: white;
        font-weight: bold;
    }
</style>
""", unsafe_allow_html=True)

# --- Diccionarios de Soporte ---
meses_es = {
    'ENE': '01', 'FEB': '02', 'MAR': '03', 'ABR': '04', 'MAY': '05', 'JUN': '06',
    'JUL': '07', 'AGO': '08', 'SEP': '09', 'OCT': '10', 'NOV': '11', 'DIC': '12'
}

# --- Lógicas de Procesamiento ---

def process_bank_bgp(file, input_period):
    """Lógica adaptativa y refinada para BGP (Excel o TXT)"""
    if file.name.lower().endswith('.txt'):
        # Leer el TXT usando punto y coma como separador (formato exacto del banco)
        df = pd.read_csv(file, sep=";", engine='python', dtype=str).dropna(how='all').reset_index(drop=True)
        
        # Mapeo de columnas flexible para manejar acentos
        col_map = {
            'Debito': 'Whitdrawals', 'Débito': 'Whitdrawals',
            'Credito': 'Deposits', 'Crédito': 'Deposits',
            'Fecha': 'Date',
            'Descripción': 'Description'
        }
        
        # Normalizar nombres de columnas existentes
        df.columns = [col_map.get(c, c) for c in df.columns]
    else:
        df = pd.read_excel(file, sheet_name="BGPCheckingMovementsExcel", skiprows=6)
        nuevos_nombres = {'Fecha': 'Date', 'Descripción': 'Description', 'Débito': 'Whitdrawals', 'Crédito': 'Deposits'}
        df = df.rename(columns=nuevos_nombres)

    # --- Limpieza y Transformación ---
    
    # 1. Limpieza de Descripciones (reemplazar ? por espacio)
    if 'Description' in df.columns:
        df['Description'] = df['Description'].astype(str).str.replace('?', ' ', regex=False).str.strip()

    # 2. Limpieza de Montos (quitar $ y ,)
    for col in ['Whitdrawals', 'Deposits']:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace(',', '', regex=False).str.replace('$', '', regex=False).str.strip()
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
    
    # 3. Lógica de Referencias Inteligentes
    # Prioridad: Referencia 1 -> Generación Secuencial
    period_str = str(input_period)
    sequence = (df.index + 1).astype(str).str.zfill(3)
    df['BG_Fallback'] = 'BG' + period_str + '-' + sequence
    
    if 'Referencia 1' in df.columns:
        # Si Referencia 1 tiene datos reales (más de 2 caracteres), los usamos. Si no, usamos el fallback.
        df['Reference Number'] = df.apply(
            lambda row: str(row['Referencia 1']) if pd.notna(row['Referencia 1']) and len(str(row['Referencia 1']).strip()) > 2 
            else row['BG_Fallback'], axis=1
        )
    else:
        df['Reference Number'] = df['BG_Fallback']

    # 4. Formateo de Fecha
    if 'Date' in df.columns:
        df['Date'] = pd.to_datetime(df['Date'], dayfirst=True, errors='coerce')
        df['Date'] = df['Date'].dt.strftime('%Y-%m-%d')

    # Columnas finales para Zoho
    df_final = df[['Date', 'Description', 'Whitdrawals', 'Deposits', 'Reference Number']].copy()
    
    return df_final.dropna(subset=['Date'])

def process_bank_mb(file, input_period):
    """Lógica para el segundo banco (MB/Motor Bank)"""
    df = pd.read_excel(file, sheet_name="Sheet1", skiprows=1)
    # Seleccionamos las columnas relevantes (Fecha, Descripción, Débito, Crédito, Referencia)
    df = df[['Fecha', df.columns[1], 'Débito', 'Crédito', 'No. de Referencia']]
    df.columns = ['Fecha', 'Descripción', 'Débito', 'Crédito', 'Referencia']

    # Procesar Fechas con meses en español
    df['Fecha'] = df['Fecha'].astype(str).str.upper()
    for mes, num in meses_es.items():
        df['Fecha'] = df['Fecha'].str.replace(f'/{mes}/', f'/{num}/', regex=False)

    df['Fecha'] = pd.to_datetime(df['Fecha'], format='%d/%m/%Y', errors='coerce').dt.date
    df['Fecha'] = df['Fecha'].astype(str)

    # Limpiar montos
    df['Débito'] = pd.to_numeric(df['Débito'], errors='coerce').fillna(0)
    df['Crédito'] = pd.to_numeric(df['Crédito'], errors='coerce').fillna(0)

    df['Monto'] = df['Crédito'] - df['Débito']
    df['Tipo'] = df['Monto'].apply(lambda x: 'Crédito' if x > 0 else 'Débito')
    df = df.reset_index(drop=True)

    procesado = []
    skip_indices = set()

    # Lógica de consolidación (6.5% - 7.5%)
    for i in range(len(df)):
        if i in skip_indices: continue
        actual = df.loc[i]
        if i + 1 < len(df):
            siguiente = df.loc[i + 1]
            if actual['Tipo'] == siguiente['Tipo']:
                monto1 = abs(actual['Monto'])
                monto2 = abs(siguiente['Monto'])
                if max(monto1, monto2) > 0:
                    porcentaje = round((min(monto1, monto2) / max(monto1, monto2)) * 100, 2)
                    if 6.5 <= porcentaje <= 7.5:
                        desc = actual['Descripción'] if monto1 >= monto2 else siguiente['Descripción']
                        ref = actual['Referencia'] if pd.notna(actual['Referencia']) and str(actual['Referencia']).strip() else siguiente['Referencia']
                        procesado.append({
                            'Fecha': actual['Fecha'], 'Descripción': desc,
                            'Monto': actual['Monto'] + siguiente['Monto'], 'Referencia': ref
                        })
                        skip_indices.add(i + 1)
                        continue

        procesado.append({
            'Fecha': actual['Fecha'], 'Descripción': actual['Descripción'],
            'Monto': actual['Monto'], 'Referencia': actual['Referencia']
        })

    resultado = pd.DataFrame(procesado)
    # PrefijoRef ahora viene del input_period manual
    resultado['PrefijoRef'] = "MB" + input_period

    # Autogenerar Referencia si falta
    resultado['Referencia'] = resultado.apply(
        lambda row: f"{row['PrefijoRef']}-{str(row.name + 1).zfill(3)}"
        if pd.isna(row['Referencia']) or str(row['Referencia']).strip() in ['0', '', 'nan']
        else row['Referencia'], axis=1
    )

    # Finalizar formato
    resultado['Whitdrawals'] = resultado['Monto'].apply(lambda x: abs(x) if x < 0 else 0)
    resultado['Deposits'] = resultado['Monto'].apply(lambda x: x if x > 0 else 0)
    resultado = resultado[['Fecha', 'Descripción', 'Whitdrawals', 'Deposits', 'Referencia']]
    resultado = resultado.rename(columns={'Fecha': 'Date', 'Referencia': 'Reference Number'})
    return resultado

# --- Interfaz de Usuario ---

st.title("🏦 Portal de Herramientas Bancarias")
st.write("Selecciona el banco y procesa tus archivos para Zoho Books.")

with st.sidebar:
    st.header("Configuración")
    banco_opcion = st.selectbox(
        "Seleccione el Banco",
        ["BGEN", "Motor Bank (MB)"],
        help="Elige el banco correspondiente para aplicar la lógica correcta."
    )
    st.write("---")
    st.info(f"Modo actual: {banco_opcion}")
    st.caption("v1.2.1 | Hetzner Cloud")

# Layout de inputs
col_file, col_pref = st.columns([2, 1])

with col_file:
    # Permitimos .txt también ya que BGP a veces usa esa extensión aunque sea legible por excel
    uploaded_file = st.file_uploader(f"Subir archivo para {banco_opcion}", type=["xlsx", "xls", "txt"])

with col_pref:
    # Ahora el periodo es manual y obligatorio para ambos bancos
    st.write("**Referencia Manual**")
    periodo = st.text_input("Periodo (AAAAMM)", value="", placeholder="Ej: 202602", help="Escribe el periodo que quieres que aparezca en la referencia (6 dígitos).")

if uploaded_file:
    if st.button("🚀 Procesar Movimientos"):
        if not periodo or len(periodo) != 6:
            st.warning("⚠️ Por favor, ingresa un periodo válido de 6 dígitos (AAAAMM) antes de procesar.")
        else:
            with st.spinner("Procesando lógica bancaria..."):
                try:
                    if banco_opcion == "BGEN":
                        df_final = process_bank_bgp(uploaded_file, periodo)
                    else:
                        df_final = process_bank_mb(uploaded_file, periodo)
                    
                    if df_final is not None:
                        st.success(f"¡Procesamiento de {banco_opcion} completado!")
                        
                        # --- Resumen de Totales ---
                        st.write("---")
                        st.subheader("📊 Resumen del Procesamiento")
                        
                        total_debitos = df_final['Whitdrawals'].sum()
                        total_creditos = df_final['Deposits'].sum()
                        
                        m_col1, m_col2, m_col3 = st.columns(3)
                        
                        with m_col1:
                            st.metric("Movimientos", f"{len(df_final):,}")
                        with m_col2:
                            st.metric("Total Débitos (Salidas)", f"${total_debitos:,.2f}", delta_color="inverse")
                        with m_col3:
                            st.metric("Total Créditos (Entradas)", f"${total_creditos:,.2f}")
                        
                        st.write("---")
                        
                        # Vista Previa
                        st.subheader("👀 Vista Previa")
                        st.dataframe(df_final.head(10), use_container_width=True)
                        
                        # Descarga
                        output = BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            df_final.to_excel(writer, index=False)
                        
                        st.download_button(
                            label=f"📥 Descargar Excel para Zoho ({banco_opcion})",
                            data=output.getvalue(),
                            file_name=f"Zoho_{banco_opcion.replace(' ', '_')}_{datetime.datetime.now().strftime('%Y%m%d')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                except Exception as e:
                    st.error(f"Error crítico en el procesamiento: {e}")
                    st.info("Asegúrate de que el formato del archivo subido corresponda al banco seleccionado.")
else:
    st.info("Seleccione un archivo de Excel para comenzar.")
