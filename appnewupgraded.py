import io
import pandas as pd
import plotly.express as px
import streamlit as st

# 1. Configuración de pantalla orientada a dispositivos móviles
st.set_page_config(
    page_title="Padrón EMR 2026",
    page_icon="📲",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# 2. Estilos CSS personalizados para interfaz móvil elegante y amigable
st.markdown(
    """
    <style>
    /* Estilo del fondo principal */
    .stApp {
        background-color: #F4F6F9;
        font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, sans-serif;
    }
    
    /* Header principal tipo App Móvil */
    .app-header {
        background: linear-gradient(135deg, #1B3B6F 0%, #21295C 100%);
        color: white;
        padding: 24px 18px;
        border-radius: 18px;
        text-align: center;
        margin-bottom: 20px;
        box-shadow: 0 8px 16px rgba(27, 59, 111, 0.15);
    }
    .app-header h1 {
        font-size: 22px;
        font-weight: 700;
        margin: 0;
        letter-spacing: 0.5px;
    }
    .app-header p {
        font-size: 13px;
        margin-top: 6px;
        opacity: 0.88;
        font-weight: 300;
    }

    /* Tarjetas de usuarios (Cards) */
    .citizen-card {
        background-color: #FFFFFF;
        border-radius: 16px;
        padding: 18px 20px;
        margin-bottom: 16px;
        border: 1px solid #E2E8F0;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.03);
        transition: transform 0.2s ease, box-shadow 0.2s ease;
    }
    
    /* Badges de DNI y Distrito */
    .dni-badge {
        background-color: #EBF8FF;
        color: #2B6CB0;
        font-weight: 700;
        font-size: 14px;
        padding: 4px 10px;
        border-radius: 20px;
        display: inline-block;
        border: 1px solid #BEE3F8;
    }
    .location-badge {
        background-color: #F0FFF4;
        color: #276749;
        font-size: 12px;
        font-weight: 600;
        padding: 3px 8px;
        border-radius: 12px;
        display: inline-block;
        border: 1px solid #C6F6D5;
        margin-left: 6px;
    }
    
    /* Nombre del ciudadano */
    .citizen-name {
        font-size: 17px;
        font-weight: 700;
        color: #1A202C;
        margin-top: 10px;
        margin-bottom: 8px;
        text-transform: uppercase;
    }
    
    /* Filas de detalles */
    .detail-row {
        font-size: 13px;
        color: #4A5568;
        margin-bottom: 4px;
    }
    .detail-label {
        font-weight: 600;
        color: #718096;
    }

    /* Ocultar elementos secundarios */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    </style>
""",
    unsafe_allow_html=True,
)


# 3. Función para convertir el DataFrame a Excel en memoria
def exportar_a_excel(df_data):
  output = io.BytesIO()
  df_export = df_data[[
      'DNI_FMT',
      'APELLIDO PATERNO',
      'APELLIDO MATERNO',
      'NOMBRES',
      'DISTRITO',
      'PROVINCIA',
      'REGION',
  ]].copy()
  df_export.columns = [
      'DNI',
      'APELLIDO PATERNO',
      'APELLIDO MATERNO',
      'NOMBRES',
      'DISTRITO',
      'PROVINCIA',
      'REGIÓN',
  ]

  with pd.ExcelWriter(output, engine='openpyxl') as writer:
    df_export.to_excel(writer, index=False, sheet_name='Padron_Filtrado')
  return output.getvalue()


# 4. Encabezado de la App
st.markdown(
    """
    <div class="app-header">
        <h1>📋 PADRÓN EMR 2026</h1>
        <p>Consulta de Ciudadanos y Análisis de Padrón</p>
    </div>
""",
    unsafe_allow_html=True,
)


# 5. Carga de datos con caché
@st.cache_data
def cargar_padron():
  archivo = 'emr2026.xlsx.xlsx'
  try:
    df_data = pd.read_excel(archivo, sheet_name='padron2026')
  except Exception:
    df_data = pd.read_excel(archivo, sheet_name=0)

  df_data.columns = df_data.columns.astype(str).str.strip()
  df_data['DNI_FMT'] = (
      df_data['DNI']
      .astype(str)
      .str.replace(r'\.0$', '', regex=True)
      .str.zfill(8)
  )

  df_data['APELLIDO MATERNO'] = df_data['APELLIDO MATERNO'].fillna('')
  df_data['FULL_NAME'] = (
      df_data['APELLIDO PATERNO']
      + ' '
      + df_data['APELLIDO MATERNO']
      + ' '
      + df_data['NOMBRES']
  ).str.strip()

  return df_data


try:
  df = cargar_padron()

  # 6. Buscador interactivo
  c1, c2 = st.columns([3, 1])
  with c1:
    busqueda = st.text_input(
        'Buscar por DNI o Apellidos',
        placeholder='Ej. 42422 o MAYTAHUARI...',
        label_visibility='collapsed',
    )
  with c2:
    distritos = ['Todos'] + sorted(list(df['DISTRITO'].dropna().unique()))
    filtro_distrito = st.selectbox(
        'Distrito', distritos, label_visibility='collapsed'
    )

  # 7. Lógica de Filtrado
  df_filtrado = df.copy()

  if filtro_distrito != 'Todos':
    df_filtrado = df_filtrado[df_filtrado['DISTRITO'] == filtro_distrito]

  if busqueda:
    query = busqueda.strip().lower()
    mask_dni = df_filtrado['DNI_FMT'].str.lower().str.contains(query)
    mask_nombre = df_filtrado['FULL_NAME'].str.lower().str.contains(query)
    df_filtrado = df_filtrado[mask_dni | mask_nombre]

  total_res = len(df_filtrado)

  # 8. PANEL DE MÉTRICAS Y ESTADÍSTICAS EN TIEMPO REAL
  with st.expander('📊 Panel de Métricas y Estadísticas', expanded=True):
    # Métricas principales en tarjetas
    m1, m2, m3 = st.columns(3)
    m1.metric('Total Registros', f'{total_res:,}')
    m2.metric('Distritos', f"{df_filtrado['DISTRITO'].nunique()}")
    m3.metric('Provincias', f"{df_filtrado['PROVINCIA'].nunique()}")

    if not df_filtrado.empty:
      # Pestañas con gráficos interactivos
      tab1, tab2 = st.tabs(['📍 Por Distrito', '👤 Top Apellidos'])

      with tab1:
        # Gráfico de barras por Distrito
        dist_counts = (
            df_filtrado['DISTRITO']
            .value_counts()
            .reset_index(name='Ciudadanos')
        )
        dist_counts.columns = ['Distrito', 'Ciudadanos']

        fig_dist = px.bar(
            dist_counts,
            x='Distrito',
            y='Ciudadanos',
            text='Ciudadanos',
            color='Distrito',
            color_discrete_sequence=px.colors.qualitative.Prism,
        )
        fig_dist.update_traces(
            textposition='outside',
            hovertemplate='<b>%{x}</b><br>Registros: %{y}',
        )
        fig_dist.update_layout(
            showlegend=False,
            height=300,
            margin=dict(l=10, r=10, t=20, b=10),
            xaxis_title=None,
            yaxis_title=None,
        )
        st.plotly_chart(fig_dist, use_container_width=True)

      with tab2:
        # Gráfico del Top 10 Apellidos Paternos más comunes
        top_ape = (
            df_filtrado['APELLIDO PATERNO']
            .value_counts()
            .head(10)
            .reset_index(name='Cantidad')
        )
        top_ape.columns = ['Apellido', 'Cantidad']

        fig_ape = px.bar(
            top_ape,
            x='Cantidad',
            y='Apellido',
            orientation='h',
            text='Cantidad',
            color='Cantidad',
            color_continuous_scale='Blues',
        )
        fig_ape.update_layout(
            yaxis=dict(autorange='reversed'),
            height=300,
            margin=dict(l=10, r=10, t=20, b=10),
            xaxis_title=None,
            yaxis_title=None,
            coloraxis_showscale=False,
        )
        st.plotly_chart(fig_ape, use_container_width=True)

  # 9. Contenedor de Botones de Exportación
  col_exp1, col_exp2 = st.columns(2)

  with col_exp1:
    excel_bytes = exportar_a_excel(df_filtrado)
    st.download_button(
        label='📊 Reporte Excel',
        data=excel_bytes,
        file_name='Reporte_Padron_EMR2026.xlsx',
        mime=(
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        ),
        use_container_width=True,
    )

  with col_exp2:
    csv_bytes = (
        df_filtrado[[
            'DNI_FMT',
            'APELLIDO PATERNO',
            'APELLIDO MATERNO',
            'NOMBRES',
            'DISTRITO',
        ]]
        .to_csv(index=False)
        .encode('utf-8')
    )
    st.download_button(
        label='📄 Listado CSV',
        data=csv_bytes,
        file_name='Reporte_Padron_EMR2026.csv',
        mime='text/csv',
        use_container_width=True,
    )

  st.divider()

  # 10. Mostrar Tarjetas de Resultados (Vista Móvil)
  if not df_filtrado.empty:
    for idx, row in df_filtrado.head(50).iterrows():
      st.markdown(
          f"""
                <div class="citizen-card">
                    <div>
                        <span class="dni-badge">🪪 DNI: {row['DNI_FMT']}</span>
                        <span class="location-badge">📍 {row['DISTRITO']}</span>
                    </div>
                    <div class="citizen-name">
                        {row['APELLIDO PATERNO']} {row['APELLIDO MATERNO']} {row['NOMBRES']}
                    </div>
                    <div class="detail-row">
                        <span class="detail-label">Provincia / Región:</span> {row['PROVINCIA']} - {row['REGION']}
                    </div>
                </div>
            """,
          unsafe_allow_html=True,
      )

    if total_res > 50:
      st.info(
          '💡 Mostrando los primeros 50 resultados. Refina tu búsqueda para'
          ' acotar los datos.'
      )
  else:
    st.warning('🔍 No se encontraron registros que coincidan con la búsqueda.')

except Exception as e:
  st.error(
      f"Ocurrió un inconveniente al cargar el archivo 'emr2026.xlsx.xlsx': {e}"
  )