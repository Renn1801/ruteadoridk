import streamlit as st
import pandas as pd
import io
import os
import unicodedata
from geopy.geocoders import Nominatim
import folium
from streamlit_folium import st_folium

# Función para normalizar: quita acentos y pasa a mayúsculas
def normalizar_texto(texto):
    if pd.isna(texto): return ""
    # Normaliza a forma NFD (separa el acento de la letra)
    nfkd_form = unicodedata.normalize('NFKD', str(texto).upper())
    # Filtra y devuelve solo los caracteres que no son combinables (sin acento)
    return "".join([c for c in nfkd_form if not unicodedata.combining(c)]).strip()

st.set_page_config(layout="wide", page_title="Ruteador BRAT", page_icon="logo.png")
st.markdown("""
    <style>
    .stApp {
        background-color: #000000;
    }
    
    /* Opcional: Cambiar el color del texto a blanco para que se lea bien sobre el fondo negro */
    h1, h2, h3, p, div {
        color: white;
    }
    </style>
    """, unsafe_allow_html=True)
st.logo("logo.png")
st.markdown("<h1 style='text-align: left; color: white; font-size: 70px;'>BRAT LOGISTICA</h1>", unsafe_allow_html=True)
col1, col2 = st.columns(2)
# Descargar Plantilla de Ejemplo
archivo_plantilla = "plantilla.xlsx"

with open(archivo_plantilla, "rb") as file:
    btn = st.download_button(
        label="📥 Descargar plantilla de ejemplo",
        data=file,
        file_name="plantilla.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
# Subida de archivos
with col1:
    up_ped = st.file_uploader("Subir plantilla_pedidos.xlsx", type=["xlsx"])
with col2:
    up_asig = st.file_uploader("Subir ASIGNACIONES.xlsx", type=["xlsx"])

if up_ped and up_asig and os.path.exists('maestro_zonas.xlsx'):
    pedidos = pd.read_excel(up_ped)
    asig = pd.read_excel(up_asig)
    reglas = pd.read_excel('maestro_zonas.xlsx')

    pedidos['_orig_index'] = pedidos.index

    if 'Codigo' in asig.columns:
        duplicados_asig = asig[asig.duplicated(subset=['Codigo'], keep=False)]
        if not duplicados_asig.empty:
            st.warning(f"Se eliminaron {duplicados_asig['Codigo'].nunique()} códigos duplicados en ASIGNACIONES; se usará el primer registro por código.")
            asig = asig.drop_duplicates(subset=['Codigo'], keep='first').copy()
    else:
        st.error("El archivo ASIGNACIONES debe tener la columna 'Codigo'.")
    
    # Preparamos las llaves normalizadas
    pedidos['key'] = pedidos['Localidad'].apply(normalizar_texto)
    
    # Asumimos que la columna 0 de reglas es la Localidad
    reglas_tmp = reglas.copy()
    reglas_tmp.columns = ['key', 'Codigo'] # Aseguramos nombres
    reglas_tmp['key'] = reglas_tmp['key'].apply(normalizar_texto)
    
    # Merge usando las llaves normalizadas
    df = pedidos.merge(reglas_tmp[['key', 'Codigo']], on='key', how='left')
    df = df.merge(asig, on='Codigo', how='left')
    
    # Evitar duplicar filas del pedido cuando ASIGNACIONES tiene el mismo Codigo repetido
    if '_orig_index' in df.columns:
        df = df.drop_duplicates(subset=['_orig_index', 'Codigo', 'Chofer'], keep='first').copy()

    if 'key' in df.columns:
        df = df.drop(columns=['key'])
    
    # Normalizar la columna 'Chofer' para evitar tipos mixtos (NaN -> cadena)
    if 'Chofer' in df.columns:
        df['Chofer'] = df['Chofer'].fillna('Sin Chofer').astype(str).str.strip()
    else:
        df['Chofer'] = 'Sin Chofer'

    # Separar filas sin chofer para revisarlas por separado (no las procesamos en el mapa)
    mask_sin_chofer = df['Chofer'].str.strip().str.lower().isin(['', 'sin chofer', 'none'])
    df_sin_chofer = df[mask_sin_chofer].copy()
    # DataFrame que sí procesaremos (solo filas con chofer asignado)
    df = df[~mask_sin_chofer].copy()

    if '_orig_index' in df.columns:
        df = df.drop(columns=['_orig_index'])
    if '_orig_index' in df_sin_chofer.columns:
        df_sin_chofer = df_sin_chofer.drop(columns=['_orig_index'])
        
    st.write("### Resultados en pantalla")
    st.dataframe(df)
    
    # Preparar archivo de exportación incluyendo filas sin chofer (para revisión)
    try:
        df_export = pd.concat([df, df_sin_chofer], ignore_index=True)
    except NameError:
        df_export = df.copy()

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        df_export.to_excel(writer, index=False)
    
    st.download_button(
        label="📥 Descargar Excel con Asignaciones",
        data=buffer.getvalue(),
        file_name="pedidos_asignados.xlsx",
        mime="application/vnd.ms-excel"
    )

elif not os.path.exists('maestro_zonas.xlsx'):
    st.error("Error: Falta el archivo 'maestro_zonas.xlsx' en la carpeta.")

if up_ped and up_asig and os.path.exists('maestro_zonas.xlsx'):
    # --- BLOQUE MAPA CON AVISO DE ERRORES ---
    st.write("### Mapa de Distribución")

    from geopy.geocoders import Nominatim, ArcGIS
    from concurrent.futures import ThreadPoolExecutor, as_completed
    
    geolocator_arcgis = ArcGIS()
    geolocator_nominatim = Nominatim(user_agent="logistica_app_v2")

    @st.cache_data
    def buscar_coordenadas(direccion, localidad):
        import re
        
        if pd.isna(direccion) or pd.isna(localidad):
            return (None, None)
        
        direccion = str(direccion).strip()
        localidad = str(localidad).strip()
        
        # Limpieza y normalización de la dirección
        direccion = re.sub(r'\bDr\b\.?', 'Doctor', direccion, flags=re.IGNORECASE)
        direccion = re.sub(r'\bAv\b\.?', 'Avenida', direccion, flags=re.IGNORECASE)
        direccion = re.sub(r'\bex\b', 'y', direccion, flags=re.IGNORECASE)
        direccion = re.sub(r'\bSN\b\.?', '', direccion, flags=re.IGNORECASE).strip()
        localidad = localidad.replace(' ex ', ' y ')
        
        # Solo 3 búsquedas (optimizado)
        busquedas = [
            f"{direccion}, {localidad}, Buenos Aires, Argentina",
            f"{localidad}, Buenos Aires, Argentina",
            f"{localidad}",
        ]
        
        # Intenta con ArcGIS primero (más rápido para Argentina)
        for busqueda in busquedas:
            try:
                location = geolocator_arcgis.geocode(busqueda, timeout=5)
                if location:
                    return (location.latitude, location.longitude)
            except:
                pass
        
        # Fallback: Nominatim
        for busqueda in busquedas:
            try:
                location = geolocator_nominatim.geocode(busqueda, timeout=5)
                if location:
                    return (location.latitude, location.longitude)
            except:
                pass
        
        return (None, None)

    with st.spinner('Procesando ubicaciones...'):
        # Procesamiento paralelo para acelerar geocodificación
        def procesar_fila(idx, row):
            coords = buscar_coordenadas(row['Direccion'], row['Localidad'])
            return idx, coords[0], coords[1]
        
        # Usar hasta 4 threads para no sobrecargar los servidores de geocodificación
        with ThreadPoolExecutor(max_workers=4) as executor:
            futures = {executor.submit(procesar_fila, idx, row): idx for idx, row in df.iterrows()}
            
            resultados_lat = {}
            resultados_lon = {}
            
            for future in as_completed(futures):
                try:
                    idx, lat, lon = future.result()
                    resultados_lat[idx] = lat
                    resultados_lon[idx] = lon
                except:
                    pass
        
        df['Lat'] = df.index.map(resultados_lat)
        df['Lon'] = df.index.map(resultados_lon)

    # 1. Mapa con colores por chofer
    df_mapa = df.dropna(subset=['Lat', 'Lon'])
    m = folium.Map(location=[-34.6, -58.4], zoom_start=10, tiles='OpenStreetMap')
    
    # Obtener todos los choferes únicos (incluso los sin coordenadas)
    todos_choferes = sorted(df['Chofer'].unique())
    num_choferes = len(todos_choferes)
    
    # Paleta de colores extendida soportada por folium
    colores_folium = ['red', 'blue', 'green', 'purple', 'orange', 'darkred', 'darkblue', 
                      'darkgreen', 'cadetblue', 'darkpurple', 'pink', 'gray', 'lightred', 
                      'lightblue', 'lightgreen', 'lightgray', 'beige', 'lightcyan', 'black']
    
    # Si hay más choferes que colores base, extender la paleta
    if num_choferes > len(colores_folium):
        colores_disponibles = (colores_folium * ((num_choferes // len(colores_folium)) + 1))[:num_choferes]
    else:
        colores_disponibles = colores_folium[:num_choferes]
    
    # Generar colores hexadecimales para el dashboard (mismo orden)
    import colorsys
    
    def generar_colores_hex(n):
        colores_hex = []
        for i in range(n):
            hue = i / max(n, 1)
            saturation = 0.8
            lightness = 0.5
            r, g, b = colorsys.hls_to_rgb(hue, lightness, saturation)
            hex_color = f"#{int(r*255):02x}{int(g*255):02x}{int(b*255):02x}"
            colores_hex.append(hex_color)
        return colores_hex
    
    colores_hex_generados = generar_colores_hex(num_choferes)
    
    # Mapeo de colores para cada chofer
    mapa_colores_chofer_folium = {chofer: colores_disponibles[i] 
                                   for i, chofer in enumerate(todos_choferes)}
    mapa_colores_chofer_hex = {chofer: colores_hex_generados[i] 
                               for i, chofer in enumerate(todos_choferes)}
    
    from folium.features import DivIcon

    # Agregar marcadores precisos con iconos SVG y hex colors
    coord_counts = {}
    for _, row in df_mapa.iterrows():
        lat = row['Lat']
        lon = row['Lon']
        if pd.isna(lat) or pd.isna(lon):
            continue

        key = (round(float(lat), 6), round(float(lon), 6))
        coord_counts[key] = coord_counts.get(key, 0) + 1

    coord_seen = {}
    for _, row in df_mapa.iterrows():
        lat = row['Lat']
        lon = row['Lon']
        if pd.isna(lat) or pd.isna(lon):
            continue

        key = (round(float(lat), 6), round(float(lon), 6))
        count = coord_counts.get(key, 1)
        seen = coord_seen.get(key, 0)
        coord_seen[key] = seen + 1
        index = seen + 1

        display_lat = float(lat)
        display_lon = float(lon)

        if count > 1:
            import math
            radius = 0.000025
            angle = (index - 1) * (360 / count)
            rad = math.radians(angle)
            display_lat = display_lat + radius * math.cos(rad)
            display_lon = display_lon + radius * math.sin(rad)

        color_hex = mapa_colores_chofer_hex.get(row['Chofer'], '#3388ff')

        svg_icon = f"""
        <svg xmlns='http://www.w3.org/2000/svg' width='26' height='34' viewBox='0 0 24 34'>
          <path d='M12 0C7 0 3 4 3 9c0 7.5 9 18.8 9 18.8S21 16.5 21 9c0-5-4-9-9-9z' fill='{color_hex}' stroke='#222' stroke-width='0.5'/>
          <circle cx='12' cy='9' r='3.5' fill='white' />
        </svg>
        """

        icon = DivIcon(
            icon_size=(26, 34),
            icon_anchor=(13, 34),
            html=svg_icon
        )

        folium.Marker(
            [display_lat, display_lon],
            popup=folium.Popup(f"<b>Chofer:</b> {row['Chofer']}<br><b>Dirección:</b> {row.get('Direccion','')}<br><b>Localidad:</b> {row.get('Localidad','')}", max_width=300),
            tooltip=row.get('Direccion', ''),
            icon=icon
        ).add_to(m)

    # Mostrar mapa a ancho completo
    st_folium(m, width=1400, height=700)

    # Dashboard con gráfico de barras
    st.write("---")
    st.write("### Distribución de Paquetes por Chofer")
    
    # Contar paquetes por chofer
    paquetes_por_chofer = df_mapa['Chofer'].value_counts().sort_values(ascending=True)
    total_paquetes = paquetes_por_chofer.sum()
    
    # Crear gráfico de barras con plotly
    import plotly.graph_objects as go
    
    # Preparar datos para el gráfico
    choferes = paquetes_por_chofer.index.tolist()
    cantidades = paquetes_por_chofer.values.tolist()
    colores = [mapa_colores_chofer_hex[chofer] for chofer in choferes]
    
    # Crear figura
    fig = go.Figure(
        data=[go.Bar(
            x=cantidades,
            y=choferes,
            orientation='h',
            marker=dict(color=colores),
            text=[f"{int(c)} paquetes" for c in cantidades],
            textposition='auto',
            hovertemplate='<b>%{y}</b><br>Paquetes: %{x}<extra></extra>'
        )]
    )
    
    # Configurar layout
    fig.update_layout(
        title=dict(text=f"Total: {total_paquetes} paquetes", font=dict(color='#ffffff', size=16)),
        xaxis_title="Cantidad de Paquetes",
        yaxis_title="Chofer",
        height=max(300, len(choferes) * 30),
        showlegend=False,
        margin=dict(l=150, r=50, t=60, b=50),
        plot_bgcolor='#1a1a1a',
        paper_bgcolor='#0e1117',
        font=dict(size=13, color='#ffffff'),
        xaxis=dict(
            showgrid=True,
            gridwidth=1,
            gridcolor='#31333d',
            zeroline=False,
            tickfont=dict(color='#ffffff')
        ),
        yaxis=dict(
            tickfont=dict(color='#ffffff', size=12),
            zeroline=False
        )
    )
    
    fig.update_traces(textfont=dict(color='#ffffff', size=11))
    
    st.plotly_chart(fig, use_container_width=True)

    # 2. LISTADO DE ERRORES (Aquí ves las que fallaron)
    df_errores = df[df['Lat'].isna()]

    if not df_errores.empty:
        st.error(f"⚠️ Atención: No se pudieron geolocalizar {len(df_errores)} direcciones:")
        # Mostramos solo la dirección y localidad para identificar rápido el error
        st.table(df_errores[['Direccion', 'Localidad']])
    else:
        st.success("✅ ¡Todas las direcciones fueron ubicadas correctamente!")
    
    # Mostrar filas que no tenían chofer asignado (no fueron procesadas en el mapa)
    if 'df_sin_chofer' in locals() and not df_sin_chofer.empty:
        st.warning(f"⚠️ Hay {len(df_sin_chofer)} filas sin chofer asignado (no procesadas en el mapa):")
        cols_show = [c for c in ['Direccion', 'Localidad', 'Codigo', 'Chofer'] if c in df_sin_chofer.columns]
        st.table(df_sin_chofer[cols_show])
    # --- FIN BLOQUE MAPA ---