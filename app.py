from geopy import geocoders
import streamlit as st
import pandas as pd
import io
import os
import unicodedata
from geopy.geocoders import ArcGIS, Nominatim, Photon
import folium
from streamlit_folium import st_folium
import re
import time

# api
geo_arcgis = ArcGIS(user_agent="app_logistica_renzo")
geo_photon = Photon(user_agent="app_logistica_renzo")
geo_osm = Nominatim(user_agent="app_logistica_renzo")

# limpieza
def limpiar_direccion_para_mapa(direccion):
    if pd.isna(direccion): return ""
    d = str(direccion).upper() # Todo a mayúsculas
    
    # 1. Quitar basura de horarios y descripciones extra
    d = re.sub(r'HASTA LAS.*', '', d)
    d = re.sub(r'ESQ\..*', '', d)
    d = re.sub(r'PISO:.*', '', d)
    d = re.sub(r'DEPARTAMENTO:.*', '', d)
    d = re.sub(r'MANZANA.*', '', d)
    d = re.sub(r'ENTRE.*', '', d)
    d = re.sub(r'CASA', '', d)
    d = re.sub(r'SN', '', d)
    
    # 2. Reemplazos críticos
    reemplazos = {
        "AV.": "AVENIDA ",
        "AV ": "AVENIDA ",
        "N°": "",
        "NRO": ""
    }
    for k, v in reemplazos.items():
        d = d.replace(k, v)
        
    # 3. Limpiar espacios extra
    d = re.sub(r'\s+', ' ', d)
    return d.strip()

# mas limpieza
def normalizar_texto(texto):
    if pd.isna(texto): return ""
    nfkd_form = unicodedata.normalize('NFKD', str(texto).upper())
    return "".join([c for c in nfkd_form if not unicodedata.combining(c)]).strip()

def procesar_fila(idx, row):
            coords = buscar_coordenadas(row['Direccion'], row['Localidad'], row.get('cp', ''))
            return idx, coords[0], coords[1]

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

# Descargar de archivos
archivo_plantilla = "plantilla.xlsx"
archivo_asignaciones = "plantilla_asignaciones.xlsx"
with col1:
    with open(archivo_plantilla, "rb") as file:
        btn1 = st.download_button(  # Cambié a btn1 para diferenciarlo
            label="📥 Descargar plantilla de ejemplo",
            data=file,
            file_name="plantilla.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

with col2:
    with open(archivo_asignaciones, "rb") as file:
        btn2 = st.download_button(  # Cambié a btn2 para diferenciarlo
            label="📥 Descargar plantilla de asignaciones",
            data=file,
            file_name="plantilla_asignaciones.xlsx",
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

    # Normalizar CP como texto para que funcione igual si viene como número o como string
    if 'cp' in pedidos.columns:
        pedidos['cp'] = pedidos['cp'].astype(str).str.strip()

    pedidos['_orig_index'] = pedidos.index

    if 'Codigo' in asig.columns:
        duplicados_asig = asig[asig.duplicated(subset=['Codigo'], keep=False)]
        if not duplicados_asig.empty:
            st.warning(f"Se eliminaron {duplicados_asig['Codigo'].nunique()} códigos duplicados en ASIGNACIONES; se usará el primer registro por código.")
            asig = asig.drop_duplicates(subset=['Codigo'], keep='first').copy()
    else:
        st.error("El archivo ASIGNACIONES debe tener la columna 'Codigo'.")
    
    pedidos['key'] = pedidos['Localidad'].apply(normalizar_texto)
    
    reglas_tmp = reglas.copy()
    reglas_tmp.columns = ['key', 'Codigo'] # Aseguramos nombres
    reglas_tmp['key'] = reglas_tmp['key'].apply(normalizar_texto)
    
    df = pedidos.merge(reglas_tmp[['key', 'Codigo']], on='key', how='left')
    df = df.merge(asig, on='Codigo', how='left')
    
    if '_orig_index' in df.columns:
        df = df.sort_values(by=['_orig_index']).drop_duplicates(subset=['_orig_index'], keep='first').copy()

    if 'key' in df.columns:
        df = df.drop(columns=['key'])
    
    if 'Chofer' in df.columns:
        df['Chofer'] = df['Chofer'].fillna('Sin Chofer').astype(str).str.strip()
    else:
        df['Chofer'] = 'Sin Chofer'

    mask_sin_chofer = df['Chofer'].str.strip().str.lower().isin(['', 'sin chofer', 'none'])
    df_sin_chofer = df[mask_sin_chofer].copy()
    df = df[~mask_sin_chofer].copy()

    if '_orig_index' in df.columns:
        df = df.drop(columns=['_orig_index'])
    if '_orig_index' in df_sin_chofer.columns:
        df_sin_chofer = df_sin_chofer.drop(columns=['_orig_index'])
        
    st.write("### Resultados en pantalla")
    st.dataframe(df)
    
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

    from geopy.geocoders import ArcGIS
    from concurrent.futures import ThreadPoolExecutor, as_completed
    
    geolocator_arcgis = ArcGIS()

    # Geocoders y límites de cobertura (Flex AMBA)
    geocoders = [geo_arcgis, geo_photon, geo_osm]
    LAT_MIN, LAT_MAX = -37.00, -33.00
    LON_MIN, LON_MAX = -60.50, -57.00

    def buscar_con_un_geocoder(geocoder, busquedas):
        for busqueda in busquedas:
            try:
                location = geocoder.geocode(busqueda, timeout=4)
                if location:
                    lat, lon = location.latitude, location.longitude
                    if LAT_MIN <= lat <= LAT_MAX and LON_MIN <= lon <= LON_MAX:
                        return lat, lon
            except:
                pass
        return (None, None)

    @st.cache_data
    def buscar_coordenadas(direccion, localidad, cp):
        direccion_limpia = limpiar_direccion_para_mapa(direccion)
        if pd.isna(direccion) or pd.isna(localidad):
            return (None, None)

        cp_str = str(cp).strip() if not pd.isna(cp) else ""
        localidad_str = str(localidad).strip()

        busquedas = []
        if cp_str:
            busquedas.extend([
                f"{direccion_limpia}, {cp_str}, Buenos Aires, Argentina",
                f"{localidad_str}, {cp_str}, Buenos Aires, Argentina",
            ])
        busquedas.extend([
            f"{direccion_limpia}, {localidad_str}, Buenos Aires, Argentina",
            f"{localidad_str}, Buenos Aires, Argentina",
        ])

        # Intentar cada geocoder hasta encontrar una ubicación válida dentro de bounds
        for g in geocoders:
            latlon = buscar_con_un_geocoder(g, busquedas)
            if latlon[0] is not None:
                return latlon
        return (None, None)

    # BARRA DE PROGRESO
    progreso_bar = st.progress(0)
    total = len(df)
    procesadas = 0
    
    resultados_lat = {}
    resultados_lon = {}

    with ThreadPoolExecutor(max_workers=3) as executor:
        futures = {executor.submit(procesar_fila, idx, row): idx for idx, row in df.iterrows()}
        
        for future in as_completed(futures):
            try:
                idx, lat, lon = future.result()
                resultados_lat[idx] = lat
                resultados_lon[idx] = lon
            except:
                pass

            procesadas += 1
            progreso_bar.progress(procesadas / total)

    # 4. Asignamos los resultados al DataFrame una vez que el bucle terminó
    df['Lat'] = df.index.map(resultados_lat)
    df['Lon'] = df.index.map(resultados_lon)

    df_mapa = df.dropna(subset=['Lat', 'Lon'])
    m = folium.Map(location=[-34.6, -58.4], zoom_start=10, tiles='OpenStreetMap')
    
    todos_choferes = sorted(df['Chofer'].unique())
    num_choferes = len(todos_choferes)
    
    colores_folium = ['red', 'blue', 'green', 'purple', 'orange', 'darkred', 'darkblue', 
                      'darkgreen', 'cadetblue', 'darkpurple', 'pink', 'gray', 'lightred', 
                      'lightblue', 'lightgreen', 'lightgray', 'beige', 'lightcyan', 'black']
    
    if num_choferes > len(colores_folium):
        colores_disponibles = (colores_folium * ((num_choferes // len(colores_folium)) + 1))[:num_choferes]
    else:
        colores_disponibles = colores_folium[:num_choferes]
    
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

    # ANCHO DE MAPA!!!
    st_folium(m, width=1400, height=700)

    #DASH
    st.write("---")
    st.write("### Distribución de Paquetes por Chofer")
    
    # Contar
    total_pedidos_reales = len(df)  # Total real de pedidos con chofer
    paquetes_por_chofer = df_mapa['Chofer'].value_counts().sort_values(ascending=True)
    total_paquetes_geocodificados = paquetes_por_chofer.sum()
    
    #PLOTLY!!!!
    import plotly.graph_objects as go
    
    choferes = paquetes_por_chofer.index.tolist()
    cantidades = paquetes_por_chofer.values.tolist()
    colores = [mapa_colores_chofer_hex[chofer] for chofer in choferes]
    
    fig = go.Figure(
        data=[go.Bar(
            x=choferes,
            y=cantidades,
            orientation='v',
            marker=dict(color=colores),
            text=[f"{int(c)}" for c in cantidades],
            textposition='outside',
            hovertemplate='<b>%{y}</b><br>Paquetes: %{x}<extra></extra>'
        )]
    )
    
    fig.update_layout(
        title=dict(text=f"Total en mapa: {total_paquetes_geocodificados}/{total_pedidos_reales} paquetes", font=dict(color='#ffffff', size=16)),
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

    # Botón de prueba: mostrar cómo se repartirían las primeras filas entre geocoders
    if st.button('Probar round-robin geocoders (6 filas)'):
        prueba = []
        sample = df.head(6).reset_index()
        for i, row in sample.iterrows():
            direccion = row['Direccion']
            localidad = row['Localidad']
            cp = row.get('cp', '') if 'cp' in row else ''
            direccion_limpia = limpiar_direccion_para_mapa(direccion)
            localidad_str = str(localidad).strip()
            cp_str = str(cp).strip() if not pd.isna(cp) else ''
            busquedas = []
            if cp_str:
                busquedas.extend([
                    f"{direccion_limpia}, {cp_str}, Buenos Aires, Argentina",
                    f"{localidad_str}, {cp_str}, Buenos Aires, Argentina",
                ])
            busquedas.extend([
                f"{direccion_limpia}, {localidad_str}, Buenos Aires, Argentina",
                f"{localidad_str}, Buenos Aires, Argentina",
            ])
            g = geocoders[i % len(geocoders)]
            lat, lon = buscar_con_un_geocoder(g, busquedas)
            prueba.append({
                'idx': row['index'],
                'Direccion': direccion,
                'Geocoder': g.__class__.__name__,
                'Lat': lat,
                'Lon': lon
            })
        st.table(pd.DataFrame(prueba))
    
    fig.update_traces(textfont=dict(color='#ffffff', size=11))
    
    st.plotly_chart(fig, use_container_width=True)

    # LISTADO DE ERRORES
    df_errores = df[df['Lat'].isna()]

    if not df_errores.empty:
        st.error(f"⚠️ Atención: No se pudieron geolocalizar {len(df_errores)} direcciones:")
        st.table(df_errores[['Direccion', 'Localidad']])
    else:
        st.success("✅ ¡Todas las direcciones fueron ubicadas correctamente!")
    
    if 'df_sin_chofer' in locals() and not df_sin_chofer.empty:
        st.warning(f"⚠️ Hay {len(df_sin_chofer)} filas sin chofer asignado (no procesadas en el mapa):")
        cols_show = [c for c in ['Direccion', 'Localidad', 'Codigo', 'Chofer'] if c in df_sin_chofer.columns]
        st.table(df_sin_chofer[cols_show])
    