import streamlit as st
import geopandas as gpd
import pandas as pd
import requests
import sqlite3
from datetime import datetime
from shapely.geometry import Point
from pyproj import Transformer
import time

DB_NAME = "PrzeSmyk_crm.db"
PLIK_SIECI = "sieci_komplet.gpkg"
PLIK_SLUPY = "slupy_komplet.gpkg"

st.set_page_config(page_title="PrzeSmyk", page_icon="⚡", layout="wide")

STREFY_SLUZEBNOSCI = {'400kv': 30.0, '220kv': 25.0, '110kv': 20.0, 'domyslna': 15.0}

def init_db():
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS historia_dzialek 
                 (id_dzialki TEXT PRIMARY KEY, status TEXT, nr_kw TEXT, wlasciciel_dane TEXT, notatka TEXT, data_aktualizacji TEXT)''')
    conn.commit(); conn.close()

def save_crm(id_d, status, kw, wlasciciel, notatka):
    conn = sqlite3.connect(DB_NAME); c = conn.cursor()
    teraz = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    c.execute('''INSERT INTO historia_dzialek VALUES (?,?,?,?,?,?)
                 ON CONFLICT(id_dzialki) DO UPDATE SET status=?, nr_kw=?, wlasciciel_dane=?, notatka=?, data_aktualizacji=?''',
              (id_d, status, kw, wlasciciel, notatka, teraz, status, kw, wlasciciel, notatka, teraz))
    conn.commit(); conn.close()

def get_visited():
    conn = sqlite3.connect(DB_NAME); c = conn.cursor()
    c.execute("SELECT id_dzialki FROM historia_dzialek WHERE status IN ('Odwiedzona', 'Finalizacja', 'Odmowa')")
    res = [r[0] for r in c.fetchall()]; conn.close(); return res

init_db()
transformer_4326_to_2177 = Transformer.from_crs("EPSG:4326", "EPSG:2177", always_xy=True)
transformer_2177_to_4326 = Transformer.from_crs("EPSG:2177", "EPSG:4326", always_xy=True)

st.title("⚡ PrzeSmyk - Centrum Dowodzenia")

tab1, tab2, tab3 = st.tabs(["🗺️ Trasa & Ranking", "📝 Notatka CRM", "📋 Baza"])

st.sidebar.header("📍 Pozycja Jeepa")
current_lat = st.sidebar.number_input("LAT", value=50.0931, format="%.4f")
current_lon = st.sidebar.number_input("LON", value=19.9525, format="%.4f")

if st.sidebar.button("🚗 PRZELICZ TRASĘ", type="primary"):
    with st.spinner("PrzeSmyk przelicza dane..."):
        dom_x, dom_y = transformer_4326_to_2177.transform(current_lon, current_lat)
        punkt_dom = Point(dom_x, dom_y)
        odwiedzone = get_visited()
        
        sieci = gpd.read_file(PLIK_SIECI).to_crs("EPSG:2177")
        sieci['dist'] = sieci.geometry.distance(punkt_dom) / 1000.0
        sieci = sieci.sort_values(by='dist')
        
        wykryte = []
        for idx, linia in sieci.head(15).iterrows():
            dl = linia.geometry.length
            for d in range(0, int(dl), 80):
                pt = linia.geometry.interpolate(d)
                lon_w, lat_w = transformer_2177_to_4326.transform(pt.x, pt.y)
                id_temp = f"DZIALKA_{int(pt.x)}_{int(pt.y)}"
                if id_temp not in odwiedzone:
                    wykryte.append({
                        'ID_Dzialki': id_temp, 'Roszczenie_PLN': 45000.0,
                        'LAT': lat_w, 'LON': lon_w,
                        'Google_Maps': f"https://www.google.com/maps?q={lat_w},{lon_w}"
                    })
        st.session_state['rank'] = pd.DataFrame(wykryte).head(10)
        st.success("✅ Trasa gotowa!")

with tab1:
    if 'rank' in st.session_state:
        for idx, row in st.session_state['rank'].iterrows():
            c1, c2, c3 = st.columns([3, 3, 2])
            c1.write(f"**{row['ID_Dzialki']}**")
            c2.write(f"💰 Roszczenie: {row['Roszczenie_PLN']} PLN")
            c3.link_button("🚗 Nawiguj", row['Google_Maps'])
            st.divider()

with tab2:
    with st.form("crm"):
        tid = st.text_input("ID Działki")
        stat = st.selectbox("Status", ["Odwiedzona", "Umówione spotkanie", "Odmowa"])
        kw = st.text_input("KW")
        wlas = st.text_input("Właściciel")
        notat = st.text_area("Notatka")
        if st.form_submit_button("💾 Zapisz"):
            save_crm(tid, stat, kw, wlas, notat)
            st.success("Zapisano!")

with tab3:
    conn = sqlite3.connect(DB_NAME)
    st.dataframe(pd.read_sql_query("SELECT * FROM historia_dzialek", conn))
    conn.close()
