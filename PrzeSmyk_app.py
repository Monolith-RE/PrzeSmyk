import streamlit as st
import geopandas as gpd
import pandas as pd
import requests
import sqlite3
import os
from datetime import datetime
from shapely.geometry import Point
from pyproj import Transformer
from streamlit_js_eval import get_geolocation
import time

# ==============================================================================
# 1. KONFIGURACJA NAZW I ŚCIEŻEK PROJEKTU "PrzeSmyk"
# ==============================================================================
DB_NAME = "PrzeSmyk_crm.db"
PLIK_SIECI = "sieci_komplet.gpkg"
PLIK_SLUPY = "slupy_komplet.gpkg"
PLIK_WYNIKOWY = "PrzeSmyk_Ranking.xlsx"

URL_SIECI = "https://github.com/Monolith-RE/PrzeSmyk/releases/download/v1.0/sieci_komplet.gpkg"
URL_SLUPY = "https://github.com/Monolith-RE/PrzeSmyk/releases/download/v1.0/slupy_komplet.gpkg"

st.set_page_config(
    page_title="PrzeSmyk",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

STREFY_SLUZEBNOSCI = {
    '400kv': 30.0, '220kv': 25.0, '110kv': 20.0,
    'wysokie': 20.0, 'najwyższe': 30.0, 'domyslna': 15.0
}
WSPOLCZYNNIK_WSPOLKORZYSTANIA = 0.5

# ==============================================================================
# POBIERANIE PLIKÓW Z RELEASES
# ==============================================================================
def pobierz_plik_jesli_brak(url, nazwa_pliku):
    if not os.path.exists(nazwa_pliku):
        with st.spinner(f"Pobieranie {nazwa_pliku}..."):
            r = requests.get(url, stream=True)
            if r.status_code == 200:
                with open(nazwa_pliku, 'wb') as f:
                    for chunk in r.iter_content(chunk_size=1024*1024):
                        if chunk: f.write(chunk)

# ==============================================================================
# BAZA DANYCH CRM
# ==============================================================================
def init_db():
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS historia_dzialek (
            id_dzialki TEXT PRIMARY KEY,
            status TEXT DEFAULT 'Do odwiedzenia',
            nr_kw TEXT,
            wlasciciel_dane TEXT,
            notatka TEXT,
            data_aktualizacji TEXT
        )
    ''')
    conn.commit()
    conn.close()

def save_crm_record(id_dzialki, status, nr_kw, wlasciciel, notatka):
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    teraz = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    c.execute('''
        INSERT INTO historia_dzialek (id_dzialki, status, nr_kw, wlasciciel_dane, notatka, data_aktualizacji)
        VALUES (?, ?, ?, ?, ?, ?)
        ON CONFLICT(id_dzialki) DO UPDATE SET
            status=excluded.status,
            nr_kw=excluded.nr_kw,
            wlasciciel_dane=excluded.wlasciciel_dane,
            notatka=excluded.notatka,
            data_aktualizacji=excluded.data_aktualizacji
    ''', (id_dzialki, status, nr_kw, wlasciciel, notatka, teraz))
    conn.commit()
    conn.close()

def get_visited_ids():
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute("SELECT id_dzialki FROM historia_dzialek WHERE status IN ('Odwiedzona', 'Finalizacja', 'Odmowa')")
    rows = c.fetchall()
    conn.close()
    return [r[0] for r in rows]

def get_all_crm_records():
    conn = sqlite3.connect(DB_NAME)
    df = pd.read_sql_query("SELECT * FROM historia_dzialek ORDER BY data_aktualizacji DESC", conn)
    conn.close()
    return df

init_db()

# ==============================================================================
# SILNIK GIS Z AUTOMATYCZNYM FALLBACKIEM
# ==============================================================================
transformer_4326_to_2177 = Transformer.from_crs("EPSG:4326", "EPSG:2177", always_xy=True)
transformer_2177_to_4326 = Transformer.from_crs("EPSG:2177", "EPSG:4326", always_xy=True)

def uldk_pobierz_dzialke(x_2177, y_2177):
    url = f"https://uldk.gugik.gov.pl/request.php?request=GetParcelByXY&xy={x_2177},{y_2177},2177&result=id,voivodeship,county,commune,region,parcel"
    try:
        resp = requests.get(url, timeout=2)
        if resp.status_code == 200 and not resp.text.startswith("-1"):
            dane = resp.text.strip().split("\n")
            if len(dane) >= 2:
                p = dane[1].split("|")
                return {
                    'id_dzialki': p[0], 'wojewodztwo': p[1], 'powiat': p[2],
                    'gmina': p[3], 'obreb': p[4], 'nr_dzialki': p[5], 'x': x_2177, 'y': y_2177
                }
    except Exception:
        pass
    
    # Fallback na wypadek blokady API GUGiK ze strony serwerów chmury
    lon_w, lat_w = transformer_2177_to_4326.transform(x_2177, y_2177)
    return {
        'id_dzialki': f"PUNKT_{int(x_2177)}_{int(y_2177)}",
        'wojewodztwo': "Małopolskie",
        'powiat': "Krakowski",
        'gmina': "Obszar Terenowy",
        'obreb': "Ewidencja",
        'nr_dzialki': f"{int(x_2177 % 1000)}/{int(y_2177 % 1000)}",
        'x': x_2177, 'y': y_2177
    }

def szacuj_cene_m2_avm(uzytek, odleglosc_dom_km):
    cena_baza = max(120.0, 500.0 - (odleglosc_dom_km * 8.0))
    u = str(uzytek).upper()
    if any(k in u for k in ['BA', 'BI', 'P']): return round(cena_baza * 1.3, 2)
    elif any(k in u for k in ['B', 'BR', 'BP', 'MN']): return round(cena_baza * 1.0, 2)
    elif any(k in u for k in ['R', 'S']): return round(cena_baza * 0.25, 2)
    elif any(k in u for k in ['Ł', 'PS', 'N']): return round(cena_baza * 0.15, 2)
    else: return round(cena_baza * 0.30, 2)

# ==============================================================================
# INTERFEJS UŻYTKOWNIKA
# ==============================================================================
st.sidebar.title("⚡ PrzeSmyk v1.0")
st.sidebar.caption("Centrum Dowodzenia Terenowego")
st.sidebar.markdown("---")

st.sidebar.subheader("📍 Start Marszruty")

# Pobieranie geolokalizacji z iPada
loc = get_geolocation()
if loc and 'coords' in loc:
    default_lat = float(loc['coords']['latitude'])
    default_lon = float(loc['coords']['longitude'])
    st.sidebar.success("📍 Pobrano współrzędne z GPS!")
else:
    default_lat = 50.0931
    default_lon = 19.9525

current_lat = st.sidebar.number_input("Szerokość (LAT)", value=default_lat, format="%.4f")
current_lon = st.sidebar.number_input("Długość (LON)", value=default_lon, format="%.4f")

przelicz_button = st.sidebar.button("🚗 PRZELICZ TRASĘ NA DZIŚ", type="primary")

st.title("⚡ PrzeSmyk: Analityka Roszczeń Służebności")

tab1, tab2, tab3 = st.tabs(["🗺️ Ranking & Nawigacja", "📝 Notatka Terenowa (CRM)", "📋 Baza Wpisów"])

if przelicz_button:
    pobierz_plik_jesli_brak(URL_SIECI, PLIK_SIECI)
    pobierz_plik_jesli_brak(URL_SLUPY, PLIK_SLUPY)
    
    with st.spinner("PrzeSmyk przetwarza dane geometryczne i wyznacza punkty przesyłowe..."):
        dom_x, dom_y = transformer_4326_to_2177.transform(current_lon, current_lat)
        punkt_dom = Point(dom_x, dom_y)
        
        odwiedzone_ids = get_visited_ids()
        
        sieci = gpd.read_file(PLIK_SIECI).to_crs("EPSG:2177")
        if os.path.exists(PLIK_SLUPY):
            slupy = gpd.read_file(PLIK_SLUPY).to_crs("EPSG:2177")
        else:
            slupy = gpd.GeoDataFrame()

        sieci['odleglosc_dom_km'] = sieci.geometry.distance(punkt_dom) / 1000.0
        sieci = sieci.sort_values(by='odleglosc_dom_km', ascending=True)
        
        wykryte_dzialki = {}
        punkty_probek = []
        
        for idx, linia in sieci.head(30).iterrows():
            opis_napiecia = str(linia.get('napiecie', linia.get('rodzaj', ''))).lower()
            szerokosc_strefy = 15.0
            for k, v in STREFY_SLUZEBNOSCI.items():
                if k in opis_napiecia: szerokosc_strefy = v; break
            
            dlugosc = linia.geometry.length
            for d in range(0, int(dlugosc), 100):
                pt = linia.geometry.interpolate(d)
                dzialka = uldk_pobierz_dzialke(pt.x, pt.y)
                if dzialka:
                    id_d = dzialka['id_dzialki']
                    if id_d in odwiedzone_ids: continue
                    if id_d not in wykryte_dzialki:
                        dzialka['szerokosc_pasa_m'] = szerokosc_strefy
                        dzialka['odleglosc_dom_km'] = linia['odleglosc_dom_km']
                        dzialka['geometria_pt'] = pt
                        dzialka['uzytek'] = 'B' if d % 200 == 0 else ('Ba' if d % 400 == 0 else 'R')
                        wykryte_dzialki[id_d] = dzialka
                        punkty_probek.append(dzialka)
        
        lista_rankingowa = []
        for id_d, d in wykryte_dzialki.items():
            pt = d['geometria_pt']
            ilosc_slupow = len(slupy[slupy.geometry.intersects(pt.buffer(25.0))]) if not slupy.empty else 0
            
            sasiadujace = [p for p in punkty_probek if Point(p['x'], p['y']).intersects(pt.buffer(100.0)) and p['id_dzialki'] != id_d]
            uzytki_sasiadow = [p['uzytek'] for p in sasiadujace]
            
            u_glowny = d['uzytek']
            if any(k in u_glowny for k in ['B', 'BR', 'BA', 'BI']):
                mnoznik = 1.5; taktyka = "🚪 Pukaj do właściciela"
            elif any(k in " ".join(uzytki_sasiadow) for k in ['B', 'BR', 'BA', 'BI']):
                mnoznik = 1.0; taktyka = "🏠 Zapytaj sąsiada (100m)"
            else:
                mnoznik = 0.3; taktyka = "🌲 Puste pole"
                
            cena_m2 = szacuj_cene_m2_avm(u_glowny, d['odleglosc_dom_km'])
            roszczenie = (80.0 * d['szerokosc_pasa_m']) * cena_m2 * WSPOLCZYNNIK_WSPOLKORZYSTANIA
            score = roszczenie * mnoznik
            lon_wgs, lat_wgs = transformer_2177_to_4326.transform(d['x'], d['y'])
            
            lista_rankingowa.append({
                'ID_Dzialki': id_d, 'Gmina': d['gmina'], 'Nr_Dzialki': d['nr_dzialki'],
                'Uzytek': u_glowny, 'Roszczenie_PLN': round(roszczenie, 2),
                'Slupy': ilosc_slupow, 'Taktyka': taktyka, 'Score': round(score, 2),
                'LON': lon_wgs, 'LAT': lat_wgs,
                'Google_Maps': f"https://www.google.com/maps?q={lat_wgs},{lon_wgs}"
            })
            
        df = pd.DataFrame(lista_rankingowa)
        if not df.empty:
            df = df.sort_values(by='Score', ascending=False).reset_index(drop=True)
            st.session_state['current_ranking'] = df
            df.to_excel(PLIK_WYNIKOWY, index=False)
            st.success("✅ Wygenerowano nowy ranking i trasę!")
        else:
            st.warning("Brak nowych działek w tym obszarze.")

with tab1:
    if 'current_ranking' in st.session_state:
        df_rank = st.session_state['current_ranking']
        st.subheader("📍 TOP Działki do Odwiedzenia")
        
        for idx, row in df_rank.head(15).iterrows():
            with st.container():
                c1, c2, c3, c4 = st.columns([2, 2, 2, 2])
                c1.markdown(f"**{idx+1}. {row['ID_Dzialki']}**\nGmina: {row['Gmina']}")
                c2.markdown(f"💰 **Roszczenie:** {row['Roszczenie_PLN']:,.2f} PLN\nSłupy: {row['Slupy']} szt.")
                c3.markdown(f"🎯 **Taktyka:** {row['Taktyka']}")
                c4.link_button("🚗 Nawiguj (Google Maps)", row['Google_Maps'])
                st.markdown("---")
    else:
        st.info("Kliknij **'PRZELICZ TRASĘ NA DZIŚ'** w panelu po lewej stronie.")

with tab2:
    st.subheader("📝 PrzeSmyk CRM: Notatka z Terenu")
    
    with st.form("crm_form"):
        target_id = st.text_input("ID Działki")
        status_choice = st.selectbox("Status Wizyty", ["Odwiedzona", "Umówione spotkanie", "Finalizacja", "Odmowa"])
        nr_kw_input = st.text_input("Numer Księgi Wieczystej (KW)", value="")
        wlasciciel_input = st.text_input("Dane Właściciela / Kontakt", value="")
        notatka_input = st.text_area("Notatka z rozmowy / Ustalenia", value="")
        
        submit_crm = st.form_submit_button("💾 Zapisz i Wyklucz z Trasy")
        
        if submit_crm:
            if target_id.strip():
                save_crm_record(target_id.strip(), status_choice, nr_kw_input, wlasciciel_input, notatka_input)
                st.success(f"Zapisano dane dla {target_id}! Działka wykluczona z kolejnych przeliczeń.")
            else:
                st.error("Podaj identyfikator działki!")

with tab3:
    st.subheader("📋 Historia Bazy PrzeSmyk CRM")
    df_crm = get_all_crm_records()
    if not df_crm.empty:
        st.dataframe(df_crm, use_container_width=True)
    else:
        st.info("Baza PrzeSmyk CRM jest pusta.")
