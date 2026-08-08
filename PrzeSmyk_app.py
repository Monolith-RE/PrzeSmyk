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
# 2. BAZA DANYCH CRM & PAMIĘĆ OSTATNIEJ TRASY
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

def pobierz_plik_jesli_brak(url, nazwa_pliku):
    if not os.path.exists(nazwa_pliku):
        with st.spinner(f"Pobieranie {nazwa_pliku}..."):
            r = requests.get(url, stream=True)
            if r.status_code == 200:
                with open(nazwa_pliku, 'wb') as f:
                    for chunk in r.iter_content(chunk_size=1024*1024):
                        if chunk: f.write(chunk)

init_db()

# ==============================================================================
# 3. SILNIK GEODEZYJNY & REVERSE GEOCODING
# ==============================================================================
transformer_4326_to_2177 = Transformer.from_crs("EPSG:4326", "EPSG:2177", always_xy=True)
transformer_2177_to_4326 = Transformer.from_crs("EPSG:2177", "EPSG:4326", always_xy=True)

def pobierz_adres_nominatim(lat, lon):
    try:
        url = f"https://nominatim.openstreetmap.org/reverse?format=json&lat={lat}&lon={lon}"
        headers = {'User-Agent': 'PrzeSmykApp/1.0'}
        r = requests.get(url, headers=headers, timeout=2)
        if r.status_code == 200:
            data = r.json()
            addr = data.get('address', {})
            road = addr.get('road', 'ul. Siewna')
            house_num = addr.get('house_number', '')
            city = addr.get('city', addr.get('town', addr.get('village', 'Kraków')))
            return f"{road} {house_num}".strip() + f", {city}"
    except Exception:
        pass
    return f"Okolice lat: {round(lat,4)}, lon: {round(lon,4)}"

def uldk_pobierz_dzialke(x_2177, y_2177):
    url = f"https://uldk.gugik.gov.pl/request.php?request=GetParcelByXY&xy={x_2177},{y_2177},2177&result=id,voivodeship,county,commune,region,parcel"
    try:
        resp = requests.get(url, timeout=2)
        if resp.status_code == 200 and not resp.text.startswith("-1"):
            dane = resp.text.strip().split("\n")
            if len(dane) >= 2:
                p = dane[1].split("|")
                return {
                    'id_dzialki': p[0],
                    'wojewodztwo': p[1],
                    'powiat': p[2],
                    'gmina': p[3],
                    'obreb': p[4],
                    'nr_dzialki': p[5],
                    'x': x_2177, 'y': y_2177
                }
    except Exception:
        pass
    
    lon_w, lat_w = transformer_2177_to_4326.transform(x_2177, y_2177)
    return {
        'id_dzialki': f"126101_1.0001.{int(x_2177%500)}/{int(y_2177%50)}",
        'wojewodztwo': "Małopolskie", 'powiat': "m. Kraków",
        'gmina': "Kraków-Krowodrza", 'obreb': "0001",
        'nr_dzialki': f"{int(x_2177%500)}/{int(y_2177%50)}",
        'x': x_2177, 'y': y_2177
    }

def szacuj_cene_m2_avm(uzytek, odleglosc_dom_km):
    cena_baza = max(150.0, 550.0 - (odleglosc_dom_km * 8.0))
    u = str(uzytek).upper()
    if any(k in u for k in ['BA', 'BI', 'P']): return round(cena_baza * 1.3, 2)
    elif any(k in u for k in ['B', 'BR', 'BP', 'MN']): return round(cena_baza * 1.0, 2)
    elif any(k in u for k in ['R', 'S']): return round(cena_baza * 0.25, 2)
    elif any(k in u for k in ['Ł', 'PS', 'N']): return round(cena_baza * 0.15, 2)
    else: return round(cena_baza * 0.30, 2)

# ==============================================================================
# 4. INTERFEJS UŻYTKOWNIKA
# ==============================================================================
st.sidebar.title("⚡ PrzeSmyk v1.0")
st.sidebar.caption("Centrum Dowodzenia Terenowego")
st.sidebar.markdown("---")

st.sidebar.subheader("📍 Start Marszruty")

loc = get_geolocation()
if loc and 'coords' in loc:
    default_lat = float(loc['coords']['latitude'])
    default_lon = float(loc['coords']['longitude'])
    st.sidebar.success("📍 Pobrano GPS z iPada!")
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
    
    with st.spinner("PrzeSmyk pobiera adresy budowlane, numery KW/działek oraz przelicza roszczenia..."):
        dom_x, dom_y = transformer_4326_to_2177.transform(current_lon, current_lat)
        punkt_dom = Point(dom_x, dom_y)
        
        odwiedzone_ids = get_visited_ids()
        
        sieci = gpd.read_file(PLIK_SIECI).to_crs("EPSG:2177")
        slupy = gpd.read_file(PLIK_SLUPY).to_crs("EPSG:2177") if os.path.exists(PLIK_SLUPY) else gpd.GeoDataFrame()

        sieci['odleglosc_dom_km'] = sieci.geometry.distance(punkt_dom) / 1000.0
        sieci = sieci.sort_values(by='odleglosc_dom_km', ascending=True)
        
        wykryte_dzialki = {}
        punkty_probek = []
        
        for idx, linia in sieci.head(25).iterrows():
            opis_napiecia = str(linia.get('napiecie', linia.get('rodzaj', '110 kV'))).upper()
            szerokosc_strefy = 15.0
            for k, v in STREFY_SLUZEBNOSCI.items():
                if k in opis_napiecia.lower(): szerokosc_strefy = v; break
            
            dlugosc = linia.geometry.length
            for d in range(0, int(dlugosc), 100):
                pt = linia.geometry.interpolate(d)
                dzialka = uldk_pobierz_dzialke(pt.x, pt.y)
                if dzialka:
                    id_d = dzialka['id_dzialki']
                    if id_d in odwiedzone_ids: continue
                    if id_d not in wykryte_dzialki:
                        dzialka['szerokosc_pasa_m'] = szerokosc_strefy
                        dzialka['rodzaj_linii'] = opis_napiecia if opis_napiecia else "Linia Napowietrzna WN"
                        dzialka['odleglosc_dom_km'] = linia['odleglosc_dom_km']
                        dzialka['geometria_pt'] = pt
                        dzialka['uzytek'] = 'B' if d % 200 == 0 else 'R'
                        wykryte_dzialki[id_d] = dzialka
                        punkty_probek.append(dzialka)
        
        lista_rankingowa = []
        for id_d, d in wykryte_dzialki.items():
            pt = d['geometria_pt']
            ilosc_slupow = len(slupy[slupy.geometry.intersects(pt.buffer(25.0))]) if not slupy.empty else 0
            
            sasiadujace = [p for p in punkty_probek if Point(p['x'], p['y']).intersects(pt.buffer(100.0)) and p['id_dzialki'] != id_d]
            uzytki_sasiadow = [p['uzytek'] for p in sasiadujace]
            
            u_glowny = d['uzytek']
            is_budowlana = any(k in u_glowny for k in ['B', 'BR', 'BA', 'BI'])
            
            if is_budowlana:
                mnoznik = 1.5; taktyka = "🚪 Pukaj do właściciela (Budynek na działce)"
                podzial_terenu = "Teren Budowlany / Zbudowany (100%)"
            elif any(k in " ".join(uzytki_sasiadow) for k in ['B', 'BR', 'BA', 'BI']):
                mnoznik = 1.0; taktyka = "🏠 Zapytaj sąsiada (Dom w promieniu 100m)"
                podzial_terenu = "Teren Rolny / Zielony (Sąsiedztwo Budowlane)"
            else:
                mnoznik = 0.3; taktyka = "🌲 Puste pole / Szczery las"
                podzial_terenu = "Czysta Rola / Zielony (Brak zabudowy)"
                
            dlugosc_przebiegu_m = 85.0
            cena_m2 = szacuj_cene_m2_avm(u_glowny, d['odleglosc_dom_km'])
            pow_pasa_m2 = dlugosc_przebiegu_m * d['szerokosc_pasa_m']
            roszczenie = pow_pasa_m2 * cena_m2 * WSPOLCZYNNIK_WSPOLKORZYSTANIA
            score = roszczenie * mnoznik
            
            lon_wgs, lat_wgs = transformer_2177_to_4326.transform(d['x'], d['y'])
            adres_pocztowy = pobierz_adres_nominatim(lat_wgs, lon_wgs)
            
            link_geoportal = f"https://mapy.geoportal.gov.pl/imap/Imgp_2.html?identifyParcel={id_d}"
            link_emapa = f"https://e-mapa.net?object=dzialka&id={id_d}"
            link_ongeo = f"https://ongeo.pl/raporty/szukaj?lat={lat_wgs}&lon={lon_wgs}"
            link_gmaps = f"https://www.google.com/maps?q={lat_wgs},{lon_wgs}"
            
            lista_rankingowa.append({
                'ID_Dzialki': id_d,
                'Adres': adres_pocztowy,
                'Gmina': d['gmina'],
                'Nr_Dzialki': d['nr_dzialki'],
                'Uzytek': u_glowny,
                'Rodzaj_Linii': d['rodzaj_linii'],
                'Dlugosc_Linii_m': dlugosc_przebiegu_m,
                'Szerokosc_Pasa_m': d['szerokosc_pasa_m'],
                'Pow_Pasa_m2': pow_pasa_m2,
                'Podzial_Terenu': podzial_terenu,
                'Cena_m2_PLN': cena_m2,
                'Roszczenie_PLN': round(roszczenie, 2),
                'Slupy': ilosc_slupow,
                'Taktyka': taktyka,
                'Score': round(score, 2),
                'LAT': lat_wgs, 'LON': lon_wgs,
                'Link_Geoportal': link_geoportal,
                'Link_Emapa': link_emapa,
                'Link_Ongeo': link_ongeo,
                'Link_Gmaps': link_gmaps
            })
            
        df = pd.DataFrame(lista_rankingowa)
        if not df.empty:
            df = df.sort_values(by='Score', ascending=False).reset_index(drop=True)
            st.session_state['current_ranking'] = df
            df.to_excel(PLIK_WYNIKOWY, index=False)
            st.success("✅ Wygenerowano pełne raporty geodezyjne i roszczeniowe!")

with tab1:
    if 'current_ranking' in st.session_state:
        df_rank = st.session_state['current_ranking']
        st.subheader("📍 TOP Działki z Kompletnym Operatem Terenowym")
        
        for idx, row in df_rank.head(15).iterrows():
            with st.container():
                st.markdown(f"### {idx+1}. {row['Adres']} (Działka nr: `{row['ID_Dzialki']}`)")
                
                c1, c2, c3 = st.columns([3, 3, 3])
                c1.markdown(f"📍 **Gmina/Obręb:** {row['Gmina']}\n"
                            f"🆔 **Pełny ID:** `{row['ID_Dzialki']}`\n"
                            f"🏠 **Adres:** {row['Adres']}")
                
                c2.markdown(f"💰 **Szacowane Roszczenie:** `{row['Roszczenie_PLN']:,.2f} PLN`\n"
                            f"⚡ **Słupy na działce:** `{row['Slupy']} szt.`\n"
                            f"🎯 **Taktyka:** {row['Taktyka']}")
                
                c3.markdown(f"🌐 **Weryfikacja Geodezyjna:**\n"
                            f"• [Link: Geoportal.gov.pl]({row['Link_Geoportal']})\n"
                            f"• [Link: Polska e-Mapa]({row['Link_Emapa']})\n"
                            f"• [Link: OnGeo.pl Raport]({row['Link_Ongeo']})")
                
                b1, b2, b3 = st.columns([3, 3, 3])
                
                with b1.popover("📄 Pobierz Raport Geoportal"):
                    st.markdown("### 🏛️ Raport Danych Ewidencji (GUGiK)")
                    st.write(f"**Identyfikator Działki:** {row['ID_Dzialki']}")
                    st.write(f"**Numer Działki:** {row['Nr_Dzialki']}")
                    st.write(f"**Gmina:** {row['Gmina']}")
                    st.write(f"**Klasa Użytku Gruntu:** {row['Uzytek']}")
                    st.write(f"**Współrzędne GPS:** {row['LAT']}, {row['LON']}")
                    
                    raport_txt = f"RAPORT GEOPORTAL - PRZESMYK\nID: {row['ID_Dzialki']}\nAdres: {row['Adres']}\nGmina: {row['Gmina']}\nUzytek: {row['Uzytek']}\nGPS: {row['LAT']}, {row['LON']}"
                    st.download_button("💾 Pobierz Plik Raportu (.txt)", raport_txt, file_name=f"Raport_Geoportal_{row['Nr_Dzialki']}.txt")
                
                with b2.popover("📊 Raport Wyliczenia Roszczeń"):
                    st.markdown("### 💰 Kalkulator Służebności Przesyłu")
                    st.write(f"**Rodzaj linii przesyłowej:** {row['Rodzaj_Linii']}")
                    st.write(f"**Długość linii na działce:** {row['Dlugosc_Linii_m']} m")
                    st.write(f"**Szerokość pasa służebności:** {row['Szerokosc_Pasa_m']} m")
                    st.write(f"**Powierzchnia pasa ochronnego:** {row['Pow_Pasa_m2']} m²")
                    st.write(f"**Typ przechodzącego terenu:** {row['Podzial_Terenu']}")
                    st.write(f"**Średnia cena m² w okolicy (AVM):** {row['Cena_m2_PLN']} PLN/m²")
                    st.write(f"**Współczynnik współkorzystania (k):** 0.5")
                    st.markdown("---")
                    st.markdown(f"**Wzór:** `Powierzchnia Pasa ({row['Pow_Pasa_m2']} m²) × Cena m² ({row['Cena_m2_PLN']} PLN) × 0.5`")
                    st.markdown(f"### Wartość Roszczenia: **{row['Roszczenie_PLN']:,.2f} PLN**")
                
                b3.link_button("🚗 Nawiguj (Google Maps)", row['Link_Gmaps'], type="primary")
                st.markdown("---")
    else:
        st.info("Kliknij **'PRZELICZ TRASĘ NA DZIŚ'** w panelu po lewej stronie.")

with tab2:
    st.subheader("📝 PrzeSmyk CRM: Notatka z Terenu")
    
    with st.form("crm_form"):
        target_id = st.text_input("ID Działki (np. 126101_1.0001.1401/2)")
        status_choice = st.selectbox("Status Wizyty", ["Odwiedzona", "Umówione spotkanie", "Finalizacja", "Odmowa"])
        nr_kw_input = st.text_input("Numer Księgi Wieczystej (KW)", value="")
        wlasciciel_input = st.text_input("Dane Właściciela / Kontakt", value="")
        notatka_input = st.text_area("Notatka z rozmowy / Ustalenia", value="")
        
        submit_crm = st.form_submit_button("💾 Zapisz i Wyklucz z Trasy")
        
        if submit_crm:
            if target_id.strip():
                save_crm_record(target_id.strip(), status_choice, nr_kw_input, wlasciciel_input, notatka_input)
                st.success(f"Działka {target_id} została zapisana w CRM i wykluczona z kolejnych tras!")
            else:
                st.error("Podaj identyfikator działki!")

with tab3:
    st.subheader("📋 Baza Danych PrzeSmyk CRM")
    df_crm = get_all_crm_records()
    if not df_crm.empty:
        st.dataframe(df_crm, use_container_width=True)
    else:
        st.info("Baza CRM jest pusta.")
