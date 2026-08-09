import streamlit as st
import geopandas as gpd
import pandas as pd
import requests
import sqlite3
import os
import re
from datetime import datetime
from shapely.geometry import Point
from shapely.wkt import loads as wkt_loads
from pyproj import Transformer
import time

# ==============================================================================
# 1. KONFIGURACJA PROJEKTU "PrzeSmyk"
# ==============================================================================
DB_NAME = "PrzeSmyk_crm.db"
PLIK_SIECI = "sieci_komplet.gpkg"
PLIK_SLUPY = "slupy_komplet.gpkg"
PLIK_WYNIKOWY = "PrzeSmyk_Ranking.xlsx"

URL_SIECI = "https://github.com/Monolith-RE/PrzeSmyk/releases/download/v1.0/sieci_komplet.gpkg"
URL_SLUPY = "https://github.com/Monolith-RE/PrzeSmyk/releases/download/v1.0/slupy_komplet.gpkg"

st.set_page_config(
    page_title="PrzeSmyk",
    page_icon="🚙",
    layout="wide",
    initial_sidebar_state="expanded"
)

STREFY_SLUZEBNOSCI = {'400kv': 30.0, '220kv': 25.0, '110kv': 20.0, 'wysokie': 20.0, 'najwyższe': 30.0, 'domyslna': 15.0}
WSPOLCZYNNIK_WSPOLKORZYSTANIA = 0.5

CZARNA_LISTA = ['osiedle', 'os.', 'blok', 'bloki', 'apartament', 'apartments', 'flats', 'wielorodzinny', 'spółdzielnia']

# ==============================================================================
# 2. CRM BAZA DANYCH
# ==============================================================================
def init_db():
    conn = sqlite3.connect(DB_NAME); c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS historia_dzialek (id_dzialki TEXT PRIMARY KEY, status TEXT, nr_kw TEXT, wlasciciel_dane TEXT, notatka TEXT, data_aktualizacji TEXT)''')
    conn.commit(); conn.close()

def save_crm_record(id_dzialki, status, nr_kw, wlasciciel, notatka):
    conn = sqlite3.connect(DB_NAME); c = conn.cursor(); teraz = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    c.execute('''INSERT INTO historia_dzialek VALUES (?, ?, ?, ?, ?, ?) ON CONFLICT(id_dzialki) DO UPDATE SET status=excluded.status, nr_kw=excluded.nr_kw, wlasciciel_dane=excluded.wlasciciel_dane, notatka=excluded.notatka, data_aktualizacji=excluded.data_aktualizacji''', (id_dzialki, status, nr_kw, wlasciciel, notatka, teraz))
    conn.commit(); conn.close()

def get_visited_ids():
    conn = sqlite3.connect(DB_NAME); c = conn.cursor()
    c.execute("SELECT id_dzialki FROM historia_dzialek WHERE status IN ('Odwiedzona', 'Finalizacja', 'Odmowa')")
    rows = c.fetchall(); conn.close(); return [r[0] for r in rows]

def get_all_crm_records():
    conn = sqlite3.connect(DB_NAME)
    df = pd.read_sql_query("SELECT * FROM historia_dzialek ORDER BY data_aktualizacji DESC", conn)
    conn.close()
    return df

def pobierz_plik_jesli_brak(url, nazwa_pliku):
    if not os.path.exists(nazwa_pliku):
        with st.spinner(f"Pobieranie bazy {nazwa_pliku}..."):
            r = requests.get(url, stream=True)
            if r.status_code == 200:
                with open(nazwa_pliku, 'wb') as f:
                    for chunk in r.iter_content(chunk_size=1024*1024):
                        if chunk: f.write(chunk)
init_db()

# ==============================================================================
# 3. SILNIK GIS Z OGRANICZENIEM DO 5 KM
# ==============================================================================
transformer_4326_to_2177 = Transformer.from_crs("EPSG:4326", "EPSG:2177", always_xy=True)
transformer_2177_to_4326 = Transformer.from_crs("EPSG:2177", "EPSG:4326", always_xy=True)

def geokoduj_wpis_startowy(tekst_wpisu):
    if not tekst_wpisu or not tekst_wpisu.strip(): return 50.0931, 19.9525, "Kraków"
    tekst = tekst_wpisu.strip()
    match = re.search(r'(-?\d+\.\d+)[\s,]+(-?\d+\.\d+)', tekst)
    if match: return float(match.group(1)), float(match.group(2)), f"GPS ({float(match.group(1)):.4f}, {float(match.group(2)):.4f})"
    try:
        url = f"https://nominatim.openstreetmap.org/search?q={requests.utils.quote(tekst)}&format=json&limit=1"
        r = requests.get(url, headers={'User-Agent': 'PrzeSmykApp/9.0'}, timeout=3)
        if r.status_code == 200 and len(r.json()) > 0:
            res = r.json()[0]
            return float(res['lat']), float(res['lon']), res.get('display_name', tekst).split(',')[0]
    except Exception: pass
    return 50.0931, 19.9525, "Kraków (Domyślnie)"

def pobierz_adres_i_filtr_zabudowy(lat, lon, nr_dzialki_ewidencja):
    try:
        url = f"https://nominatim.openstreetmap.org/reverse?format=json&lat={lat}&lon={lon}&zoom=18&addressdetails=1&extratags=1"
        r = requests.get(url, headers={'User-Agent': 'PrzeSmykApp/9.0'}, timeout=2)
        if r.status_code == 200:
            data = r.json()
            raw_text = str(data).lower()
            
            if any(s in raw_text for s in CZARNA_LISTA): return "", False, "Osiedle / Bloki"
            if data.get('extratags', {}).get('building') in ['apartments', 'residential', 'dormitory', 'terrace']: 
                return "", False, "Wielorodzinny"

            addr = data.get('address', {})
            miasto = addr.get('city') or addr.get('town') or addr.get('village') or "Kraków"
            ulica = addr.get('road') or ""
            numer = addr.get('house_number') or ""
            
            adres_czysty = f"{miasto}, ul. {ulica} {numer}".strip() if ulica and numer else f"{miasto}, dz. nr {nr_dzialki_ewidencja}"
            
            if any(b in raw_text for b in ['commercial', 'industrial', 'warehouse', 'company']):
                return adres_czysty, True, "🏭 Firma / Przemysł"
            return adres_czysty, True, "🏠 Dom Jednorodzinny / Posesja"
    except Exception: pass
    return f"Kraków, dz. nr {nr_dzialki_ewidencja}", True, "🏠 Posesja"

def uldk_pobierz_dzialke_z_geometria(x_2177, y_2177):
    url = f"https://uldk.gugik.gov.pl/request.php?request=GetParcelByXY&xy={x_2177},{y_2177},2177&result=id,commune,parcel,geom_wkt"
    try:
        resp = requests.get(url, timeout=3)
        if resp.status_code == 200 and not resp.text.startswith("-1"):
            lines = resp.text.strip().split("\n")
            if len(lines) >= 2:
                p = lines[1].split("|")
                id_d = p[0]
                gmina = p[1]
                nr_d = p[2]
                wkt_str = p[3] if len(p) > 3 else None
                geom = wkt_loads(wkt_str) if wkt_str else None
                return {'id_dzialki': id_d, 'gmina': gmina, 'nr_dzialki': nr_d, 'geom': geom, 'x': x_2177, 'y': y_2177}
    except Exception: pass
    return None

def szacuj_cene_m2_avm(odleglosc_dom_km):
    return max(180.0, 600.0 - (odleglosc_dom_km * 8.0))

# ==============================================================================
# 4. INTERFEJS UŻYTKOWNIKA
# ==============================================================================
st.sidebar.title("🚙⚡🔥 PrzeSmyk v2.5")
st.sidebar.caption("Centrum Dowodzenia Terenowego")
st.sidebar.markdown("---")

st.sidebar.subheader("📍 Start Marszruty")
wpis_lokalizacji = st.sidebar.text_input("Wpisz z palca miasto, ulicę lub numer domu:", value="Kraków, ul. Nad Sudołem 32")
current_lat, current_lon, opis_lokalizacji = geokoduj_wpis_startowy(wpis_lokalizacji)
st.sidebar.caption(f"🎯 Zlokalizowano: **{opis_lokalizacji}**")

przelicz_button = st.sidebar.button("🚙 PRZELICZ TRASĘ", type="primary")

# BRAK DUŻEGO NAGŁÓWKA W OKNIE GŁÓWNYM
tab1, tab2, tab3 = st.tabs(["🗺️ Trasa & Operat", "📝 CRM Terenowy", "🗂️ Baza Działek"])

if przelicz_button:
    pobierz_plik_jesli_brak(URL_SIECI, PLIK_SIECI)
    pobierz_plik_jesli_brak(URL_SLUPY, PLIK_SLUPY)
    
    # SPINNER - OBRACAJĄCA SIĘ KLEPSYDRA
    with st.spinner("⏳ Trwa analiza terenu w promieniu 5 km od punktu startu..."):
        dom_x, dom_y = transformer_4326_to_2177.transform(current_lon, current_lat)
        punkt_dom = Point(dom_x, dom_y)
        bufor_5km = punkt_dom.buffer(5000.0)  # Sztywny promień 5 km
        
        odwiedzone_ids = get_visited_ids()
        
        sieci = gpd.read_file(PLIK_SIECI).to_crs("EPSG:2177")
        slupy = gpd.read_file(PLIK_SLUPY).to_crs("EPSG:2177") if os.path.exists(PLIK_SLUPY) else gpd.GeoDataFrame()

        # Wycięcie przestrzenne do bufora 5 km
        sieci_5km = sieci[sieci.geometry.intersects(bufor_5km)].copy()
        
        if sieci_5km.empty:
            st.warning("Nie znaleziono żadnych linii przesyłowych w promieniu 5 km od wskazanego punktu.")
        else:
            sieci_5km['geometry'] = sieci_5km.geometry.intersection(bufor_5km)
            sieci_5km['dist'] = sieci_5km.geometry.distance(punkt_dom) / 1000.0
            sieci_5km = sieci_5km.sort_values(by='dist', ascending=True)
            
            TARGET_COUNT = 100  # Docelowa liczba działek
            wykryte_dzialki = {}
            
            for idx, linia in sieci_5km.iterrows():
                if len(wykryte_dzialki) >= TARGET_COUNT:
                    break
                    
                opis_nap = str(linia.get('napiecie', linia.get('rodzaj', '110 kV'))).upper()
                szerokosc_strefy = 15.0
                for k, v in STREFY_SLUZEBNOSCI.items():
                    if k in opis_nap.lower(): szerokosc_strefy = v; break
                
                dlugosc = linia.geometry.length
                for d in range(0, int(dlugosc), 100):
                    if len(wykryte_dzialki) >= TARGET_COUNT:
                        break
                        
                    pt = linia.geometry.interpolate(d)
                    dzialka = uldk_pobierz_dzialke_z_geometria(pt.x, pt.y)
                    if dzialka:
                        id_d = dzialka['id_dzialki']
                        if id_d in odwiedzone_ids or id_d in wykryte_dzialki: continue
                        
                        poly = dzialka.get('geom')
                        if poly and not poly.is_empty:
                            rdzen_dzialki = poly.buffer(-5.0)
                            if rdzen_dzialki.is_empty or not rdzen_dzialki.intersects(linia.geometry):
                                continue
                        
                        lon_wgs, lat_wgs = transformer_2177_to_4326.transform(pt.x, pt.y)
                        adres_czysty, ok, typ_terenu = pobierz_adres_i_filtr_zabudowy(lat_wgs, lon_wgs, dzialka['nr_dzialki'])
                        
                        if not ok: continue
                            
                        dzialka.update({
                            'szer_pasa': szerokosc_strefy, 'rodzaj': opis_nap,
                            'adres': adres_czysty, 'typ': typ_terenu,
                            'lat': lat_wgs, 'lon': lon_wgs, 'dist': linia['dist']
                        })
                        wykryte_dzialki[id_d] = dzialka
            
            lista_rankingowa = []
            for id_d, d in wykryte_dzialki.items():
                ilosc_slupow = len(slupy[slupy.geometry.intersects(Point(d['x'], d['y']).buffer(30.0))]) if not slupy.empty else 0
                cena = szacuj_cene_m2_avm(d['dist'])
                pow_pasa = 85.0 * d['szer_pasa']
                roszczenie = pow_pasa * cena * WSPOLCZYNNIK_WSPOLKORZYSTANIA
                
                link_geo = f"https://mapy.geoportal.gov.pl/imap/Imgp_2.html?identifyParcel={id_d}"
                link_ema = f"https://polska.e-mapa.net/?dzialka={id_d}"
                link_ong = f"https://ongeo.pl/mapa?x={d['lon']:.6f}&y={d['lat']:.6f}&zoom=19"
                link_gmaps = f"https://www.google.com/maps?q={d['lat']},{d['lon']}"
                
                rodzaj_mediow = "🔥 Gazociąg" if "GAZ" in d['rodzaj'] else "⚡ Linia Elektroenergetyczna"
                
                lista_rankingowa.append({
                    'ID': id_d, 'Adres': d['adres'], 'Gmina': d['gmina'], 'Nr': d['nr_dzialki'], 'Typ': d['typ'],
                    'Linia': f"{rodzaj_mediow} ({d['rodzaj']})", 'Pow': pow_pasa, 'Cena': cena, 'Roszczenie': round(roszczenie, 2),
                    'Slupy': ilosc_slupow, 'Dist': d['dist'],
                    'LinkG': link_geo, 'LinkE': link_ema, 'LinkO': link_ong, 'LinkM': link_gmaps
                })
                
            df = pd.DataFrame(lista_rankingowa)
            if not df.empty:
                df = df.sort_values(by=['Roszczenie', 'Dist'], ascending=[False, True]).reset_index(drop=True)
                st.session_state['rank'] = df
                df.to_excel(PLIK_WYNIKOWY, index=False)
                
                if len(df) < 100:
                    st.warning(f"⚠️ W promieniu 5 km od miejsca startu znaleziono {len(df)} działek spełniających wszystkie kryteria (mniej niż docelowe 100).")
                else:
                    st.success(f"✅ Wygenerowano pełną marszrutę obejmującą TOP 100 działek w promieniu 5 km!")
            else:
                st.warning("Nie znaleziono działek spełniających kryteria w promieniu 5 km od wskazanego punktu.")

with tab1:
    if 'rank' in st.session_state:
        df_rank = st.session_state['rank']
        st.subheader(f"📍 Marszruta Terenowa: {len(df_rank)} działek")
        
        for idx, row in df_rank.iterrows():
            st.markdown(f"### {idx+1}. {row['Adres']}")
            c1, c2, c3 = st.columns([3, 3, 3])
            c1.markdown(f"📍 **Gmina:** {row['Gmina']}\n🆔 **ID Działki:** `{row['ID']}`\n{row['Typ']}")
            c2.markdown(f"💰 **Roszczenie:** `{row['Roszczenie']:,.2f} PLN`\n🗼 **Wieża / Słup:** `{row['Slupy']} szt.`\n📐 **Pas:** `{row['Pow']} m²`")
            c3.markdown(f"🌐 **Weryfikacja Geodezyjna:**\n• [Otwórz Działkę - Geoportal]({row['LinkG']})\n• [Otwórz Działkę - e-Mapa]({row['LinkE']})\n• [Otwórz Mapę - OnGeo.pl]({row['LinkO']})")
            
            b1, b2, b3 = st.columns([3, 3, 3])
            with b1.popover("📄 Pobierz Raport"):
                st.write(f"ID: {row['ID']}\nAdres: {row['Adres']}\nKwalifikacja: {row['Typ']}")
                st.download_button("💾 Pobierz Plik", f"ID: {row['ID']}\nAdres: {row['Adres']}", file_name=f"Raport_{row['Nr']}.txt")
            with b2.popover("📊 Kalkulator Roszczeń"):
                st.write(f"**Infrastruktura:** {row['Linia']}\n**Wzór:** `{row['Pow']} m² × {row['Cena']} PLN × 0.5`\n### Wartość: {row['Roszczenie']:,.2f} PLN")
            b3.link_button("🚗 Nawiguj (Google Maps)", row['LinkM'], type="primary")
            st.divider()

with tab2:
    with st.form("crm"):
        tid = st.text_input("ID Działki")
        stat = st.selectbox("Status Wizyty", ["Odwiedzona", "Umówione spotkanie", "Odmowa"])
        kw = st.text_input("Numer KW")
        wlas = st.text_input("Dane Kontaktowe Właściciela")
        notat = st.text_area("Notatka z rozmowy")
        if st.form_submit_button("💾 Zapisz w CRM"):
            save_crm_record(tid, stat, kw, wlas, notat)
            st.success("Zapisano rekord w bazie PrzeSmyk!")

with tab3:
    st.dataframe(get_all_crm_records(), use_container_width=True)
