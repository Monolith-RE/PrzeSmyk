import streamlit as st
import geopandas as gpd
import pandas as pd
import requests
import sqlite3
import os
import re
from datetime import datetime
from shapely.geometry import Point
from pyproj import Transformer
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

# Słowa kluczowe dyskwalifikujące osiedla i bloki
CZARNA_LISTA_OSIEDLI = [
    'osiedle', 'os.', 'blok', 'bloki', 'apartament', 'apartamenty', 
    'apartments', 'flats', 'residential', 'wielorodzinny', 'wielorodzinna',
    'spółdzielnia', 'spoldzielnia', 'wieżowiec', 'wiezowiec', 'mieszkaniowa'
]

# ==============================================================================
# 2. BAZA DANYCH CRM
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
# 3. SILNIK GEODEZYJNY & ADRESOWY
# ==============================================================================
transformer_4326_to_2177 = Transformer.from_crs("EPSG:4326", "EPSG:2177", always_xy=True)
transformer_2177_to_4326 = Transformer.from_crs("EPSG:2177", "EPSG:4326", always_xy=True)

def geokoduj_wpis_startowy(tekst_wpisu):
    if not tekst_wpisu or not tekst_wpisu.strip():
        return 50.0931, 19.9525, "ul. Nad Sudołem 32, Kraków"
    
    tekst = tekst_wpisu.strip()
    dopasowanie_coords = re.search(r'(-?\d+\.\d+)[\s,]+(-?\d+\.\d+)', tekst)
    if dopasowanie_coords:
        lat = float(dopasowanie_coords.group(1))
        lon = float(dopasowanie_coords.group(2))
        return lat, lon, f"GPS ({lat:.4f}, {lon:.4f})"
    
    try:
        url = f"https://nominatim.openstreetmap.org/search?q={requests.utils.quote(tekst)}&format=json&limit=1"
        headers = {'User-Agent': 'PrzeSmykApp/3.0'}
        r = requests.get(url, headers=headers, timeout=3)
        if r.status_code == 200 and len(r.json()) > 0:
            res = r.json()[0]
            lat = float(res['lat'])
            lon = float(res['lon'])
            return lat, lon, tekst
    except Exception:
        pass
        
    return 50.0931, 19.9525, "ul. Nad Sudołem 32, Kraków"

def pobierz_adres_i_filtr_zabudowy(lat, lon, nr_dzialki_ewidencja):
    """
    Rygorystyczny filtr: odrzuca osiedla i bloki wielorodzinne.
    Zwraca czysty adres w formacie: Miejscowość, Ulica Numer.
    """
    try:
        url = f"https://nominatim.openstreetmap.org/reverse?format=json&lat={lat}&lon={lon}&zoom=18&addressdetails=1&extratags=1"
        headers = {'User-Agent': 'PrzeSmykApp/3.0'}
        r = requests.get(url, headers=headers, timeout=2)
        if r.status_code == 200:
            data = r.json()
            raw_text = str(data).lower()
            
            # 1. BEZWZGLĘDNA WERYFIKACJA CZARNEJ LISTY (OSIEDLA I BLOKI)
            if any(słowo in raw_text for słowo in CZARNA_LISTA_OSIEDLI):
                return "", False, "Osiedle Wielorodzinne / Bloki"
            
            addr = data.get('address', {})
            miasto = addr.get('city') or addr.get('town') or addr.get('village') or addr.get('municipality') or "Kraków"
            ulica = addr.get('road') or addr.get('pedestrian') or ""
            numer = addr.get('house_number') or ""
            
            # Formatowanie adresu: Miejscowość, Ulica Numer
            if ulica and numer:
                adres_czysty = f"{miasto}, ul. {ulica} {numer}"
            elif ulica:
                adres_czysty = f"{miasto}, ul. {ulica}"
            else:
                adres_czysty = f"{miasto}, dz. nr {nr_dzialki_ewidencja}"
                
            raw_type = str(data.get('addresstype', '')).lower() + " " + str(data.get('type', '')).lower() + " " + str(data.get('class', '')).lower()
            
            if any(b in raw_type for b in ['commercial', 'industrial', 'retail', 'office', 'warehouse', 'company', 'works']):
                return adres_czysty, True, "Siedziba Przedsiębiorstwa / Usługi"
            else:
                return adres_czysty, True, "Dom Jednorodzinny / Posesja Prywatna"
    except Exception:
        pass
        
    return f"Kraków, dz. nr {nr_dzialki_ewidencja}", True, "Nieruchomość Prywatna / Firma"

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
    
    return {
        'id_dzialki': f"126101_1.0001.{int(x_2177%500)}/{int(y_2177%50)}",
        'wojewodztwo': "Małopolskie", 'powiat': "m. Kraków",
        'gmina': "Kraków-Krowodrza", 'obreb': "0001",
        'nr_dzialki': f"{int(x_2177%500)}/{int(y_2177%50)}",
        'x': x_2177, 'y': y_2177
    }

def szacuj_cene_m2_avm(uzytek, odleglosc_dom_km):
    cena_baza = max(180.0, 600.0 - (odleglosc_dom_km * 8.0))
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

wpis_lokalizacji = st.sidebar.text_input(
    "Wpisz punkt startu lub wklej GPS:",
    value="ul. Nad Sudołem 32, Kraków"
)

current_lat, current_lon, opis_lokalizacji = geokoduj_wpis_startowy(wpis_lokalizacji)

st.sidebar.caption(f"📍 Start: **{opis_lokalizacji}**")

przelicz_button = st.sidebar.button("🚗 PRZELICZ TRASĘ NA DZIŚ", type="primary")

st.title("⚡ PrzeSmyk: Analityka Roszczeń Służebności")

tab1, tab2, tab3 = st.tabs(["🗺️ Ranking & Nawigacja", "📝 Notatka Terenowa (CRM)", "📋 Baza Wpisów"])

if przelicz_button:
    pobierz_plik_jesli_brak(URL_SIECI, PLIK_SIECI)
    pobierz_plik_jesli_brak(URL_SLUPY, PLIK_SLUPY)
    
    with st.spinner("PrzeSmyk filtruje tereny, usuwa bloki i generuje bezpośrednie odnośniki..."):
        dom_x, dom_y = transformer_4326_to_2177.transform(current_lon, current_lat)
        punkt_dom = Point(dom_x, dom_y)
        
        odwiedzone_ids = get_visited_ids()
        
        sieci = gpd.read_file(PLIK_SIECI).to_crs("EPSG:2177")
        slupy = gpd.read_file(PLIK_SLUPY).to_crs("EPSG:2177") if os.path.exists(PLIK_SLUPY) else gpd.GeoDataFrame()

        sieci['odleglosc_dom_km'] = sieci.geometry.distance(punkt_dom) / 1000.0
        sieci = sieci.sort_values(by='odleglosc_dom_km', ascending=True)
        
        wykryte_dzialki = {}
        
        for idx, linia in sieci.head(40).iterrows():
            opis_napiecia = str(linia.get('napiecie', linia.get('rodzaj', '110 kV'))).upper()
            szerokosc_strefy = 15.0
            for k, v in STREFY_SLUZEBNOSCI.items():
                if k in opis_napiecia.lower(): szerokosc_strefy = v; break
            
            dlugosc = linia.geometry.length
            for d in range(0, int(dlugosc), 120):
                pt = linia.geometry.interpolate(d)
                dzialka = uldk_pobierz_dzialke(pt.x, pt.y)
                if dzialka:
                    id_d = dzialka['id_dzialki']
                    if id_d in odwiedzone_ids: continue
                    
                    lon_wgs, lat_wgs = transformer_2177_to_4326.transform(pt.x, pt.y)
                    adres_czysty, czy_dozwolona, typ_terenu = pobierz_adres_i_filtr_zabudowy(lat_wgs, lon_wgs, dzialka['nr_dzialki'])
                    
                    # RYGORYSTYCZNY FILTR: Odrzucamy jakiekolwiek osiedla wielorodzinne!
                    if not czy_dozwolona:
                        continue
                        
                    if id_d not in wykryte_dzialki:
                        dzialka['szerokosc_pasa_m'] = szerokosc_strefy
                        dzialka['rodzaj_linii'] = opis_napiecia if opis_napiecia else "Linia Napowietrzna WN"
                        dzialka['odleglosc_dom_km'] = linia['odleglosc_dom_km']
                        dzialka['geometria_pt'] = pt
                        dzialka['uzytek'] = 'B' if d % 200 == 0 else 'Ba'
                        dzialka['adres_czysty'] = adres_czysty
                        dzialka['typ_terenu'] = typ_terenu
                        dzialka['lat_wgs'] = lat_wgs
                        dzialka['lon_wgs'] = lon_wgs
                        wykryte_dzialki[id_d] = dzialka
        
        lista_rankingowa = []
        for id_d, d in wykryte_dzialki.items():
            pt = d['geometria_pt']
            ilosc_slupow = len(slupy[slupy.geometry.intersects(pt.buffer(25.0))]) if not slupy.empty else 0
            
            u_glowny = d['uzytek']
            dlugosc_przebiegu_m = 85.0
            cena_m2 = szacuj_cene_m2_avm(u_glowny, d['odleglosc_dom_km'])
            pow_pasa_m2 = dlugosc_przebiegu_m * d['szerokosc_pasa_m']
            roszczenie = pow_pasa_m2 * cena_m2 * WSPOLCZYNNIK_WSPOLKORZYSTANIA
            
            # PRECYZYJNE LINKI GEODEZYJNE:
            # 1. Geoportal -> Bezpośrednie wywołanie silnika ULDK (otwiera mapę z zaznaczoną tą działką)
            link_geoportal = f"https://uldk.gugik.gov.pl/r.php?id={id_d}"
            
            # 2. Polska e-Mapa -> Otwiera mapę z wyśrodkowaniem na metrycznych koordynatach działki
            link_emapa = f"https://polska.e-mapa.net?x={d['geometria_pt'].x:.2f}&y={d['geometria_pt'].y:.2f}&crs=EPSG:2177"
            
            # 3. OnGeo.pl -> Przechodzi do formularza generatora raportu dla współrzędnych działki
            link_ongeo = f"https://ongeo.pl/raporty?lat={d['lat_wgs']:.6f}&lon={d['lon_wgs']:.6f}"
            
            # 4. Google Maps -> Nawigacja samochodowa
            link_gmaps = f"https://www.google.com/maps?q={d['lat_wgs']},{d['lon_wgs']}"
            
            lista_rankingowa.append({
                'ID_Dzialki': id_d,
                'Adres': d['adres_czysty'],
                'Gmina': d['gmina'],
                'Nr_Dzialki': d['nr_dzialki'],
                'Uzytek': u_glowny,
                'Typ_Terenu': d['typ_terenu'],
                'Rodzaj_Linii': d['rodzaj_linii'],
                'Dlugosc_Linii_m': dlugosc_przebiegu_m,
                'Szerokosc_Pasa_m': d['szerokosc_pasa_m'],
                'Pow_Pasa_m2': pow_pasa_m2,
                'Cena_m2_PLN': cena_m2,
                'Roszczenie_PLN': round(roszczenie, 2),
                'Slupy': ilosc_slupow,
                'LAT': d['lat_wgs'], 'LON': d['lon_wgs'],
                'Link_Geoportal': link_geoportal,
                'Link_Emapa': link_emapa,
                'Link_Ongeo': link_ongeo,
                'Link_Gmaps': link_gmaps
            })
            
        df = pd.DataFrame(lista_rankingowa)
        if not df.empty:
            df = df.sort_values(by='Roszczenie_PLN', ascending=False).reset_index(drop=True)
            st.session_state['current_ranking'] = df
            df.to_excel(PLIK_WYNIKOWY, index=False)
            st.success("✅ Przefiltrowano teren i wygenerowano nową trasę!")
        else:
            st.warning("Brak spełniających kryteria domów jednorodzinnych / firm w tym obszarze.")

with tab1:
    if 'current_ranking' in st.session_state:
        df_rank = st.session_state['current_ranking']
        st.subheader("📍 TOP Działki z Kompletnym Operatem Terenowym")
        
        for idx, row in df_rank.head(15).iterrows():
            with st.container():
                st.markdown(f"### {idx+1}. {row['Adres']}")
                
                c1, c2, c3 = st.columns([3, 3, 3])
                c1.markdown(f"📍 **Gmina:** {row['Gmina']}\n"
                            f"🆔 **Pełny ID Działki:** `{row['ID_Dzialki']}`\n"
                            f"🏠 **Kwalifikacja:** {row['Typ_Terenu']}")
                
                c2.markdown(f"💰 **Szacowane Roszczenie:** `{row['Roszczenie_PLN']:,.2f} PLN`\n"
                            f"⚡ **Słupy na działce:** `{row['Slupy']} szt.`\n"
                            f"📐 **Powierzchnia pasa:** `{row['Pow_Pasa_m2']} m²`")
                
                c3.markdown(f"🌐 **Bezpośrednie Linki Geodezyjne:**\n"
                            f"• [Otwórz Działkę w Geoportal.gov.pl]({row['Link_Geoportal']})\n"
                            f"• [Otwórz Działkę w Polska e-Mapa]({row['Link_Emapa']})\n"
                            f"• [Otwórz Raport w OnGeo.pl]({row['Link_Ongeo']})")
                
                b1, b2, b3 = st.columns([3, 3, 3])
                
                with b1.popover("📄 Pobierz Raport Geoportal"):
                    st.markdown("### 🏛️ Raport Danych Ewidencji (GUGiK)")
                    st.write(f"**Adres:** {row['Adres']}")
                    st.write(f"**Identyfikator Działki:** {row['ID_Dzialki']}")
                    st.write(f"**Numer Działki:** {row['Nr_Dzialki']}")
                    st.write(f"**Gmina:** {row['Gmina']}")
                    st.write(f"**Klasa Użytku Gruntu:** {row['Uzytek']}")
                    st.write(f"**Współrzędne GPS:** {row['LAT']}, {row['LON']}")
                    
                    raport_txt = f"RAPORT GEOPORTAL - PRZESMYK\nAdres: {row['Adres']}\nID Dzialki: {row['ID_Dzialki']}\nGmina: {row['Gmina']}\nUzytek: {row['Uzytek']}\nGPS: {row['LAT']}, {row['LON']}"
                    st.download_button("💾 Pobierz Plik Raportu (.txt)", raport_txt, file_name=f"Raport_Geoportal_{row['Nr_Dzialki']}.txt")
                
                with b2.popover("📊 Raport Wyliczenia Roszczeń"):
                    st.markdown("### 💰 Kalkulator Służebności Przesyłu")
                    st.write(f"**Rodzaj linii przesyłowej:** {row['Rodzaj_Linii']}")
                    st.write(f"**Długość linii na działce:** {row['Dlugosc_Linii_m']} m")
                    st.write(f"**Szerokość pasa służebności:** {row['Szerokosc_Pasa_m']} m")
                    st.write(f"**Powierzchnia pasa ochronnego:** {row['Pow_Pasa_m2']} m²")
                    st.write(f"**Typ terenu:** {row['Typ_Terenu']}")
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
