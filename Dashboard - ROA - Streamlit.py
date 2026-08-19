import streamlit as st
import pandas as pd
import os
import json
import datetime
import time
import uuid
import openpyxl
from openpyxl.styles import PatternFill, Font

# --- PAGINA CONFIGURATIE ---
st.set_page_config(page_title="ROA Dashboard", layout="wide", initial_sidebar_state="expanded")

# --- BEVEILIGING CONFIGURATIE ---
GEBRUIKERS_LIJST = {
    'A': 's',
    'eqh_user': 'Welkom01'
}

# --- CONFIGURATIE & BASISINSTELLINGEN ---
STATUS_FILE = "live_voorraad_status.json"
ACTIEF_LOG_BESTAND = "trolley_logging.xlsx"

START_VOORRAAD_BASIS = {'UTR': 52, 'T11': 26, 'T12': 52, '2T12': 52, 'TW': 52, 'TWA': 52}
CUSTOM_START_VOORRAAD = {'UTR': 52, 'T11': 26, 'T12': 52, '2T12': 52, 'TW': 52, 'TWA': 52}

alle_excel_kolommen = ['UTR', 'T11', 'T12', '2T12', 'TW', 'TWA']
balk_kolommen = ['UTR', 'T11', 'T12', '2T12', 'TW', 'TWA']

AFBEELDING_CONFIG = {
    'UTR': 'assets/UTR.jpg',
    'T12': 'assets/T12.jpg',
    'TW': 'assets/TW.jpg',
    'T11': 'assets/T11.jpg',
    '2T12': 'assets/2T12.jpg',
    'TWA': 'assets/TWA.jpg'
}

BLOK_TIJDEN = {
    '1': datetime.time(6, 0, 0),
    '2': datetime.time(7, 25, 0),
    '3': datetime.time(8, 50, 0),
    '4': datetime.time(10, 15, 0),
    '5': datetime.time(11, 40, 0),
    '6': datetime.time(13, 5, 0),
    '7': datetime.time(14, 30, 0),
    '8': datetime.time(15, 55, 0),
    '9': datetime.time(17, 20, 0),
    '10': datetime.time(18, 45, 0),
    '11': datetime.time(20, 10, 0),
    '12': datetime.time(21, 35, 0)
}

# --- HULPFUNCTIES ---
def probeer_getal(waarde):
    try:
        return float(waarde) if '.' in str(waarde) else int(waarde)
    except ValueError:
        return waarde

def bepaal_voorraad_status_en_kleur(actuele_waarde, basis_waarde):
    if basis_waarde == 0:
        percentage = 100
    else:
        percentage = (actuele_waarde / basis_waarde) * 100
        
    if percentage >= 60:
        return "Groen", "Normaal", "#27ae60"
    elif percentage >= 40:
        return "Geel", "Let op: Lage voorraad", "#f1c40f"
    else:
        return "Rood", "WAARSCHUWING: Onvoldoende voorraad!", "#e74c3c"

def is_blok_vrijgegeven(df, blok_naam, handmatige_vrijgaven):
    blok_str = str(blok_naam).replace("🕒", "").replace("Block", "").strip()
    if handmatige_vrijgaven and blok_str in handmatige_vrijgaven:
        return True
    if blok_str == '1':
        return True
    nu_tijd = datetime.datetime.now().time()
    if blok_str in BLOK_TIJDEN and nu_tijd >= BLOK_TIJDEN[blok_str]:
        return True
    try:
        blok_nr = int(blok_str)
        vorig_blok_str = str(blok_nr - 1)
        if 'Block' in df.columns and 'Total' in df.columns:
            totaal_vorig = df[df['Block'].astype(str).str.replace("🕒", "").str.replace("Block", "").str.strip() == vorig_blok_str]['Total'].sum()
            if totaal_vorig == 0:
                return True
    except ValueError:
        pass
    return False

def log_klik_naar_excel(trolley_type, block, vluchtnummer, tijd_nu, actuele_voorraad=None, basis_voorraad=None, status="In gebruik", huidige_plus_stand=0):
    nieuwe_datum = tijd_nu.strftime('%Y-%m-%d')
    nieuwe_tijd = tijd_nu.strftime('%H:%M:%S')
    equipment = 'EUR' if trolley_type in ['MUTR-CCL', 'UTR', 'T12', 'TW', 'OIS', 'A-OIS'] else ('KLC' if trolley_type in ['T11', '2T12', 'TWA'] else 'ONBEKEND')

    _, waarschuwing, _ = bepaal_voorraad_status_en_kleur(actuele_voorraad or 0, basis_voorraad or 1)
    log_doel_voorraad = (actuele_voorraad + huidige_plus_stand) if actuele_voorraad is not None else 'N.v.t.'

    nieuw_record = {
        'Datum': [nieuwe_datum], 'Tijd': [nieuwe_tijd], 'Trolley Type': [trolley_type],
        'Equipment': [equipment], 'Block': [block], 'Vluchtnummer': [vluchtnummer],
        'Status': [status], 'Voorraad Niveau': [actuele_voorraad if actuele_voorraad is not None else 'N.v.t.'],
        'Verwachte Voorraad': [log_doel_voorraad], 'Waarschuwing': [waarschuwing]
    }
    df_nieuw = pd.DataFrame(nieuw_record)
    kolommen = ['Datum', 'Tijd', 'Trolley Type', 'Equipment', 'Block', 'Vluchtnummer', 'Status', 'Voorraad Niveau', 'Verwachte Voorraad', 'Waarschuwing']
    
    try:
        if os.path.exists(ACTIEF_LOG_BESTAND):
            df_bestaand = pd.read_excel(ACTIEF_LOG_BESTAND)
            df_totaal = pd.concat([df_bestaand, df_nieuw], ignore_index=True)
        else:
            df_totaal = df_nieuw
        df_totaal[kolommen].to_excel(ACTIEF_LOG_BESTAND, index=False)
    except Exception as e:
        st.warning(f"Log schrijven mislukt: {e}")

# --- SESSIE STATUS INITIALISATIE ---
if 'logged_in' not in st.session_state:
    st.session_state.logged_in = False

if 'df_data' not in st.session_state:
    st.session_state.df_data = pd.DataFrame(columns=['Block', 'Vluchttijd', 'Vluchtnummer'] + alle_excel_kolommen + ['Total'])

if 'handmatige_plus' not in st.session_state:
    st.session_state.handmatige_plus = {col: 0 for col in balk_kolommen}

if 'handmatige_min' not in st.session_state:
    st.session_state.handmatige_min = {col: max(0, START_VOORRAAD_BASIS[col] - int(CUSTOM_START_VOORRAAD.get(col, START_VOORRAAD_BASIS[col]))) for col in balk_kolommen}

if 'actieve_timers' not in st.session_state:
    st.session_state.actieve_timers = []

if 'actie_historie' not in st.session_state:
    st.session_state.actie_historie = []

if 'handmatige_vrijgaven' not in st.session_state:
    st.session_state.handmatige_vrijgaven = []

# --- INLOGSCHERM ---
if not st.session_state.logged_in:
    st.markdown("<h2 style='text-align: center;'>ROA Dashboard Login</h2>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        with st.form("login_form"):
            user = st.text_input("Gebruikersnaam")
            pw = st.text_input("Wachtwoord", type="password")
            btn = st.form_submit_button("Inloggen", use_container_width=True)
            if btn:
                if user.strip() in GEBRUIKERS_LIJST and GEBRUIKERS_LIJST[user.strip()] == pw.strip():
                    st.session_state.logged_in = True
                    st.rerun()
                else:
                    st.error("Onjuiste gebruikersnaam of wachtwoord!")
    st.stop()

# --- SIDEBAR (BESTAND UPLOAD & FILTERS) ---
st.sidebar.title("Instellingen & Upload")

uploaded_file = st.sidebar.file_uploader("Upload Excel Schema (Map1)", type=['xlsx', 'xls'])
if uploaded_file is not None and st.sidebar.button("Laad nieuw bestand in"):
    df_temp = pd.read_excel(uploaded_file)
    for col in alle_excel_kolommen + ['Total']:
        if col in df_temp.columns:
            df_temp[col] = df_temp[col].fillna(0).astype(int)
    if 'Block' in df_temp.columns:
        df_temp['Block'] = df_temp['Block'].astype(str).str.replace("🕒", "").str.replace("Block", "").str.strip()
    st.session_state.df_data = df_temp
    st.sidebar.success("Excel succesvol ingeladen!")

filter_dienst = st.sidebar.selectbox("Kies Filter / Dienst", ["Alle", "Ochtend (Block 1-6)", "Middag (Block 7-12)"])
view_mode = st.sidebar.radio("Visualisatie Mode", ["Vlucht Details", "Totaal Per Blok"])

if st.sidebar.button("↩️ Laatste Actie Ongedaan Maken", use_container_width=True):
    if st.session_state.actie_historie:
        laatste = st.session_state.actie_historie.pop()
        t_col, t_block, t_flight = laatste['trolley_type'], laatste['block'], laatste['vluchtnummer']
        df = st.session_state.df_data
        match = df[(df['Block'] == t_block) & (df['Vluchtnummer'].astype(str).str.strip() == str(t_flight))]
        if not match.empty:
            idx = match.index[0]
            df.at[idx, t_col] = int(df.at[idx, t_col]) + 1
            df.at[idx, 'Total'] = df.loc[idx, alle_excel_kolommen].astype(int).sum()
            st.session_state.handmatige_plus[t_col] = max(0, st.session_state.handmatige_plus[t_col] - 1)
            st.success(f"↩️ {t_col} hersteld op vlucht {t_flight}")
            st.rerun()

# --- HOOFDSCHERM LAYOUT ---
st.title("Beschikbare Vlucht Informatie - ROA")
st.caption(f"Live Tijd: {datetime.datetime.now().strftime('%d-%m-%Y | %H:%M:%S')}")

# --- ACTUELE VOORRAADBALKEN ---
st.subheader("Actuele voorraad in EQH-afdeling")

nu = time.time()
st.session_state.actieve_timers = [t for t in st.session_state.actieve_timers if nu < t['vrijgave_tijd']]

def get_stats(col_id):
    basis = START_VOORRAAD_BASIS[col_id]
    t_plus = st.session_state.handmatige_plus[col_id]
    in_wacht = sum(1 for t in st.session_state.actieve_timers if t['type'] == col_id)
    afgeboekt = st.session_state.handmatige_min[col_id]
    doel = max(0, basis + t_plus - afgeboekt)
    actueel = max(0, doel - in_wacht)
    return actueel, doel, basis

col_eur, col_klc = st.columns(2)

with col_eur:
    st.markdown("**EUR-equipment**")
    c1, c2, c3 = st.columns(3)
    for idx, item in enumerate(['T12', 'TW', 'UTR']):
        act, doel, bas = get_stats(item)
        _, _, hex_kleur = bepaal_voorraad_status_en_kleur(act, bas)
        with [c1, c2, c3][idx]:
            st.markdown(f"<div style='background-color:{hex_kleur}; padding:10px; border-radius:8px; color:white; text-align:center;'><b>{item}</b><br><h3>{act} / {doel}</h3></div>", unsafe_allow_html=True)

with col_klc:
    st.markdown("**KLC-equipment**")
    c1, c2, c3 = st.columns(3)
    for idx, item in enumerate(['2T12', 'TWA', 'T11']):
        act, doel, bas = get_stats(item)
        _, _, hex_kleur = bepaal_voorraad_status_en_kleur(act, bas)
        with [c1, c2, c3][idx]:
            st.markdown(f"<div style='background-color:{hex_kleur}; padding:10px; border-radius:8px; color:white; text-align:center;'><b>{item}</b><br><h3>{act} / {doel}</h3></div>", unsafe_allow_html=True)

st.divider()

# --- DATATABEL EN AFBOEKEN ---
df_display = st.session_state.df_data.copy()

if not df_display.empty:
    if filter_dienst == "Ochtend (Block 1-6)":
        df_display = df_display[df_display['Block'].isin(['1','2','3','4','5','6'])]
    elif filter_dienst == "Middag (Block 7-12)":
        df_display = df_display[df_display['Block'].isin(['7','8','9','10','11','12'])]

    st.subheader("Vluchten Overzicht")
    
    # Knoppen om eenvoudig trolleys af te boeken per vlucht (iPad vriendelijk)
    for idx, row in df_display.iterrows():
        block_nr = str(row['Block']).strip()
        is_open = is_blok_vrijgegeven(st.session_state.df_data, block_nr, st.session_state.handmatige_vrijgaven)
        
        with st.expander(f"Block {block_nr} | Vlucht: {row.get('Vluchtnummer', 'N.v.t.')} | Tijd: {row.get('Vluchttijd', '')} {'🔓' if is_open else '🔒'}"):
            if not is_open:
                st.warning("Dit blok is vergrendeld.")
                if st.button(f"Handmatig ontgrendelen Block {block_nr}", key=f"unlock_{idx}"):
                    st.session_state.handmatige_vrijgaven.append(block_nr)
                    st.success(f"Block {block_nr} vrijgegeven!")
                    st.rerun()
            else:
                st.write("Klik op een trolley-type om er 1 af te boeken:")
                btn_cols = st.columns(len(alle_excel_kolommen))
                for c_idx, t_type in enumerate(alle_excel_kolommen):
                    aantal = int(row.get(t_type, 0))
                    with btn_cols[c_idx]:
                        if st.button(f"{t_type} ({aantal})", key=f"btn_{idx}_{t_type}", disabled=(aantal <= 0)):
                            st.session_state.df_data.at[idx, t_type] = aantal - 1
                            st.session_state.df_data.at[idx, 'Total'] = st.session_state.df_data.loc[idx, alle_excel_kolommen].astype(int).sum()
                            st.session_state.handmatige_plus[t_type] += 1
                            st.session_state.actieve_timers.append({'type': t_type, 'vrijgave_tijd': time.time() + 900.0})
                            
                            st.session_state.actie_historie.append({
                                'trolley_type': t_type, 'block': block_nr,
                                'vluchtnummer': row.get('Vluchtnummer', ''), 'timestamp': time.time()
                            })
                            
                            act, _, bas = get_stats(t_type)
                            log_klik_naar_excel(t_type, block_nr, row.get('Vluchtnummer', ''), datetime.datetime.now(), act, bas)
                            st.success(f"1x {t_type} afgeboekt!")
                            st.rerun()

    st.divider()
    st.dataframe(df_display, use_container_width=True)

else:
    st.info("Upload een Excel-bestand via het zijmenu om te beginnen.")