import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import numpy as np
import re
import json
import io

# Konfigurasi Halaman Full Wide (Tanpa Sidebar Kiri)
st.set_page_config(
    page_title="Komparasi PL Yamaha Jember",
    page_icon="🏍️",
    layout="wide"
)

DEFAULT_FILEPATH = "MASTER DISKON SEPTEMBER 2026.xls"

def format_rp_kosong(val):
    if pd.isna(val) or val is None or val <= 0:
        return ""
    try:
        return f"Rp {int(round(float(val))):,}".replace(",", ".")
    except:
        return ""

def format_rp_angka(val):
    if pd.isna(val) or val is None:
        return "Rp 0"
    try:
        return f"Rp {int(round(float(val))):,}".replace(",", ".")
    except:
        return "Rp 0"

def normalize_motor_name(s):
    s = str(s).strip().upper().replace('"', '')
    s = s.replace('/', ' / ')
    s = re.sub(r'\s+', ' ', s).strip()
    s = s.replace('N - MAX', 'N-MAX').replace('N MAX', 'N-MAX')
    alias_map = {
        'FINO 125 PREMIUM / SPORTI': 'NEW FINO PREMIUM / SPORTY 125 BLUE CORE',
        'NEW FINO PREMIUM 125 BLUE CORE': 'NEW FINO PREMIUM / SPORTY 125 BLUE CORE',
        'NEW FINO SPORTY 125 BLUE CORE': 'NEW FINO PREMIUM / SPORTY 125 BLUE CORE',
        'FAZZIO HYBRID LUX': 'FAZZIO HYBRID LUX / FAZZIO LUX',
        'FAZZIO LUX': 'FAZZIO HYBRID LUX / FAZZIO LUX',
        'LEXI LX 155 CONNECT ABS MAX': 'LEXI LX 155 CONNECTION ABS MAX',
        'MX KING 150 CYBERCITY': 'MX KING / MX KING 150 CYBERCITY',
        'MX KING': 'MX KING / MX KING 150 CYBERCITY',
        'WR 155': 'WR 155 R',
        'X - MAX TECH MAX': 'X-MAX TECH MAX',
        'XMAX TECH MAX': 'X-MAX TECH MAX',
        'G - ULTIMA HYBIRD': 'G - ULTIMA HYBRID',
        'G - ULTIMA HYBIRD S': 'G - ULTIMA HYBRID S',
        'G - ULTIMA HYBIRD SOLID': 'G - ULTIMA HYBRID SOLID',
        'G - ULTIMA HYBIRD SMART': 'G - ULTIMA HYBRID SMART',
        'ALL NEW N-MAX 155': 'ALL NEW N-MAX',
        'ALL NEW N-MAX 155 S': 'ALL NEW N-MAX S',
        'ALL NEW N-MAX 155 CONNECTED': 'ALL NEW N-MAX CONNECTED',
        'ALL NEW N-MAX 155 ABS CONNECTED': 'ALL NEW N-MAX ABS C',
    }
    return alias_map.get(s, s)

def categorize_motor(name):
    n = name.upper()
    if 'NMAX' in n or 'N-MAX' in n:
        return 'NMAX SERIES'
    elif 'FILANO' in n or 'GRAND FILANO' in n:
        return 'FILLANO SERIES'
    elif 'FAZZIO' in n:
        return 'FAZZIO SERIES'
    elif 'GEAR' in n:
        return 'GEAR SERIES'
    else:
        return 'OTHERS'

@st.cache_data
def load_all_programs_from_bytes(file_bytes, file_name):
    if isinstance(file_bytes, bytes):
        excel_file = io.BytesIO(file_bytes)
    else:
        excel_file = file_bytes

    # 0. Sheet CASH
    df_cash_raw = pd.read_excel(excel_file, sheet_name='CASH')
    df_cash = df_cash_raw[pd.to_numeric(df_cash_raw.iloc[:, 1], errors='coerce').notna()].copy()
    cash_cols = ['Empty', 'NO', 'TYPE', 'OTR', 'BD', 'DISC_2023', 'DISC_2024', 'DISC_2025', 'DISC_2026', 'INCV_SALES', 'INCV_COUNTER', 'Dummy', 'Dummy2', 'PROFIT_2023', 'PROFIT_2024', 'PROFIT_2025', 'PROFIT_2026']
    df_cash = df_cash.iloc[:, :len(cash_cols)]
    df_cash.columns = cash_cols
    df_cash = df_cash.drop(columns=['Empty', 'Dummy', 'Dummy2']).reset_index(drop=True)
    df_cash['TYPE_NORM'] = df_cash['TYPE'].apply(normalize_motor_name)
    for col in ['OTR', 'BD', 'DISC_2023', 'DISC_2024', 'DISC_2025', 'DISC_2026', 'INCV_SALES', 'INCV_COUNTER', 'PROFIT_2025', 'PROFIT_2026']:
        if col in df_cash.columns:
            df_cash[col] = pd.to_numeric(df_cash[col], errors='coerce').fillna(0)

    # 1. Sheet REGULER
    df_reg_raw = pd.read_excel(excel_file, sheet_name='REGULER ')
    df_reg = df_reg_raw[
        df_reg_raw.iloc[:, 0].notna() & 
        (df_reg_raw.iloc[:, 0].astype(str).str.strip() != 'BACK') &
        pd.to_numeric(df_reg_raw.iloc[:, 1], errors='coerce').notna() &
        pd.to_numeric(df_reg_raw.iloc[:, 3], errors='coerce').notna()
    ].copy()
    
    reg_cols = [
        'TYPE', 'OTR', 'DP_PCT', 'DP_GROSS', 
        'ANGS_11', 'ANGS_17', 'ANGS_23', 'ANGS_29', 'ANGS_35', 'ANGS_47', 'Empty_10',
        'TAC_11', 'TAC_17', 'TAC_23', 'TAC_29', 'TAC_35', 'TAC_47', 'BD',
        'ICTV_KOOR', 'ICTV_SALES', 'ICTV_COUNTER', 'ICTV_SM', 'INCV_TOTAL', 'Empty_23',
        'DISC25_11', 'DISC25_17', 'DISC25_23', 'DISC25_29', 'DISC25_35', 'DISC25_47',
        'DISC26_11', 'DISC26_17', 'DISC26_23', 'DISC26_29', 'DISC26_35',
        'PROF25_11', 'PROF25_17', 'PROF25_23', 'PROF25_29', 'PROF25_35', 'Empty_41',
        'PROF26_11', 'PROF26_17', 'PROF26_23', 'PROF26_29', 'PROF26_35'
    ]
    df_reg = df_reg.iloc[:, :len(reg_cols)]
    df_reg.columns = reg_cols
    df_reg['TYPE_NORM'] = df_reg['TYPE'].apply(normalize_motor_name)
    for col in df_reg.columns[1:-1]:
        df_reg[col] = pd.to_numeric(df_reg[col], errors='coerce').fillna(0)
    df_reg['INCV_SALES'] = df_reg['ICTV_SALES']

    # 2. Sheet ALL LEASING 33 & ADIRA 33
    cols_33 = [
        'TYPE', 'OTR', 'DP_PCT', 'DP_GROSS', 'ANGS_33', 'TAC', 'BD',
        'DISC_2023', 'DISC_2025', 'DISC_2026',
        'INCV_KOOR', 'INCV_SALES', 'INCV_COUNTER', 'INCV_SM', 'INCV_TOTAL',
        'PROFIT_2023', 'PROFIT_2025', 'PROFIT_2026'
    ]
    def parse_sheet_33(sheet_name):
        if isinstance(file_bytes, bytes):
            ef = io.BytesIO(file_bytes)
        else:
            ef = file_bytes
        raw = pd.read_excel(ef, sheet_name=sheet_name)
        clean = raw[
            raw.iloc[:, 0].notna() & 
            (raw.iloc[:, 0].astype(str).str.strip() != 'BACK') &
            pd.to_numeric(raw.iloc[:, 1], errors='coerce').notna() &
            pd.to_numeric(raw.iloc[:, 3], errors='coerce').notna()
        ].copy()
        clean = clean.iloc[:, :len(cols_33)]
        clean.columns = cols_33
        clean['TYPE_NORM'] = clean['TYPE'].apply(normalize_motor_name)
        for c in cols_33[1:]:
            clean[c] = pd.to_numeric(clean[c], errors='coerce').fillna(0)
        return clean.reset_index(drop=True)

    df_all33 = parse_sheet_33('ALL LEASING  33')
    df_adira33 = parse_sheet_33('ADIRA 33 ')

    # 3. Sheet BAF TOP
    if isinstance(file_bytes, bytes):
        ef = io.BytesIO(file_bytes)
    else:
        ef = file_bytes
    df_baf_raw = pd.read_excel(ef, sheet_name='BAF TOP')
    df_baf = df_baf_raw[
        df_baf_raw.iloc[:, 1].notna() & 
        (~df_baf_raw.iloc[:, 1].astype(str).str.strip().isin(['BACK', 'TYPE'])) &
        pd.to_numeric(df_baf_raw.iloc[:, 3], errors='coerce').notna() &
        pd.to_numeric(df_baf_raw.iloc[:, 4], errors='coerce').notna()
    ].copy()
    baf_cols = [
        'Empty', 'TYPE', 'DP_PCT', 'OTR', 'DP_GROSS', 'ANGS_11', 'ANGS_17',
        'TAC_11', 'TAC_17', 'BD', 'DISC25_11', 'DISC25_17', 'DISC26_11', 'DISC26_17',
        'INCV_KOOR', 'INCV_SALES', 'INCV_COUNTER', 'INCV_SM', 'INCV_TOTAL',
        'PROF25_11', 'PROF25_17', 'PROF26_11', 'PROF26_17'
    ]
    df_baf = df_baf.iloc[:, :len(baf_cols)]
    df_baf.columns = baf_cols
    df_baf = df_baf.drop(columns=['Empty']).reset_index(drop=True)
    df_baf['TYPE_NORM'] = df_baf['TYPE'].apply(normalize_motor_name)
    for col in df_baf.columns[1:-1]:
        df_baf[col] = pd.to_numeric(df_baf[col], errors='coerce').fillna(0)
    df_baf = df_baf.reset_index(drop=True)

    return df_cash, df_reg, df_all33, df_adira33, df_baf

# --- HEADER UTAMA ---
col_head1, col_head2 = st.columns([2.2, 1.8])

with col_head1:
    st.title("Komparasi PL Yamaha Jember")
    st.markdown("⚠️ *hanya dipergunakan untuk kepentingan internal dealer MDS Area Jember*")
    if is_admin := True: # placeholder
        pass

with col_head2:
    mode_akses = st.text_input("🔑 Masukkan kode akses:", value="", type="password")

is_admin = (mode_akses.strip().upper() == "PAPINYACICOZRA")

if is_admin:
    st.caption("🔓 **Mode Admin (PAPINYACICOZRA):** Rincian 5 Parameter (TAC, BD, Insentif Total, Profit, Disc) + Fitur Copy Gambar HD.")
    col_head1.caption("💡 **Fitur Interaktif:** Klik angka angsuran/CASH untuk melihat rincian lengkap & Copy Gambar HD.")
else:
    st.caption("🔒 **Mode Dealer (Sales):** Tampilan standar (Disc, Insentif Salesman, DP - Disc, DP - Disc - Insentif).")
    col_head1.caption("💡 **Fitur Interaktif:** Klik angka angsuran atau tombol CASH untuk melihat rincian kalkulasi.")

uploaded_file = None
if is_admin:
    uploaded_file = st.file_uploader("📂 Upload Price List Baru (.xls / .xlsx)", type=["xls", "xlsx"])

try:
    if uploaded_file is not None:
        file_bytes = uploaded_file.getvalue()
        df_cash, df_reg, df_all33, df_adira33, df_baf = load_all_programs_from_bytes(file_bytes, uploaded_file.name)
    else:
        df_cash, df_reg, df_all33, df_adira33, df_baf = load_all_programs_from_bytes(DEFAULT_FILEPATH, DEFAULT_FILEPATH)
except Exception as e:
    st.error(f"Gagal memuat file Excel. Pastikan file '{DEFAULT_FILEPATH}' berada di folder yang sama. Error: {e}")
    st.stop()

all_motor_types = sorted(list(
    set(df_reg['TYPE_NORM']) | set(df_all33['TYPE_NORM']) | set(df_adira33['TYPE_NORM']) | set(df_baf['TYPE_NORM'])
))

motor_category_map = {m: categorize_motor(m) for m in all_motor_types}
available_categories = ['NMAX SERIES', 'FILLANO SERIES', 'FAZZIO SERIES', 'GEAR SERIES', 'OTHERS']

st.divider()

# --- FILTER 2 TINGKAT ---
col_f1, col_f2, col_f3 = st.columns([1.2, 1.8, 2.0])

with col_f1:
    pilih_kategori = st.selectbox("1️⃣ Pilih Kategori Seri:", available_categories)

filtered_motors = sorted([m for m in all_motor_types if motor_category_map[m] == pilih_kategori])
if not filtered_motors:
    filtered_motors = all_motor_types

with col_f2:
    pilih_motor = st.selectbox("2️⃣ Pilih Model Spesifik:", filtered_motors)

sub_cash = df_cash[df_cash['TYPE_NORM'] == pilih_motor]
sub_reg = df_reg[df_reg['TYPE_NORM'] == pilih_motor]
sub_all33 = df_all33[df_all33['TYPE_NORM'] == pilih_motor]
sub_adi33 = df_adira33[df_adira33['TYPE_NORM'] == pilih_motor]
sub_baf = df_baf[df_baf['TYPE_NORM'] == pilih_motor]

available_dps = sorted(list(
    set(sub_reg['DP_GROSS']) | set(sub_all33['DP_GROSS']) | set(sub_adi33['DP_GROSS']) | set(sub_baf['DP_GROSS'])
))

otr_val = 0
for s in [sub_cash, sub_reg, sub_all33, sub_adi33, sub_baf]:
    if not s.empty:
        otr_val = s['OTR'].iloc[0]
        break

with col_f3:
    selected_dps = st.multiselect(
        "3️⃣ Pilih DP Gross (Bisa pilih beberapa / semua):",
        options=available_dps,
        default=available_dps,
        format_func=format_rp_kosong
    )

st.markdown(
    f"**Model Terpilih:** `{pilih_motor}` &nbsp;|&nbsp; "
    f"**Harga OTR:** `{format_rp_kosong(otr_val)}` &nbsp;|&nbsp; "
    f"**Acuan Tahun Rangka:** `Noka 2026 (Berjalan)`"
)

# --- TOMBOL CASH ---
cash_row = sub_cash.iloc[0] if not sub_cash.empty else None
cash_json_attr = "{}"
if cash_row is not None:
    c_otr = float(cash_row['OTR'])
    c_disc26 = float(cash_row['DISKON_2026'] if 'DISKON_2026' in cash_row else cash_row.get('DISC_2026', 0.0))
    c_incv_sales = float(cash_row['INCV_SALES'] if 'INCV_SALES' in cash_row else 0.0)
    c_bd = float(cash_row['BD']) if 'BD' in cash_row else 0.0
    c_prof = float(cash_row['PROFIT_2026'] if 'PROFIT_2026' in cash_row else 0.0)
    
    if is_admin:
        cash_payload = {
            "mode": "admin",
            "motor": pilih_motor,
            "program": "CASH (INSTANSI / PO)",
            "dp_gross": format_rp_angka(c_otr),
            "angsuran": "TUNAI (CASH)",
            "tac": "Rp 0 (Cash)",
            "bd": format_rp_angka(c_bd),
            "incv_total": format_rp_angka(c_incv_sales),
            "profit": format_rp_angka(c_prof),
            "disc": format_rp_angka(c_disc26)
        }
    else:
        cash_payload = {
            "mode": "dealer",
            "motor": pilih_motor,
            "program": "CASH (INSTANSI / PO)",
            "dp_gross": format_rp_angka(c_otr),
            "angsuran": "TUNAI (CASH)",
            "disc": format_rp_angka(c_disc26),
            "incv": format_rp_angka(c_incv_sales),
            "dp_min_disc": format_rp_angka(c_otr - c_disc26),
            "dp_min_disc_incv": format_rp_angka(c_otr - c_disc26 - c_incv_sales)
        }
    cash_json_attr = json.dumps(cash_payload).replace('"', '&quot;')

st.markdown("---")

if not selected_dps:
    st.warning("Silakan pilih minimal satu nominal **DP Gross** pada kotak pilihan di atas.")
else:
    columns_order = [
        'REGULER (11x)', 'REGULER (17x)', 'REGULER (23x)', 'REGULER (29x)', 'REGULER (35x)', 'REGULER (47x)',
        'ALL LEASING 33 (33x)', 'ADIRA 33 (33x)', 'BAF TOP (11x)', 'BAF TOP (17x)'
    ]

    table_rows_html = ""
    csv_export_rows = []

    for dp in sorted(selected_dps):
        r_row = sub_reg[sub_reg['DP_GROSS'] == dp]
        al_row = sub_all33[sub_all33['DP_GROSS'] == dp]
        ad_row = sub_adi33[sub_adi33['DP_GROSS'] == dp]
        bf_row = sub_baf[sub_baf['DP_GROSS'] == dp]
        
        dp_str = format_rp_kosong(dp)
        csv_row = {'DP Gross': dp_str}
        
        cells_html = f'<td class="sticky-col">{dp_str}</td>'
        
        # REGULER
        for t in [11, 17, 23, 29, 35, 47]:
            col_label = f'REGULER ({t}x)'
            if not r_row.empty and r_row.iloc[0][f'ANGS_{t}'] > 0:
                r = r_row.iloc[0]
                angs = float(r[f'ANGS_{t}'])
                disc_val = float(r[f'DISC26_{t}']) if t != 47 else 0.0
                incv_sales = float(r['INCV_SALES'])
                bd_val = float(r['BD'])
                tac_val = float(r[f'TAC_{t}']) if f'TAC_{t}' in r else 0.0
                incv_tot = float(r['INCV_TOTAL'])
                prof_val = float(r[f'PROF26_{t}']) if f'PROF26_{t}' in r else 0.0
                
                if is_admin:
                    payload = {
                        "mode": "admin",
                        "motor": pilih_motor,
                        "program": col_label,
                        "dp_gross": format_rp_angka(dp),
                        "angsuran": format_rp_angka(angs),
                        "tac": format_rp_angka(tac_val),
                        "bd": format_rp_angka(bd_val),
                        "incv_total": format_rp_angka(incv_tot),
                        "profit": format_rp_angka(prof_val),
                        "disc": format_rp_angka(disc_val)
                    }
                else:
                    payload = {
                        "mode": "dealer",
                        "motor": pilih_motor,
                        "program": col_label,
                        "dp_gross": format_rp_angka(dp),
                        "angsuran": format_rp_angka(angs),
                        "disc": format_rp_angka(disc_val),
                        "incv": format_rp_angka(incv_sales),
                        "dp_min_disc": format_rp_angka(dp - disc_val),
                        "dp_min_disc_incv": format_rp_angka(dp - disc_val - incv_sales)
                    }
                json_attr = json.dumps(payload).replace('"', '&quot;')
                val_str = format_rp_kosong(angs)
                cells_html += f'<td class="clickable-cell" onclick="showNotif(event, {json_attr})">{val_str}</td>'
                csv_row[col_label] = val_str
            else:
                cells_html += '<td class="empty-cell"></td>'
                csv_row[col_label] = ""

        # ALL LEASING 33
        col_label = 'ALL LEASING 33 (33x)'
        if not al_row.empty and al_row.iloc[0]['ANGS_33'] > 0:
            al = al_row.iloc[0]
            angs = float(al['ANGS_33'])
            disc_val = float(al['DISC_2026'])
            incv_sales = float(al['INCV_SALES'])
            bd_val = float(al['BD'])
            tac_val = float(al['TAC'])
            incv_tot = float(al['INCV_TOTAL'])
            prof_val = float(al['PROFIT_2026'])
            
            if is_admin:
                payload = {
                    "mode": "admin",
                    "motor": pilih_motor,
                    "program": col_label,
                    "dp_gross": format_rp_angka(dp),
                    "angsuran": format_rp_angka(angs),
                    "tac": format_rp_angka(tac_val),
                    "bd": format_rp_angka(bd_val),
                    "incv_total": format_rp_angka(incv_tot),
                    "profit": format_rp_angka(prof_val),
                    "disc": format_rp_angka(disc_val)
                }
            else:
                payload = {
                    "mode": "dealer",
                    "motor": pilih_motor,
                    "program": col_label,
                    "dp_gross": format_rp_angka(dp),
                    "angsuran": format_rp_angka(angs),
                    "disc": format_rp_angka(disc_val),
                    "incv": format_rp_angka(incv_sales),
                    "dp_min_disc": format_rp_angka(dp - disc_val),
                    "dp_min_disc_incv": format_rp_angka(dp - disc_val - incv_sales)
                }
            json_attr = json.dumps(payload).replace('"', '&quot;')
            val_str = format_rp_kosong(angs)
            cells_html += f'<td class="clickable-cell" onclick="showNotif(event, {json_attr})">{val_str}</td>'
            csv_row[col_label] = val_str
        else:
            cells_html += '<td class="empty-cell"></td>'
            csv_row[col_label] = ""

        # ADIRA 33
        col_label = 'ADIRA 33 (33x)'
        if not ad_row.empty and ad_row.iloc[0]['ANGS_33'] > 0:
            ad = ad_row.iloc[0]
            angs = float(ad['ANGS_33'])
            disc_val = float(ad['DISC_2026'])
            incv_sales = float(ad['INCV_SALES'])
            bd_val = float(ad['BD'])
            tac_val = float(ad['TAC'])
            incv_tot = float(ad['INCV_TOTAL'])
            prof_val = float(ad['PROFIT_2026'])
            
            if is_admin:
                payload = {
                    "mode": "admin",
                    "motor": pilih_motor,
                    "program": col_label,
                    "dp_gross": format_rp_angka(dp),
                    "angsuran": format_rp_angka(angs),
                    "tac": format_rp_angka(tac_val),
                    "bd": format_rp_angka(bd_val),
                    "incv_total": format_rp_angka(incv_tot),
                    "profit": format_rp_angka(prof_val),
                    "disc": format_rp_angka(disc_val)
                }
            else:
                payload = {
                    "mode": "dealer",
                    "motor": pilih_motor,
                    "program": col_label,
                    "dp_gross": format_rp_angka(dp),
                    "angsuran": format_rp_angka(angs),
                    "disc": format_rp_angka(disc_val),
                    "incv": format_rp_angka(incv_sales),
                    "dp_min_disc": format_rp_angka(dp - disc_val),
                    "dp_min_disc_incv": format_rp_angka(dp - disc_val - incv_sales)
                }
            json_attr = json.dumps(payload).replace('"', '&quot;')
            val_str = format_rp_kosong(angs)
            cells_html += f'<td class="clickable-cell" onclick="showNotif(event, {json_attr})">{val_str}</td>'
            csv_row[col_label] = val_str
        else:
            cells_html += '<td class="empty-cell"></td>'
            csv_row[col_label] = ""

        # BAF TOP
        for t in [11, 17]:
            col_label = f'BAF TOP ({t}x)'
            if not bf_row.empty and bf_row.iloc[0][f'ANGS_{t}'] > 0:
                bf = bf_row.iloc[0]
                angs = float(bf[f'ANGS_{t}'])
                disc_val = float(bf[f'DISC26_{t}'])
                incv_sales = float(bf['INCV_SALES'])
                bd_val = float(bf['BD'])
                tac_val = float(bf[f'TAC_{t}'])
                incv_tot = float(bf['INCV_TOTAL'])
                prof_val = float(bf[f'PROF26_{t}'])
                
                if is_admin:
                    payload = {
                        "mode": "admin",
                        "motor": pilih_motor,
                        "program": col_label,
                        "dp_gross": format_rp_angka(dp),
                        "angsuran": format_rp_angka(angs),
                        "tac": format_rp_angka(tac_val),
                        "bd": format_rp_angka(bd_val),
                        "incv_total": format_rp_angka(incv_tot),
                        "profit": format_rp_angka(prof_val),
                        "disc": format_rp_angka(disc_val)
                    }
                else:
                    payload = {
                        "mode": "dealer",
                        "motor": pilih_motor,
                        "program": col_label,
                        "dp_gross": format_rp_angka(dp),
                        "angsuran": format_rp_angka(angs),
                        "disc": format_rp_angka(disc_val),
                        "incv": format_rp_angka(incv_sales),
                        "dp_min_disc": format_rp_angka(dp - disc_val),
                        "dp_min_disc_incv": format_rp_angka(dp - disc_val - incv_sales)
                    }
                json_attr = json.dumps(payload).replace('"', '&quot;')
                val_str = format_rp_kosong(angs)
                cells_html += f'<td class="clickable-cell" onclick="showNotif(event, {json_attr})">{val_str}</td>'
                csv_row[col_label] = val_str
            else:
                cells_html += '<td class="empty-cell"></td>'
                csv_row[col_label] = ""

        table_rows_html += f"<tr>{cells_html}</tr>"
        csv_export_rows.append(csv_row)

    headers_html = '<th class="sticky-header">DP Gross</th>' + "".join([f"<th>{c}</th>" for c in columns_order])

    # HTML Template disesuaikan berdasarkan is_admin
    if is_admin:
        html_code = f"""
        <!DOCTYPE html>
        <html>
        <head>
        <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
        <style>
            body {{ font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; margin: 0; padding: 0; background: transparent; }}
            .cash-btn-container {{ margin-bottom: 12px; }}
            .cash-btn {{ background-color: #047857; color: white; border: none; padding: 10px 18px; border-radius: 6px; font-weight: 600; font-size: 14px; cursor: pointer; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }}
            .cash-btn:hover {{ background-color: #065f46; }}
            .table-container {{ width: 100%; max-height: 580px; overflow: auto; border: 1px solid #cbd5e1; border-radius: 8px; box-shadow: 0 2px 6px rgba(0,0,0,0.06); background: #ffffff; }}
            table {{ border-collapse: separate; border-spacing: 0; width: 100%; font-size: 13.5px; }}
            th, td {{ padding: 11px 16px; text-align: center; white-space: nowrap; border-bottom: 1px solid #e2e8f0; border-right: 1px solid #e2e8f0; }}
            thead th {{ position: sticky; top: 0; background-color: #1e3a8a; color: #ffffff; font-weight: 600; z-index: 2; }}
            .sticky-header {{ position: sticky; left: 0; top: 0; z-index: 4 !important; background-color: #172554 !important; }}
            .sticky-col {{ position: sticky; left: 0; background-color: #eff6ff; font-weight: 700; color: #1e3a8a; z-index: 1; border-right: 2px solid #93c5fd; }}
            tr:hover td:not(.sticky-col):not(.empty-cell) {{ background-color: #f8fafc; }}
            .clickable-cell {{ cursor: pointer; color: #0f172a; font-weight: 500; transition: all 0.12s ease; }}
            .clickable-cell:hover {{ background-color: #fef08a !important; color: #854d0e; font-weight: 700; }}
            .active-cell {{ background-color: #fde047 !important; outline: 2px solid #ea580c; font-weight: 700; color: #9a3412; }}
            .empty-cell {{ background-color: #f8fafc; }}

            #notif-overlay {{ display: none; position: fixed; top: 0; left: 0; right: 0; bottom: 0; background: rgba(15, 23, 42, 0.4); z-index: 999; align-items: center; justify-content: center; }}
            #notif-box {{ background: #0f172a; color: #f8fafc; width: 440px; max-width: 92vw; border-radius: 12px; border: 2px solid #38bdf8; box-shadow: 0 20px 25px -5px rgba(0, 0, 0, 0.5); padding: 18px 22px; position: relative; animation: popIn 0.15s ease-out; }}
            @keyframes popIn {{ from {{ transform: scale(0.92); opacity: 0; }} to {{ transform: scale(1); opacity: 1; }} }}
            .notif-title {{ font-size: 15px; font-weight: 700; color: #38bdf8; margin-bottom: 4px; display: flex; justify-content: space-between; align-items: center; }}
            .notif-sub {{ font-size: 12.5px; color: #cbd5e1; margin-bottom: 12px; padding-bottom: 10px; border-bottom: 1px solid #334155; }}
            .metric-grid {{ display: grid; grid-template-columns: 1fr; gap: 8px; }}
            .metric-row {{ background: #1e293b; padding: 8px 12px; border-radius: 7px; display: flex; justify-content: space-between; align-items: center; border-left: 4px solid #38bdf8; cursor: pointer; transition: all 0.1s ease; }}
            .metric-row:hover {{ background: #334155; transform: scale(1.01); }}
            .metric-label {{ font-size: 12px; color: #e2e8f0; font-weight: 600; }}
            .metric-val {{ font-size: 14px; font-weight: 700; color: #ffffff; }}
            .notif-footer {{ margin-top: 10px; text-align: center; font-size: 11px; color: #94a3b8; }}
        </style>
        </head>
        <body>

        <div class="cash-btn-container">
            <button class="cash-btn" onclick='showNotif(event, {cash_json_attr})'>
                💵 CEK CASH (INSTANSI / PO) — Klik untuk Rincian
            </button>
        </div>

        <div class="table-container">
            <table>
                <thead><tr>{headers_html}</tr></thead>
                <tbody>{table_rows_html}</tbody>
            </table>
        </div>

        <div id="notif-overlay" onclick="closeNotif(event)">
            <div id="notif-box" onclick="event.stopPropagation()">
                <div id="capture-header">
                    <div class="notif-title">
                        <span id="n-program">PROGRAM</span>
                        <span style="font-size:13px; color:#fde047;" id="n-angsuran"></span>
                    </div>
                    <div class="notif-sub">
                        Tipe: <strong id="n-motor" style="color:#fff;"></strong> &nbsp;|&nbsp; OTR / DP: <strong id="n-dp" style="color:#fff;"></strong>
                    </div>
                </div>

                <div class="metric-grid" id="capture-grid">
                    <div class="metric-row" id="row-tac" style="border-left-color: #60a5fa;" onclick="copySingleRow('row-tac', '1. TAC', document.getElementById('n-tac').innerText)">
                        <span class="metric-label">1. TAC</span>
                        <span class="metric-val" id="n-tac"></span>
                    </div>
                    <div class="metric-row" id="row-bd" style="border-left-color: #38bdf8;" onclick="copySingleRow('row-bd', '2. BASIC DISC (BD)', document.getElementById('n-bd').innerText)">
                        <span class="metric-label">2. BASIC DISC (BD)</span>
                        <span class="metric-val" id="n-bd"></span>
                    </div>
                    <div class="metric-row" id="row-incv" style="border-left-color: #22c55e;" onclick="copySingleRow('row-incv', '3. INSENTIF TOTAL', document.getElementById('n-incv').innerText)">
                        <span class="metric-label">3. INSENTIF TOTAL</span>
                        <span class="metric-val" style="color:#4ade80;" id="n-incv"></span>
                    </div>
                    <div class="metric-row" id="row-prof" style="border-left-color: #f59e0b;" onclick="copySingleRow('row-prof', '4. PROFIT', document.getElementById('n-prof').innerText)">
                        <span class="metric-label">4. PROFIT</span>
                        <span class="metric-val" style="color:#fbbf24;" id="n-prof"></span>
                    </div>
                    <div class="metric-row" id="row-disc" style="border-left-color: #f472b6;" onclick="copySingleRow('row-disc', '5. DISCOUNT (Noka 2026)', document.getElementById('n-disc').innerText)">
                        <span class="metric-label">5. DISCOUNT (Noka 2026)</span>
                        <span class="metric-val" id="n-disc"></span>
                    </div>
                </div>

                <div class="notif-footer" id="capture-footer">
                    💡 <b>Klik langsung pada salah satu baris di atas</b> untuk menyalin bagian tersebut ke Clipboard (HD Image). Atau klik di luar untuk menutup.
                </div>
            </div>
        </div>

        <script>
            let currentActiveCell = null;
            function showNotif(event, data) {{
                event.stopPropagation();
                if (currentActiveCell && currentActiveCell.classList) {{ currentActiveCell.classList.remove('active-cell'); }}
                if (event.currentTarget && event.currentTarget.classList) {{ currentActiveCell = event.currentTarget; currentActiveCell.classList.add('active-cell'); }}

                document.getElementById('n-program').innerText = data.program;
                document.getElementById('n-angsuran').innerText = data.angsuran.includes("T") ? data.angsuran : "Angsuran: " + data.angsuran;
                document.getElementById('n-motor').innerText = data.motor;
                document.getElementById('n-dp').innerText = data.dp_gross;
                
                document.getElementById('n-tac').innerText = data.tac;
                document.getElementById('n-bd').innerText = data.bd;
                document.getElementById('n-incv').innerText = data.incv_total;
                document.getElementById('n-prof').innerText = data.profit;
                document.getElementById('n-disc').innerText = data.disc;

                document.getElementById('notif-overlay').style.display = 'flex';
            }}
            function closeNotif(event) {{
                document.getElementById('notif-overlay').style.display = 'none';
                if (currentActiveCell && currentActiveCell.classList) {{ currentActiveCell.classList.remove('active-cell'); currentActiveCell = null; }}
            }}
            function copySingleRow(rowId, labelName, valText) {{
                const tempBox = document.createElement('div');
                tempBox.style.background = '#0f172a';
                tempBox.style.color = '#f8fafc';
                tempBox.style.width = '420px';
                tempBox.style.borderRadius = '12px';
                tempBox.style.border = '2px solid #38bdf8';
                tempBox.style.padding = '18px 22px';
                tempBox.style.fontFamily = '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif';
                tempBox.style.boxSizing = 'border-box';

                const headerHtml = document.getElementById('capture-header').innerHTML;
                tempBox.innerHTML = headerHtml + `
                    <div class="metric-grid" style="margin-top: 10px;">
                        <div class="metric-row" style="background: #1e293b; padding: 10px 14px; border-radius: 7px; display: flex; justify-content: space-between; align-items: center; border-left: 4px solid #38bdf8;">
                            <span style="font-size: 13px; color: #e2e8f0; font-weight: 600;">${{labelName}}</span>
                            <span style="font-size: 15px; font-weight: 700; color: #ffffff;">${{valText}}</span>
                        </div>
                    </div>
                `;
                document.body.appendChild(tempBox);
                html2canvas(tempBox, {{ scale: 3, backgroundColor: '#0f172a', logging: false }}).then(canvas => {{
                    document.body.removeChild(tempBox);
                    canvas.toBlob(blob => {{
                        try {{
                            const item = new ClipboardItem({{'image/png': blob}});
                            navigator.clipboard.write([item]).then(() => {{
                                const rowEl = document.getElementById(rowId);
                                const originalBg = rowEl.style.background;
                                rowEl.style.background = '#16a34a';
                                setTimeout(() => {{ rowEl.style.background = originalBg; }}, 1000);
                            }});
                        }} catch (e) {{ alert("Gagal menyalin gambar."); }}
                    }});
                }});
            }}
        </script>
        </body>
        </html>
        """
    else:
        # MODE DEALER (Klasik: Disc, Insentif, DP - Disc, DP - Disc - Insentif)
        html_code = f"""
        <!DOCTYPE html>
        <html>
        <head>
        <style>
            body {{ font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; margin: 0; padding: 0; background: transparent; }}
            .cash-btn-container {{ margin-bottom: 12px; }}
            .cash-btn {{ background-color: #047857; color: white; border: none; padding: 10px 18px; border-radius: 6px; font-weight: 600; font-size: 14px; cursor: pointer; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }}
            .cash-btn:hover {{ background-color: #065f46; }}
            .table-container {{ width: 100%; max-height: 580px; overflow: auto; border: 1px solid #cbd5e1; border-radius: 8px; box-shadow: 0 2px 6px rgba(0,0,0,0.06); background: #ffffff; }}
            table {{ border-collapse: separate; border-spacing: 0; width: 100%; font-size: 13.5px; }}
            th, td {{ padding: 11px 16px; text-align: center; white-space: nowrap; border-bottom: 1px solid #e2e8f0; border-right: 1px solid #e2e8f0; }}
            thead th {{ position: sticky; top: 0; background-color: #1e3a8a; color: #ffffff; font-weight: 600; z-index: 2; }}
            .sticky-header {{ position: sticky; left: 0; top: 0; z-index: 4 !important; background-color: #172554 !important; }}
            .sticky-col {{ position: sticky; left: 0; background-color: #eff6ff; font-weight: 700; color: #1e3a8a; z-index: 1; border-right: 2px solid #93c5fd; }}
            tr:hover td:not(.sticky-col):not(.empty-cell) {{ background-color: #f8fafc; }}
            .clickable-cell {{ cursor: pointer; color: #0f172a; font-weight: 500; transition: all 0.12s ease; }}
            .clickable-cell:hover {{ background-color: #fef08a !important; color: #854d0e; font-weight: 700; }}
            .active-cell {{ background-color: #fde047 !important; outline: 2px solid #ea580c; font-weight: 700; color: #9a3412; }}
            .empty-cell {{ background-color: #f8fafc; }}

            #notif-overlay {{ display: none; position: fixed; top: 0; left: 0; right: 0; bottom: 0; background: rgba(15, 23, 42, 0.4); z-index: 999; align-items: center; justify-content: center; }}
            #notif-box {{ background: #0f172a; color: #f8fafc; width: 440px; max-width: 92vw; border-radius: 12px; border: 2px solid #38bdf8; box-shadow: 0 20px 25px -5px rgba(0, 0, 0, 0.5); padding: 18px 22px; position: relative; animation: popIn 0.15s ease-out; }}
            @keyframes popIn {{ from {{ transform: scale(0.92); opacity: 0; }} to {{ transform: scale(1); opacity: 1; }} }}
            .notif-title {{ font-size: 15px; font-weight: 700; color: #38bdf8; margin-bottom: 4px; display: flex; justify-content: space-between; align-items: center; }}
            .notif-sub {{ font-size: 12.5px; color: #cbd5e1; margin-bottom: 12px; padding-bottom: 10px; border-bottom: 1px solid #334155; }}
            .metric-grid {{ display: grid; grid-template-columns: 1fr; gap: 8px; }}
            .metric-row {{ background: #1e293b; padding: 9px 12px; border-radius: 7px; display: flex; justify-content: space-between; align-items: center; border-left: 4px solid #38bdf8; }}
            .metric-row.highlight-1 {{ border-left-color: #22c55e; background: #14532d33; }}
            .metric-row.highlight-2 {{ border-left-color: #f59e0b; background: #78350f33; }}
            .metric-label {{ font-size: 12.5px; color: #e2e8f0; font-weight: 600; }}
            .metric-note {{ font-size: 11px; color: #94a3b8; display: block; margin-top: 2px; font-weight: 400; }}
            .metric-val {{ font-size: 15px; font-weight: 700; color: #ffffff; }}
            .notif-footer {{ margin-top: 12px; text-align: center; font-size: 11.5px; color: #94a3b8; }}
        </style>
        </head>
        <body>

        <div class="cash-btn-container">
            <button class="cash-btn" onclick='showNotif(event, {cash_json_attr})'>
                💵 CEK CASH (INSTANSI / PO) — Klik untuk Rincian
            </button>
        </div>

        <div class="table-container">
            <table>
                <thead><tr>{headers_html}</tr></thead>
                <tbody>{table_rows_html}</tbody>
            </table>
        </div>

        <div id="notif-overlay" onclick="closeNotif(event)">
            <div id="notif-box" onclick="event.stopPropagation()">
                <div class="notif-title">
                    <span id="n-program">PROGRAM</span>
                    <span style="font-size:13px; color:#fde047;" id="n-angsuran"></span>
                </div>
                <div class="notif-sub">
                    Tipe: <strong id="n-motor" style="color:#fff;"></strong> &nbsp;|&nbsp; OTR / DP: <strong id="n-dp" style="color:#fff;"></strong>
                </div>

                <div class="metric-grid">
                    <div class="metric-row">
                        <div>
                            <span class="metric-label">1. DISC (Noka 2026)</span>
                            <span class="metric-note">Pedoman diskon tahun berjalan</span>
                        </div>
                        <span class="metric-val" id="n-disc"></span>
                    </div>

                    <div class="metric-row">
                        <div>
                            <span class="metric-label">2. INSENTIF (Salesman)</span>
                            <span class="metric-note">Insentif khusus salesman per unit</span>
                        </div>
                        <span class="metric-val" id="n-incv"></span>
                    </div>

                    <div class="metric-row highlight-1">
                        <div>
                            <span class="metric-label">3. DP - DISC (DP Bayar)</span>
                            <span class="metric-note">DP Gross dikurangi Diskon 2026</span>
                        </div>
                        <span class="metric-val" style="color:#4ade80;" id="n-dp-disc"></span>
                    </div>

                    <div class="metric-row highlight-2">
                        <div>
                            <span class="metric-label">4. DP - DISC - INSENTIF</span>
                            <span class="metric-note">Nett setelah dikurangi Diskon & Insentif Salesman</span>
                        </div>
                        <span class="metric-val" style="color:#fbbf24;" id="n-dp-disc-incv"></span>
                    </div>
                </div>

                <div class="notif-footer">
                    💡 Klik di mana saja di luar kotak notifikasi ini untuk menutup
                </div>
            </div>
        </div>

        <script>
            let currentActiveCell = null;
            function showNotif(event, data) {{
                event.stopPropagation();
                if (currentActiveCell && currentActiveCell.classList) {{ currentActiveCell.classList.remove('active-cell'); }}
                if (event.currentTarget && event.currentTarget.classList) {{ currentActiveCell = event.currentTarget; currentActiveCell.classList.add('active-cell'); }}

                document.getElementById('n-program').innerText = data.program;
                document.getElementById('n-angsuran').innerText = data.angsuran.includes("T") ? data.angsuran : "Angsuran: " + data.angsuran;
                document.getElementById('n-motor').innerText = data.motor;
                document.getElementById('n-dp').innerText = data.dp_gross;
                
                document.getElementById('n-disc').innerText = data.disc;
                document.getElementById('n-incv').innerText = data.incv;
                document.getElementById('n-dp-disc').innerText = data.dp_min_disc;
                document.getElementById('n-dp-disc-incv').innerText = data.dp_min_disc_incv;

                document.getElementById('notif-overlay').style.display = 'flex';
            }}
            function closeNotif(event) {{
                document.getElementById('notif-overlay').style.display = 'none';
                if (currentActiveCell && currentActiveCell.classList) {{ currentActiveCell.classList.remove('active-cell'); currentActiveCell = null; }}
            }}
        </script>
        </body>
        </html>
        """

    components.html(html_code, height=660, scrolling=False)

    if is_admin:
        df_csv = pd.DataFrame(csv_export_rows)
        csv_bytes = df_csv.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📥 Download Tabel Komparasi (CSV)",
            data=csv_bytes,
            file_name=f"Komparasi_Angsuran_{pilih_motor.replace(' ', '_')}.csv",
            mime="text/csv"
        )
