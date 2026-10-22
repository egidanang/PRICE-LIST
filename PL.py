import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import numpy as np
import re
import json
import io

# Konfigurasi Halaman Full Wide (Tanpa Sidebar Kiri)
st.set_page_config(
    page_title="Dashboard Komparasi Angsuran Yamaha",
    page_icon="🏍️",
    layout="wide"
)

DEFAULT_FILEPATH = "PL #1 SEPTEMBER 26.xls"

def format_rp_kosong(val):
    """Format angka ke Rupiah. Jika tidak ada data / <= 0 / NaN, kembalikan string kosong."""
    if pd.isna(val) or val is None or val <= 0:
        return ""
    try:
        return f"Rp {int(round(float(val))):,}".replace(",", ".")
    except:
        return ""

def format_rp_angka(val):
    """Format angka ke Rupiah termasuk Rp 0 untuk ditampilkan di dalam notifikasi popup."""
    if pd.isna(val) or val is None:
        return "Rp 0"
    try:
        return f"Rp {int(round(float(val))):,}".replace(",", ".")
    except:
        return "Rp 0"

def normalize_motor_name(s):
    """Menyamakan format penulisan nama motor antar sheet agar semua program terhubung."""
    s = str(s).strip().upper().replace('"', '')
    s = s.replace('/', ' / ')
    s = re.sub(r'\s+', ' ', s).strip()
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
    }
    return alias_map.get(s, s)

def categorize_motor(name):
    """Mengelompokkan tipe motor ke dalam kategori seri."""
    n = name.upper()
    if 'NMAX' in n or 'N - MAX' in n:
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
    # Gunakan BytesIO untuk membaca file excel dari uploader atau default path
    if isinstance(file_bytes, bytes):
        excel_file = io.BytesIO(file_bytes)
    else:
        excel_file = file_bytes

    # 1. Sheet REGULER
    df_reg_raw = pd.read_excel(excel_file, sheet_name='REGULER ')
    df_reg = df_reg_raw[
        df_reg_raw.iloc[:, 0].notna() & 
        (df_reg_raw.iloc[:, 0].astype(str).str.strip() != 'BACK') &
        pd.to_numeric(df_reg_raw.iloc[:, 1], errors='coerce').notna() &
        pd.to_numeric(df_reg_raw.iloc[:, 3], errors='coerce').notna()
    ].copy()
    df_reg.columns = [
        'TYPE', 'OTR', 'DP_PCT', 'DP_GROSS', 
        'ANGS_11', 'ANGS_17', 'ANGS_23', 'ANGS_29', 'ANGS_35', 'ANGS_47', 'Empty_10',
        'INCV_KOOR', 'INCV_SALES', 'INCV_COUNTER', 'INCV_SM', 'INCV_TOTAL', 'Empty_16',
        'DISC25_11', 'DISC25_17', 'DISC25_23', 'DISC25_29', 'DISC25_35', 'DISC25_47',
        'DISC26_11', 'DISC26_17', 'DISC26_23', 'DISC26_29', 'DISC26_35'
    ]
    df_reg = df_reg.drop(columns=['Empty_10', 'Empty_16']).reset_index(drop=True)
    df_reg['TYPE_NORM'] = df_reg['TYPE'].apply(normalize_motor_name)
    for col in df_reg.columns[1:-1]:
        df_reg[col] = pd.to_numeric(df_reg[col], errors='coerce').fillna(0)

    # 2. Sheet ALL LEASING 33 & ADIRA 33
    cols_33 = [
        'TYPE', 'OTR', 'DP_PCT', 'DP_GROSS', 'ANGS_33',
        'DISC_2023', 'DISC_2025', 'DISC_2026',
        'INCV_KOOR', 'INCV_SALES', 'INCV_COUNTER', 'INCV_SM', 'INCV_TOTAL', 'PROFIT_2023'
    ]
    def parse_sheet_33(sheet_name):
        if isinstance(file_bytes, bytes):
            ef = io.BytesIO(file_bytes)
        else:
            ef = file_bytes
        raw = pd.read_excel(ef, sheet_name=sheet_name).iloc[:, :14]
        clean = raw[
            raw.iloc[:, 0].notna() & 
            (raw.iloc[:, 0].astype(str).str.strip() != 'BACK') &
            pd.to_numeric(raw.iloc[:, 1], errors='coerce').notna() &
            pd.to_numeric(raw.iloc[:, 3], errors='coerce').notna()
        ].copy()
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
    df_baf_raw = pd.read_excel(ef, sheet_name='BAF TOP').iloc[:, 1:16]
    df_baf = df_baf_raw[
        df_baf_raw.iloc[:, 0].notna() & 
        (~df_baf_raw.iloc[:, 0].astype(str).str.strip().isin(['BACK', 'TYPE'])) &
        pd.to_numeric(df_baf_raw.iloc[:, 2], errors='coerce').notna() &
        pd.to_numeric(df_baf_raw.iloc[:, 3], errors='coerce').notna()
    ].copy()
    df_baf.columns = [
        'TYPE', 'DP_PCT', 'OTR', 'DP_GROSS', 'ANGS_11', 'ANGS_17',
        'DISC25_11', 'DISC25_17', 'DISC26_11', 'DISC26_17',
        'INCV_KOOR', 'INCV_SALES', 'INCV_COUNTER', 'INCV_SM', 'INCV_TOTAL'
    ]
    df_baf['TYPE_NORM'] = df_baf['TYPE'].apply(normalize_motor_name)
    for col in df_baf.columns[1:-1]:
        df_baf[col] = pd.to_numeric(df_baf[col], errors='coerce').fillna(0)
    df_baf = df_baf.reset_index(drop=True)

    return df_reg, df_all33, df_adira33, df_baf

# --- HEADER UTAMA & TOMBOL UPLOAD DI POJOK KANAN ATAS ---
col_head1, col_head2 = st.columns([2.5, 1.5])

with col_head1:
    st.title("🏍️ Dashboard Komparasi Angsuran Yamaha")
    st.caption("💡 **Fitur Interaktif:** Klik angka angsuran untuk melihat rincian DISC, Insentif Salesman, & DP Bayar.")

with col_head2:
    uploaded_file = st.file_uploader("📂 Upload Price List Baru (.xls / .xlsx)", type=["xls", "xlsx"])

# Load data (menggunakan file uploader jika ada, atau file default)
try:
    if uploaded_file is not None:
        file_bytes = uploaded_file.getvalue()
        df_reg, df_all33, df_adira33, df_baf = load_all_programs_from_bytes(file_bytes, uploaded_file.name)
        st.success(f"Berhasil memuat file baru: **{uploaded_file.name}**")
    else:
        df_reg, df_all33, df_adira33, df_baf = load_all_programs_from_bytes(DEFAULT_FILEPATH, DEFAULT_FILEPATH)
except Exception as e:
    st.error(f"Gagal memuat file Excel. Pastikan format file sesuai. Error: {e}")
    st.stop()

# Kumpulkan seluruh tipe motor gabungan dari semua program
all_motor_types = sorted(list(
    set(df_reg['TYPE_NORM']) | set(df_all33['TYPE_NORM']) | set(df_adira33['TYPE_NORM']) | set(df_baf['TYPE_NORM'])
))

# Pemetaan Kategori Seri Motor
motor_category_map = {m: categorize_motor(m) for m in all_motor_types}
available_categories = ['NMAX SERIES', 'FILLANO SERIES', 'FAZZIO SERIES', 'GEAR SERIES', 'OTHERS']

st.divider()

# --- FILTER 2 TINGKAT: KATEGORI SERI & MODEL SPESIFIK ---
col_f1, col_f2, col_f3 = st.columns([1.2, 1.8, 2.0])

with col_f1:
    pilih_kategori = st.selectbox("1️⃣ Pilih Kategori Seri:", available_categories)

# Filter daftar motor berdasarkan kategori yang dipilih di dropdown 1
filtered_motors = sorted([m for m in all_motor_types if motor_category_map[m] == pilih_kategori])
if not filtered_motors:
    filtered_motors = all_motor_types # fallback jika kosong

with col_f2:
    pilih_motor = st.selectbox("2️⃣ Pilih Model Spesifik:", filtered_motors)

# Filter masing-masing sheet berdasarkan motor terpilih
sub_reg = df_reg[df_reg['TYPE_NORM'] == pilih_motor]
sub_all33 = df_all33[df_all33['TYPE_NORM'] == pilih_motor]
sub_adi33 = df_adira33[df_adira33['TYPE_NORM'] == pilih_motor]
sub_baf = df_baf[df_baf['TYPE_NORM'] == pilih_motor]

available_dps = sorted(list(
    set(sub_reg['DP_GROSS']) | set(sub_all33['DP_GROSS']) | set(sub_adi33['DP_GROSS']) | set(sub_baf['DP_GROSS'])
))

otr_val = 0
for s in [sub_reg, sub_all33, sub_adi33, sub_baf]:
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
    f"**Acuan Tahun Rangka:** `Noka 2026 (Berjalan)` &nbsp;|&nbsp; "
    f"**Acuan Insentif:** `Insentif Salesman`"
)
st.markdown("---")

# --- MEMBANGUN TABEL HTML INTERAKTIF FULL LEBAR ---
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
        
        # 1. REGULER (11x - 47x)
        for t in [11, 17, 23, 29, 35, 47]:
            col_label = f'REGULER ({t}x)'
            if not r_row.empty and r_row.iloc[0][f'ANGS_{t}'] > 0:
                r = r_row.iloc[0]
                angs = float(r[f'ANGS_{t}'])
                disc26 = float(r[f'DISC26_{t}']) if t != 47 else 0.0
                incv_sales = float(r['INCV_SALES'])
                
                payload = {
                    "motor": pilih_motor,
                    "program": col_label,
                    "dp_gross": format_rp_angka(dp),
                    "angsuran": format_rp_angka(angs),
                    "tahun_noka": "2026",
                    "disc": format_rp_angka(disc26),
                    "insentif": format_rp_angka(incv_sales),
                    "dp_min_disc": format_rp_angka(dp - disc26),
                    "dp_min_disc_incv": format_rp_angka(dp - disc26 - incv_sales)
                }
                json_attr = json.dumps(payload).replace('"', '&quot;')
                val_str = format_rp_kosong(angs)
                cells_html += f'<td class="clickable-cell" onclick="showNotif(event, {json_attr})">{val_str}</td>'
                csv_row[col_label] = val_str
            else:
                cells_html += '<td class="empty-cell"></td>'
                csv_row[col_label] = ""

        # 2. ALL LEASING 33 (33x)
        col_label = 'ALL LEASING 33 (33x)'
        if not al_row.empty and al_row.iloc[0]['ANGS_33'] > 0:
            al = al_row.iloc[0]
            angs = float(al['ANGS_33'])
            disc26 = float(al['DISC_2026'])
            incv_sales = float(al['INCV_SALES'])
            
            payload = {
                "motor": pilih_motor,
                "program": col_label,
                "dp_gross": format_rp_angka(dp),
                "angsuran": format_rp_angka(angs),
                "tahun_noka": "2026",
                "disc": format_rp_angka(disc26),
                "insentif": format_rp_angka(incv_sales),
                "dp_min_disc": format_rp_angka(dp - disc26),
                "dp_min_disc_incv": format_rp_angka(dp - disc26 - incv_sales)
            }
            json_attr = json.dumps(payload).replace('"', '&quot;')
            val_str = format_rp_kosong(angs)
            cells_html += f'<td class="clickable-cell" onclick="showNotif(event, {json_attr})">{val_str}</td>'
            csv_row[col_label] = val_str
        else:
            cells_html += '<td class="empty-cell"></td>'
            csv_row[col_label] = ""

        # 3. ADIRA 33 (33x)
        col_label = 'ADIRA 33 (33x)'
        if not ad_row.empty and ad_row.iloc[0]['ANGS_33'] > 0:
            ad = ad_row.iloc[0]
            angs = float(ad['ANGS_33'])
            disc26 = float(ad['DISC_2026'])
            incv_sales = float(ad['INCV_SALES'])
            
            payload = {
                "motor": pilih_motor,
                "program": col_label,
                "dp_gross": format_rp_angka(dp),
                "angsuran": format_rp_angka(angs),
                "tahun_noka": "2026",
                "disc": format_rp_angka(disc26),
                "insentif": format_rp_angka(incv_sales),
                "dp_min_disc": format_rp_angka(dp - disc26),
                "dp_min_disc_incv": format_rp_angka(dp - disc26 - incv_sales)
            }
            json_attr = json.dumps(payload).replace('"', '&quot;')
            val_str = format_rp_kosong(angs)
            cells_html += f'<td class="clickable-cell" onclick="showNotif(event, {json_attr})">{val_str}</td>'
            csv_row[col_label] = val_str
        else:
            cells_html += '<td class="empty-cell"></td>'
            csv_row[col_label] = ""

        # 4. BAF TOP (11x & 17x)
        for t in [11, 17]:
            col_label = f'BAF TOP ({t}x)'
            if not bf_row.empty and bf_row.iloc[0][f'ANGS_{t}'] > 0:
                bf = bf_row.iloc[0]
                angs = float(bf[f'ANGS_{t}'])
                disc26 = float(bf[f'DISC26_{t}'])
                incv_sales = float(bf['INCV_SALES'])
                
                payload = {
                    "motor": pilih_motor,
                    "program": col_label,
                    "dp_gross": format_rp_angka(dp),
                    "angsuran": format_rp_angka(angs),
                    "tahun_noka": "2026",
                    "disc": format_rp_angka(disc26),
                    "insentif": format_rp_angka(incv_sales),
                    "dp_min_disc": format_rp_angka(dp - disc26),
                    "dp_min_disc_incv": format_rp_angka(dp - disc26 - incv_sales)
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

    html_code = f"""
    <!DOCTYPE html>
    <html>
    <head>
    <style>
        body {{
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
            margin: 0;
            padding: 0;
            background: transparent;
        }}
        .table-container {{
            width: 100%;
            max-height: 600px;
            overflow: auto;
            border: 1px solid #cbd5e1;
            border-radius: 8px;
            box-shadow: 0 2px 6px rgba(0,0,0,0.06);
            background: #ffffff;
        }}
        table {{
            border-collapse: separate;
            border-spacing: 0;
            width: 100%;
            font-size: 13.5px;
        }}
        th, td {{
            padding: 11px 16px;
            text-align: center;
            white-space: nowrap;
            border-bottom: 1px solid #e2e8f0;
            border-right: 1px solid #e2e8f0;
        }}
        thead th {{
            position: sticky;
            top: 0;
            background-color: #1e3a8a;
            color: #ffffff;
            font-weight: 600;
            z-index: 2;
        }}
        .sticky-header {{
            position: sticky;
            left: 0;
            top: 0;
            z-index: 4 !important;
            background-color: #172554 !important;
        }}
        .sticky-col {{
            position: sticky;
            left: 0;
            background-color: #eff6ff;
            font-weight: 700;
            color: #1e3a8a;
            z-index: 1;
            border-right: 2px solid #93c5fd;
        }}
        tr:hover td:not(.sticky-col):not(.empty-cell) {{
            background-color: #f8fafc;
        }}
        .clickable-cell {{
            cursor: pointer;
            color: #0f172a;
            font-weight: 500;
            transition: all 0.12s ease;
        }}
        .clickable-cell:hover {{
            background-color: #fef08a !important;
            color: #854d0e;
            font-weight: 700;
        }}
        .active-cell {{
            background-color: #fde047 !important;
            outline: 2px solid #ea580c;
            font-weight: 700;
            color: #9a3412;
        }}
        .empty-cell {{
            background-color: #f8fafc;
        }}

        /* Overlay & Popup Modal Notification */
        #notif-overlay {{
            display: none;
            position: fixed;
            top: 0; left: 0; right: 0; bottom: 0;
            background: rgba(15, 23, 42, 0.4);
            z-index: 999;
            align-items: center;
            justify-content: center;
        }}
        #notif-box {{
            background: #0f172a;
            color: #f8fafc;
            width: 440px;
            max-width: 92vw;
            border-radius: 12px;
            border: 2px solid #38bdf8;
            box-shadow: 0 20px 25px -5px rgba(0, 0, 0, 0.5);
            padding: 18px 22px;
            position: relative;
            animation: popIn 0.15s ease-out;
        }}
        @keyframes popIn {{
            from {{ transform: scale(0.92); opacity: 0; }}
            to {{ transform: scale(1); opacity: 1; }}
        }}
        .notif-title {{
            font-size: 15px;
            font-weight: 700;
            color: #38bdf8;
            margin-bottom: 4px;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }}
        .notif-sub {{
            font-size: 12.5px;
            color: #cbd5e1;
            margin-bottom: 12px;
            padding-bottom: 10px;
            border-bottom: 1px solid #334155;
        }}
        .metric-grid {{
            display: grid;
            grid-template-columns: 1fr;
            gap: 8px;
        }}
        .metric-row {{
            background: #1e293b;
            padding: 9px 12px;
            border-radius: 7px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            border-left: 4px solid #38bdf8;
        }}
        .metric-row.highlight-1 {{
            border-left-color: #22c55e;
            background: #14532d33;
        }}
        .metric-row.highlight-2 {{
            border-left-color: #f59e0b;
            background: #78350f33;
        }}
        .metric-label {{
            font-size: 12.5px;
            color: #e2e8f0;
            font-weight: 600;
        }}
        .metric-note {{
            font-size: 11px;
            color: #94a3b8;
            display: block;
            margin-top: 2px;
            font-weight: 400;
        }}
        .metric-val {{
            font-size: 15px;
            font-weight: 700;
            color: #ffffff;
        }}
        .notif-footer {{
            margin-top: 12px;
            text-align: center;
            font-size: 11.5px;
            color: #94a3b8;
        }}
    </style>
    </head>
    <body>

    <div class="table-container">
        <table>
            <thead>
                <tr>{headers_html}</tr>
            </thead>
            <tbody>
                {table_rows_html}
            </tbody>
        </table>
    </div>

    <!-- Kotak Notifikasi Popup (Klik di luar untuk menutup) -->
    <div id="notif-overlay" onclick="closeNotif(event)">
        <div id="notif-box" onclick="event.stopPropagation()">
            <div class="notif-title">
                <span id="n-program">PROGRAM</span>
                <span style="font-size:13px; color:#fde047;" id="n-angsuran"></span>
            </div>
            <div class="notif-sub">
                Tipe: <strong id="n-motor" style="color:#fff;"></strong> &nbsp;|&nbsp; DP Gross: <strong id="n-dp" style="color:#fff;"></strong>
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
            if (currentActiveCell) {{
                currentActiveCell.classList.remove('active-cell');
            }}
            currentActiveCell = event.currentTarget;
            currentActiveCell.classList.add('active-cell');

            document.getElementById('n-program').innerText = data.program;
            document.getElementById('n-angsuran').innerText = "Angsuran: " + data.angsuran;
            document.getElementById('n-motor').innerText = data.motor;
            document.getElementById('n-dp').innerText = data.dp_gross;
            document.getElementById('n-disc').innerText = data.disc;
            document.getElementById('n-incv').innerText = data.insentif;
            document.getElementById('n-dp-disc').innerText = data.dp_min_disc;
            document.getElementById('n-dp-disc-incv').innerText = data.dp_min_disc_incv;

            document.getElementById('notif-overlay').style.display = 'flex';
        }}

        function closeNotif(event) {{
            document.getElementById('notif-overlay').style.display = 'none';
            if (currentActiveCell) {{
                currentActiveCell.classList.remove('active-cell');
                currentActiveCell = null;
            }}
        }}
    </script>
    </body>
    </html>
    """

    # Render Tabel HTML Melebar Full Width
    components.html(html_code, height=620, scrolling=False)

    # Tombol Download CSV
    df_csv = pd.DataFrame(csv_export_rows)
    csv_bytes = df_csv.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="📥 Download Tabel Komparasi (CSV)",
        data=csv_bytes,
        file_name=f"Komparasi_Angsuran_{pilih_motor.replace(' ', '_')}.csv",
        mime="text/csv"
    )