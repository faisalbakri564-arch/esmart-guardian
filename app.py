import pandas as pd
import streamlit as st
from io import BytesIO
import os
import sqlite3

# ==========================================
# DATABASE HELPER (ANTI HILANG DATA)
# ==========================================
DB_NAME = 'esmart_guardian.db'

def save_df(table_name, df):
    with sqlite3.connect(DB_NAME) as conn:
        df.astype(str).to_sql(table_name, conn, if_exists='replace', index=False)

def load_df(table_name):
    try:
        with sqlite3.connect(DB_NAME) as conn:
            return pd.read_sql(f'SELECT * FROM {table_name}', conn)
    except:
        return pd.DataFrame()

def save_config(key, value):
    with sqlite3.connect(DB_NAME) as conn:
        pd.DataFrame({'key': [key], 'value': [value]}).to_sql(f'config_{key}', conn, if_exists='replace', index=False)

def load_config(key, default=""):
    try:
        with sqlite3.connect(DB_NAME) as conn:
            df = pd.read_sql(f'SELECT value FROM config_{key}', conn)
            return df['value'].iloc[0]
    except:
        return default

# ==========================================
# KONFIGURASI HALAMAN
# ==========================================
st.set_page_config(page_title="E-Smart Guardian Pro", page_icon="🛡️", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
    <style>
    .block-container { padding-top: 2rem; padding-bottom: 2rem; }
    .main { background-color: #f8fafc; font-family: 'Segoe UI', sans-serif; }
    .stButton>button { width: 100%; background-color: #3b82f6; color: white; border-radius: 8px; height: 48px; border: none; }
    .stButton>button:hover { background-color: #2563eb; }
    .stDownloadButton>button { width: 100%; background-color: #10b981; color: white; border-radius: 8px; height: 48px; border: none; }
    .stDownloadButton>button:hover { background-color: #059669; }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# INISIALISASI SESSION STATE DARI DATABASE
# ==========================================
if 'db_loaded' not in st.session_state:
    st.session_state['raw_data'] = load_df('raw_data')
    st.session_state['processed_data'] = load_df('processed_data')
    
    staff_df = load_df('staff_list')
    st.session_state['staff_list'] = staff_df['Staff'].tolist() if not staff_df.empty else []
    
    mapping_df = load_df('template_mapping')
    if not mapping_df.empty:
        st.session_state['template_mapping_dict'] = dict(zip(mapping_df['Dept'], mapping_df['Nama Staff']))
    else:
        st.session_state['template_mapping_dict'] = {}
        
    st.session_state['store_code'] = load_config('store_code', "")
    st.session_state['db_loaded'] = True

# Baca Price Zone Lokal
if 'master_harga' not in st.session_state:
    if os.path.exists("FILE PRICE ZONE.xlsx"):
        try:
            xls = pd.ExcelFile("FILE PRICE ZONE.xlsx")
            st.session_state['master_harga'] = pd.read_excel(xls, sheet_name=0, dtype=str)
            st.session_state['master_zone'] = pd.read_excel(xls, sheet_name=1, dtype=str)
        except:
            st.session_state['master_harga'] = pd.DataFrame()
            st.session_state['master_zone'] = pd.DataFrame()
    else:
        st.session_state['master_harga'] = pd.DataFrame()
        st.session_state['master_zone'] = pd.DataFrame()

# ==========================================
# SIDEBAR MENU
# ==========================================
with st.sidebar:
    st.markdown("<h1 style='text-align:center; font-size: 50px;'>🛡️</h1>", unsafe_allow_html=True)
    st.markdown("### E-Smart Guardian")
    st.caption("Aman dari Refresh & Mati Lampu (Auto-Save DB)")
    st.markdown("---")
    
    menu_selection = st.radio(
        "📍 Navigasi Menu:",
        ["1. Setup & Upload Data", "2. Analisis & Unduhan", "3. Portal Staff (Scanner)", "4. Dashboard SL & MOD"]
    )
    st.markdown("---")
    if not st.session_state['master_harga'].empty:
        st.success("✅ File Price Zone Aktif")
    else:
        st.error("❌ FILE PRICE ZONE.xlsx Tidak Ditemukan!")
        
    if st.button("🗑️ Reset Semua Data (Hapus Database)"):
        if os.path.exists(DB_NAME): os.remove(DB_NAME)
        st.session_state.clear()
        st.rerun()

# ==========================================
# HALAMAN 1: SETUP & UPLOAD
# ==========================================
if menu_selection == "1. Setup & Upload Data":
    st.title("📂 1. Setup & Upload Laporan")
    
    with st.expander("⚙ Pengaturan Daftar Staf & Pembagian Dept (Wajib)", expanded=True):
        col1, col2 = st.columns(2)
        
        # TOMBOL DOWNLOAD TEMPLATE DIKEMBALIKAN
        with col1:
            st.caption("1. Download Template Kosong:")
            def generate_template():
                output = BytesIO()
                all_departments = [
                    'BABY', 'BATH', 'BEVERAGE', 'COTTON', 'DENTAL', 'DEO & FRAGRANCE', 'DERMA', 'EYE', 'FACE', 
                    'FOOD', 'HAIR', 'HAND & BODY', 'HOME HEALTHCARE', 'JAPAN & KOREA', 'KIDS', 'LIP', 
                    'MASS SKIN CARE', 'MEN PERSONAL CARE', 'MEN SKINCARE', 'NAIL', 'OTC EXTERNAL', 'OTC INTERNAL', 
                    'OTHERS', 'PAPER & CLEANING', 'PHARMACY', 'SANITARY PROTECTION', 'VITAMIN & SUPPLEMENT', 'WOMAN SHAVING'
                ]
                df_tpl = pd.DataFrame({"Dept": all_departments, "Nama Staff": ["" for _ in range(len(all_departments))]})
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_tpl.to_excel(writer, index=False, sheet_name='Template_Mapping')
                return output.getvalue()
                
            st.download_button("📥 Download Template Staff", data=generate_template(), file_name="Template_Dept_Staff.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            
        with col2:
            st.caption("2. Upload Template Staf Terisi:")
            uploaded_template = st.file_uploader("Upload Template (.xlsx)", type=["xlsx"])
            if uploaded_template:
                df_mapping = pd.read_excel(uploaded_template)
                df_clean = df_mapping.dropna(subset=['Dept', 'Nama Staff'])
                st.session_state['staff_list'] = df_clean['Nama Staff'].astype(str).str.strip().unique().tolist()
                st.session_state['template_mapping_dict'] = dict(zip(df_clean['Dept'].astype(str).str.strip(), df_clean['Nama Staff'].astype(str).str.strip()))
                
                # Simpan ke Database
                save_df('staff_list', pd.DataFrame({'Staff': st.session_state['staff_list']}))
                save_df('template_mapping', df_clean)
                st.success("✅ Daftar Staf berhasil dimuat dan tersimpan di Database!")
                
    with st.container(border=True):
        st.markdown("#### 📁 Unggah Laporan Pusat (ED & Markdown)")
        uploaded_files = st.file_uploader("Upload file laporan dari pusat", type=["csv", "xlsx", "xls"], accept_multiple_files=True)
        if st.button("Simpan Laporan & Lanjut"):
            if uploaded_files:
                all_data = []
                for file in uploaded_files:
                    if file.name.endswith('.csv'): df_temp = pd.read_csv(file, sep=';', header=1, on_bad_lines='skip', dtype=str)
                    else: df_temp = pd.read_excel(file, header=1, dtype=str)
                    all_data.append(df_temp)
                
                combined_df = pd.concat(all_data, ignore_index=True)
                combined_df.columns = [str(c).strip() for c in combined_df.columns]
                st.session_state['raw_data'] = combined_df
                
                # Simpan ke Database
                save_df('raw_data', combined_df)
                st.success("✅ Data Laporan Tersimpan di Database! Lanjut ke Menu 2 di kiri.")
            else:
                st.warning("Harap unggah File Laporan.")

# ==========================================
# HALAMAN 2: ANALISIS & UNDUHAN LAPORAN
# ==========================================
elif menu_selection == "2. Analisis & Unduhan":
    st.title("⚙️ 2. Konfigurasi Toko & Unduhan")
    
    if st.session_state['raw_data'].empty or not st.session_state['staff_list']:
        st.warning("⚠️ Selesaikan Setup di Menu 1 terlebih dahulu.")
    else:
        store_code_input = st.text_input("Masukkan 4 Digit Kode Toko (Contoh: 6809):", value=st.session_state['store_code'])
        if store_code_input:
            st.session_state['store_code'] = store_code_input
            save_config('store_code', store_code_input) # Simpan kode toko ke DB
            
            raw_df = st.session_state['raw_data']
            
            code_col = next((c for c in raw_df.columns if 'code' in str(c).lower() or 'site' in str(c).lower()), None)
            filtered_df = raw_df[raw_df[code_col].astype(str).str.contains(store_code_input, na=False)].copy() if code_col else raw_df.copy()
            
            st.markdown("#### 👥 Mapping Dept ke Staf")
            dept_col = next((c for c in raw_df.columns if 'dept' in str(c).lower()), None)
            
            if dept_col:
                unique_depts = filtered_df[dept_col].dropna().unique()
                mapping_input = {}
                mapping_tpl = st.session_state.get('template_mapping_dict', {})
                cols = st.columns(4)
                for i, dept in enumerate(unique_depts):
                    d_clean = str(dept).strip()
                    def_idx = st.session_state['staff_list'].index(mapping_tpl[d_clean]) if d_clean in mapping_tpl and mapping_tpl[d_clean] in st.session_state['staff_list'] else 0
                    with cols[i % 4]: mapping_input[dept] = st.selectbox(f"{dept}", st.session_state['staff_list'], index=def_idx)

                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("🚀 Proses Logika Price Zone & Diskon Markdown"):
                    # 1. Staff Mapping
                    filtered_df['Staff_Penanggung_Jawab'] = filtered_df[dept_col].map(mapping_input).fillna("-")
                    filtered_df['Status_Pengerjaan'] = "Belum"
                    filtered_df['Feedback_Staff'] = "-"

                    # 2. Ekstrak Besaran Diskon dari Kolom Remark
                    remark_col = next((c for c in filtered_df.columns if 'remark' in str(c).lower()), None)
                    def get_discount(remark):
                        r = str(remark).lower()
                        if '90' in r: return 90
                        if '75' in r: return 75
                        if '50' in r: return 50
                        return 0
                    
                    if remark_col and remark_col in filtered_df.columns:
                        filtered_df['Diskon_Persen'] = filtered_df[remark_col].apply(get_discount)
                    else:
                        filtered_df['Diskon_Persen'] = 0

                    # 3. Logika Price Zone
                    df_zone = st.session_state['master_zone']
                    df_harga = st.session_state['master_harga']
                    
                    current_zone = "Zone_1"
                    if not df_zone.empty:
                        zone_row = df_zone[df_zone.astype(str).apply(lambda x: x.str.contains(store_code_input, case=False)).any(axis=1)]
                        if not zone_row.empty:
                            for col in df_zone.columns:
                                if 'zone' in str(col).lower() or 'price' in str(col).lower():
                                    current_zone = str(zone_row.iloc[0][col])
                    
                    art_col_rep = next((c for c in filtered_df.columns if 'article' in str(c).lower()), None)
                    
                    # Pastikan kolom Harga_Normal dan Diskon_Persen ada untuk menghindari KeyError
                    if 'Harga_Normal' not in filtered_df.columns:
                        filtered_df['Harga_Normal'] = 0
                        
                    if not df_harga.empty and art_col_rep:
                        art_col_har = next((c for c in df_harga.columns if 'article' in str(c).lower() or 'plu' in str(c).lower()), None)
                        zone_price_col = next((c for c in df_harga.columns if current_zone.lower() in str(c).lower() or 'price' in str(c).lower()), None)
                        if art_col_har and zone_price_col:
                            harga_dict = dict(zip(df_harga[art_col_har].astype(str).str.strip(), df_harga[zone_price_col]))
                            filtered_df['Harga_Normal'] = filtered_df[art_col_rep].astype(str).str.strip().map(harga_dict).fillna(0)

                    # 4. Kalkulasi Harga Akhir Label (Perbaikan KeyError)
                    filtered_df['Harga_Normal'] = pd.to_numeric(filtered_df['Harga_Normal'], errors='coerce').fillna(0)
                    filtered_df['Diskon_Persen'] = pd.to_numeric(filtered_df['Diskon_Persen'], errors='coerce').fillna(0)
                    filtered_df['Harga_Akhir'] = filtered_df['Harga_Normal'] * (1 - (filtered_df['Diskon_Persen'] / 100))
                    
                    st.session_state['processed_data'] = filtered_df
                    save_df('processed_data', filtered_df) # Simpan proses ke DB
                    
                    st.success("✅ Price Zone & Diskon Berhasil Dikalkulasi!")

            if not st.session_state['processed_data'].empty:
                st.markdown("---")
                st.markdown("### 🖨️ Area Unduhan Staf & Label")
                final_df = st.session_state['processed_data']
                
                c_dl1, c_dl2 = st.columns(2)
                
                with c_dl1:
                    st.markdown("#### 1. Laporan Tugas Staf (Bersih)")
                    def get_clean_col(df, keywords):
                        for c in df.columns:
                            cl = str(c).lower().replace(' ', '')
                            if any(k in cl for k in keywords): return c
                        return None
                    
                    esensial_cols = []
                    for kw in [['code'], ['site', 'desc'], ['article'], ['article', 'desc'], ['dept'], ['month'], ['year'], ['qty', 'exp'], ['remark'], ['staff'], ['status'], ['feedback']]:
                        found = get_clean_col(final_df, kw)
                        if found and found not in esensial_cols: esensial_cols.append(found)
                    
                    clean_df = final_df[esensial_cols]
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as w: clean_df.to_excel(w, index=False)
                    
                    st.download_button(
                        label="📥 Download Laporan Tugas (.xlsx)",
                        data=output.getvalue(),
                        file_name=f"Laporan_Tugas_{store_code_input}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                with c_dl2:
                    st.markdown("#### 2. Label Clearance Siap Cetak")
                    st.caption("File berformat HTML. Buka di Browser lalu tekan **Ctrl+P** untuk mencetak dengan layout 1.5 x 3.")
                    
                    if 'Diskon_Persen' in final_df.columns:
                        markdown_df = final_df[pd.to_numeric(final_df['Diskon_Persen'], errors='coerce') > 0].copy()
                        
                        if not markdown_df.empty:
                            html_content = """
                            <html><head><style>
                            body { margin: 0; padding: 0; font-family: 'Arial', sans-serif; background: #fff;}
                            .page { width: 210mm; margin: 0 auto; padding: 2mm; display: flex; flex-wrap: wrap; }
                            .label { 
                                width: 32mm; height: 16mm;
                                border: 1px dashed #999; margin: 1mm; padding: 1.5mm; 
                                box-sizing: border-box; display: flex; flex-direction: column; justify-content: space-between;
                            }
                            .top { display: flex; justify-content: space-between; align-items: flex-start; }
                            .disc { font-size: 7pt; line-height: 1; font-weight: bold; }
                            .price { font-size: 11pt; font-weight: 900; letter-spacing: -0.5px;}
                            .bot { display: flex; justify-content: space-between; font-size: 6pt; font-weight: bold; margin-top: 1px;}
                            .strike { text-decoration: line-through; font-weight: normal; }
                            .art { font-size: 5pt; text-align: left; margin-top: -2px; }
                            @media print { .label { border: 1px dashed #fff; } .page { padding: 0; } }
                            </style></head><body><div class="page">
                            """
                            for _, row in markdown_df.iterrows():
                                disc = int(row['Diskon_Persen'])
                                h_norm = f"{int(row['Harga_Normal']):,}".replace(',', '.')
                                h_akhir = f"{int(row['Harga_Akhir']):,}".replace(',', '.')
                                art_val = str(row.get(get_clean_col(markdown_df, ['article']), ''))
                                
                                html_content += f"""
                                <div class="label">
                                    <div class="top">
                                        <div class="disc">Disc.<br>{disc}%</div>
                                        <div class="price">{h_akhir}</div>
                                    </div>
                                    <div class="bot">
                                        <span>{store_code_input}</span>
                                        <span class="strike">{h_norm}</span>
                                    </div>
                                    <div class="art">{art_val}</div>
                                </div>
                                """
                            html_content += "</div></body></html>"
                            
                            st.download_button(
                                label="🖨️ Download Label HTML (Siap Print)",
                                data=html_content,
                                file_name=f"Label_Clearance_{store_code_input}.html",
                                mime="text/html"
                            )
                        else:
                            st.info("Tidak ada item berstatus Markdown 50/75/90%.")

# ==========================================
# HALAMAN 3: PORTAL STAFF (SCANNER)
# ==========================================
elif menu_selection == "3. Portal Staff (Scanner)":
    st.title("📱 3. Portal Kerja Staf")
    
    if st.session_state['processed_data'].empty:
        st.warning("Belum ada data tugas diproses.")
    else:
        df_current = st.session_state['processed_data']
        
        c_staff, c_scan = st.columns([1, 2])
        with c_staff: aktif_staff = st.selectbox("Nama Staf:", st.session_state['staff_list'])
        with c_scan: scan_input = st.text_input("Scan Barcode / Ketik Nomor Artikel:")
            
        staff_note = st.text_input("Catatan / Feedback Kerja (Opsional):")
        
        if st.button("Kirim Laporan Selesai"):
            if scan_input:
                match_idx = df_current[df_current.astype(str).apply(lambda x: x.str.contains(scan_input, case=False)).any(axis=1)].index
                if not match_idx.empty:
                    df_current.loc[match_idx, 'Status_Pengerjaan'] = "Selesai"
                    df_current.loc[match_idx, 'Feedback_Staff'] = staff_note if staff_note else f"Selesai di-display"
                    
                    st.session_state['processed_data'] = df_current
                    save_df('processed_data', df_current) # UPDATE DATABASE SAAT STAF SCAN BARANG
                    
                    st.success(f"✅ Artikel {scan_input} diselesaikan oleh {aktif_staff}!")
                else:
                    st.error("❌ Artikel tidak ditemukan.")

# ==========================================
# HALAMAN 4: DASHBOARD SL & MOD
# ==========================================
elif menu_selection == "4. Dashboard SL & MOD":
    st.title("📊 4. Dashboard Monitoring SL")
    
    if 'auth_mod' not in st.session_state: st.session_state['auth_mod'] = False
        
    if not st.session_state['auth_mod']:
        st.warning("🔒 Halaman Terkunci. Masukkan PIN SL/MOD.")
        pin_input = st.text_input("PIN Akses:", type="password")
        if st.button("Login"):
            valid_pins = ["99231078", "99230454", "99250259"]
            if pin_input in valid_pins: 
                st.session_state['auth_mod'] = True
                st.rerun()
            else:
                st.error("PIN Salah atau Tidak Terdaftar!")
    else:
        if st.button("🔒 Logout"):
            st.session_state['auth_mod'] = False
            st.rerun()
            
        if st.session_state['processed_data'].empty:
            st.info("Belum ada data.")
        else:
            df_dash = st.session_state['processed_data']
            total = len(df_dash)
            selesai = len(df_dash[df_dash['Status_Pengerjaan'] == "Selesai"])
            
            c1, c2, c3 = st.columns(3)
            c1.metric("Total Tugas Item", total)
            c2.metric("Selesai Dikerjakan", selesai)
            c3.metric("Belum Dikerjakan", total - selesai)
            
            st.markdown("#### Detail Tabel Progress")
            
            art_col = next((c for c in df_dash.columns if 'article' in str(c).lower()), 'Article')
            desc_col = next((c for c in df_dash.columns if 'desc' in str(c).lower() and 'site' not in str(c).lower()), 'Description')
            view_cols = [c for c in [art_col, desc_col, 'Staff_Penanggung_Jawab', 'Status_Pengerjaan', 'Feedback_Staff'] if c in df_dash.columns]
            
            st.dataframe(df_dash[view_cols], use_container_width=True)
            
            out = BytesIO()
            with pd.ExcelWriter(out, engine='openpyxl') as w: df_dash.to_excel(w, index=False)
            st.download_button(
                "📥 Download Rekap Akhir Hasil Pengerjaan", 
                data=out.getvalue(), 
                file_name=f"Monitoring_Akhir_{st.session_state.get('store_code', 'Toko')}.xlsx", 
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
