import pandas as pd
import streamlit as st
from io import BytesIO
import os
import re

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
# INISIALISASI & BACA PRICE ZONE LOKAL
# ==========================================
if 'staff_list' not in st.session_state: st.session_state['staff_list'] = []
if 'store_code' not in st.session_state: st.session_state['store_code'] = ""
if 'template_mapping_dict' not in st.session_state: st.session_state['template_mapping_dict'] = {}
if 'raw_data' not in st.session_state: st.session_state['raw_data'] = pd.DataFrame()
if 'processed_data' not in st.session_state: st.session_state['processed_data'] = pd.DataFrame()

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
    st.caption("Alur: Setup ➡️ Analisis ➡️ Staff Scan ➡️ Dashboard SL")
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

# ==========================================
# HALAMAN 1: SETUP & UPLOAD
# ==========================================
if menu_selection == "1. Setup & Upload Data":
    st.title("📂 1. Setup & Upload Laporan")
    
    with st.expander("⚙ Pengaturan Daftar Staf & Pembagian Dept (Wajib)", expanded=True):
        col1, col2 = st.columns(2)
        with col1:
            st.caption("Upload template staf di sini:")
            uploaded_template = st.file_uploader("Upload Template Staf (.xlsx)", type=["xlsx"])
            if uploaded_template:
                df_mapping = pd.read_excel(uploaded_template)
                df_clean = df_mapping.dropna(subset=['Dept', 'Nama Staff'])
                st.session_state['staff_list'] = df_clean['Nama Staff'].astype(str).str.strip().unique().tolist()
                st.session_state['template_mapping_dict'] = dict(zip(df_clean['Dept'].astype(str).str.strip(), df_clean['Nama Staff'].astype(str).str.strip()))
                st.success("✅ Daftar Staf dimuat!")
                
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
                st.success("✅ Data Laporan Tersimpan! Lanjut ke Menu 2 di kiri.")
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
                    filtered_df['Diskon_Persen'] = filtered_df[remark_col].apply(get_discount) if remark_col else 0

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
                    if not df_harga.empty and art_col_rep:
                        art_col_har = next((c for c in df_harga.columns if 'article' in str(c).lower() or 'plu' in str(c).lower()), None)
                        zone_price_col = next((c for c in df_harga.columns if current_zone.lower() in str(c).lower() or 'price' in str(c).lower()), None)
                        if art_col_har and zone_price_col:
                            harga_dict = dict(zip(df_harga[art_col_har].astype(str).str.strip(), df_harga[zone_price_col]))
                            filtered_df['Harga_Normal'] = pd.to_numeric(filtered_df[art_col_rep].astype(str).str.strip().map(harga_dict).fillna(0), errors='coerce')
                    else:
                        filtered_df['Harga_Normal'] = 0

                    # 4. Kalkulasi Harga Akhir Label
                    filtered_df['Harga_Akhir'] = filtered_df['Harga_Normal'] * (1 - (filtered_df['Diskon_Persen'] / 100))
                    
                    st.session_state['processed_data'] = filtered_df
                    st.success("✅ Price Zone & Diskon Berhasil Dikalkulasi!")

            if not st.session_state['processed_data'].empty:
                st.markdown("---")
                st.markdown("### 🖨️ Area Unduhan Staf & Label")
                final_df = st.session_state['processed_data']
                
                c_dl1, c_dl2 = st.columns(2)
                
                # UNDUHAN 1: LAPORAN BERSIH (Hanya kolom yang diminta)
                with c_dl1:
                    st.markdown("#### 1. Laporan Tugas Staf (Bersih)")
                    def get_clean_col(df, keywords):
                        for c in df.columns:
                            cl = str(c).lower().replace(' ', '')
                            if any(k in cl for k in keywords): return c
                        return None
                    
                    # Pemetaan kolom esensial
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

                # UNDUHAN 2: HTML LABEL CLEARANCE (Persis seperti gambar 1,5 x 3)
                with c_dl2:
                    st.markdown("#### 2. Label Clearance Siap Cetak")
                    st.caption("File berformat HTML. Buka di Browser (Chrome) lalu tekan **Ctrl+P** untuk mencetak dengan layout pas.")
                    
                    markdown_df = final_df[final_df['Diskon_Persen'] > 0].copy()
                    
                    if not markdown_df.empty:
                        # Buat Template HTML untuk Label
                        html_content = """
                        <html><head><style>
                        body { margin: 0; padding: 0; font-family: 'Arial', sans-serif; background: #fff;}
                        .page { width: 210mm; margin: 0 auto; padding: 2mm; display: flex; flex-wrap: wrap; }
                        .label { 
                            width: 32mm; height: 16mm; /* Ukuran mendekati 1.5x3 */
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
                    st.success(f"✅ Artikel {scan_input} diselesaikan oleh {aktif_staff}!")
                else:
                    st.error("❌ Artikel tidak ditemukan.")

# ==========================================
# HALAMAN 4: DASHBOARD SL & MOD (PIN KHUSUS)
# ==========================================
elif menu_selection == "4. Dashboard SL & MOD":
    st.title("📊 4. Dashboard Monitoring SL")
    
    if 'auth_mod' not in st.session_state: st.session_state['auth_mod'] = False
        
    if not st.session_state['auth_mod']:
        st.warning("🔒 Halaman Terkunci. Masukkan PIN SL/MOD.")
        pin_input = st.text_input("PIN Akses:", type="password")
        if st.button("Login"):
            # 3 PIN Otorisasi yang diminta
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
            
            # Kolom untuk SL view
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
