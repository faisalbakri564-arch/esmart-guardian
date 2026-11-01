import pandas as pd
import streamlit as st
from io import BytesIO
import os

# ==========================================
# KONFIGURASI HALAMAN & CUSTOM CSS
# ==========================================
st.set_page_config(
    page_title="E-Smart Guardian Pro",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
    <style>
    .block-container { padding-top: 2rem; padding-bottom: 2rem; }
    .main { background-color: #f8fafc; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
    h1, h2, h3, h4 { color: #1e293b; font-weight: 700; }
    .stButton>button { 
        width: 100%; background-color: #3b82f6; color: white; font-weight: 600; 
        border-radius: 8px; height: 48px; border: none; transition: all 0.3s ease;
    }
    .stButton>button:hover { background-color: #2563eb; transform: translateY(-2px); box-shadow: 0 4px 12px rgba(59, 130, 246, 0.3); }
    .stDownloadButton>button { 
        width: 100%; background-color: #10b981; color: white; font-weight: 600; 
        border-radius: 8px; height: 48px; border: none; transition: all 0.3s ease;
    }
    .stDownloadButton>button:hover { background-color: #059669; transform: translateY(-2px); box-shadow: 0 4px 12px rgba(16, 185, 129, 0.3); }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# INISIALISASI SESSION STATE & MEMBACA PRICE ZONE LOKAL
# ==========================================
if 'staff_list' not in st.session_state: st.session_state['staff_list'] = []
if 'store_code' not in st.session_state: st.session_state['store_code'] = ""
if 'template_mapping_dict' not in st.session_state: st.session_state['template_mapping_dict'] = {}
if 'raw_data' not in st.session_state: st.session_state['raw_data'] = pd.DataFrame()
if 'processed_data' not in st.session_state: st.session_state['processed_data'] = pd.DataFrame()

# Baca file FILE PRICE ZONE.xlsx secara otomatis (Lokal)
if 'master_harga' not in st.session_state:
    if os.path.exists("FILE PRICE ZONE.xlsx"):
        try:
            xls = pd.ExcelFile("FILE PRICE ZONE.xlsx")
            st.session_state['master_harga'] = pd.read_excel(xls, sheet_name=0, dtype=str)
            st.session_state['master_zone'] = pd.read_excel(xls, sheet_name=1, dtype=str)
        except Exception as e:
            st.sidebar.error(f"Error membaca Price Zone: {e}")
            st.session_state['master_harga'] = pd.DataFrame()
            st.session_state['master_zone'] = pd.DataFrame()
    else:
        st.session_state['master_harga'] = pd.DataFrame()
        st.session_state['master_zone'] = pd.DataFrame()

# ==========================================
# SIDEBAR MENU (NAVIGASI KIRI)
# ==========================================
with st.sidebar:
    if os.path.exists("logo.png"):
        st.image("logo.png", width=180)
    elif os.path.exists("logo.jpg"):
        st.image("logo.jpg", width=180)
    else:
        st.markdown("<h1 style='font-size: 50px;'>🛡️</h1>", unsafe_allow_html=True)
        
    st.markdown("### E-Smart Guardian")
    st.caption("Manajemen ED, Price Zone & Monitoring Staf.")
    st.markdown("---")
    
    menu_selection = st.radio(
        "📍 Navigasi Menu:",
        ["1. Setup & Upload Data", "2. Analisis & Unduhan", "3. Portal Staff (Scanner)", "4. Dashboard SL & MOD"]
    )
    
    st.markdown("---")
    st.caption("Status Sistem:")
    if not st.session_state['master_harga'].empty:
        st.success("✅ Master Price Zone Terhubung")
    else:
        st.error("❌ FILE PRICE ZONE.xlsx Tidak Ditemukan!")

# ==========================================
# HALAMAN 1: SETUP & UPLOAD DATA
# ==========================================
if menu_selection == "1. Setup & Upload Data":
    st.title("📂 1. Setup & Upload Laporan")
    
    with st.expander("⚙ Pengaturan Daftar Staf & Pembagian Dept (Wajib)", expanded=True):
        st.info("Upload template pembagian departemen staf terlebih dahulu.")
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
            
        col_tpl1, col_tpl2 = st.columns(2)
        with col_tpl1:
            st.download_button("📥 Download Template Dept", data=generate_template(), file_name="Template_Dept_Staff.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with col_tpl2:
            uploaded_template = st.file_uploader("Upload Template Terisi (.xlsx)", type=["xlsx", "xls"])
            if uploaded_template:
                try:
                    df_mapping = pd.read_excel(uploaded_template)
                    df_clean = df_mapping.dropna(subset=['Dept', 'Nama Staff'])
                    df_clean = df_clean[df_clean['Nama Staff'].astype(str).str.strip() != '']
                    st.session_state['staff_list'] = df_clean['Nama Staff'].astype(str).str.strip().unique().tolist()
                    st.session_state['template_mapping_dict'] = dict(zip(df_clean['Dept'].astype(str).str.strip(), df_clean['Nama Staff'].astype(str).str.strip()))
                    st.success("✅ Daftar Staf berhasil dimuat!")
                except Exception as e:
                    st.error(f"Error format: {e}")

    with st.container(border=True):
        st.markdown("#### 📁 Unggah Laporan Pusat (ED & Markdown Pak Yanyan)")
        uploaded_files = st.file_uploader("Upload file laporan (Bisa multi-file, Maks 5)", type=["csv", "xlsx", "xls", "xlsb"], accept_multiple_files=True)
        
        if st.button("Simpan Laporan & Lanjut"):
            if uploaded_files:
                all_data = []
                for file in uploaded_files:
                    if file.name.endswith('.csv'):
                        df_temp = pd.read_csv(file, sep=';', header=1, on_bad_lines='skip', dtype=str)
                    else:
                        df_temp = pd.read_excel(file, header=1, dtype=str)
                    all_data.append(df_temp)
                
                combined_df = pd.concat(all_data, ignore_index=True)
                combined_df.columns = [str(c).strip() for c in combined_df.columns]
                combined_df = combined_df.loc[:, ~combined_df.columns.str.contains('^Unnamed|^None', case=False, na=False)]
                
                st.session_state['raw_data'] = combined_df
                st.success("✅ Data Laporan Tersimpan! Silakan klik menu '2. Analisis & Unduhan' di sebelah kiri.")
            else:
                st.warning("Harap unggah File Laporan terlebih dahulu.")

# ==========================================
# HALAMAN 2: ANALISIS & UNDUHAN
# ==========================================
elif menu_selection == "2. Analisis & Unduhan":
    st.title("⚙️ 2. Konfigurasi Toko & Unduhan Laporan")
    
    if st.session_state['raw_data'].empty:
        st.warning("⚠️ Data laporan belum diunggah. Silakan ke menu '1. Setup & Upload Data' terlebih dahulu.")
    elif not st.session_state['staff_list']:
        st.warning("⚠️ Template staf belum diunggah. Silakan kembali ke menu '1. Setup & Upload Data'.")
    else:
        store_code_input = st.text_input("Masukkan 4 Digit Kode Toko Anda (Contoh: 6809):", value=st.session_state['store_code'])
        
        if store_code_input:
            st.session_state['store_code'] = store_code_input
            raw_df = st.session_state['raw_data']
            
            # Filter Toko
            code_col = next((col for col in raw_df.columns if 'code' in str(col).lower() or 'site' in str(col).lower()), None)
            if code_col:
                filtered_df = raw_df[raw_df[code_col].astype(str).str.contains(store_code_input, na=False)].copy()
            else:
                filtered_df = raw_df.copy()
            
            st.markdown("#### 👥 Mapping Dept ke Staf Toko (Cek Kembali)")
            dept_col = next((col for col in raw_df.columns if str(col).upper() == 'DEPT' or 'dept' in str(col).lower()), None)
            
            if dept_col:
                unique_depts = filtered_df[dept_col].dropna().unique()
                mapping_input = {}
                mapping_tpl = st.session_state.get('template_mapping_dict', {})
                
                cols = st.columns(4)
                for i, dept in enumerate(unique_depts):
                    d_clean = str(dept).strip()
                    def_idx = 0
                    if d_clean in mapping_tpl and mapping_tpl[d_clean] in st.session_state['staff_list']:
                        def_idx = st.session_state['staff_list'].index(mapping_tpl[d_clean])
                    with cols[i % 4]:
                        mapping_input[dept] = st.selectbox(f"{dept}", st.session_state['staff_list'], index=def_idx, key=f"map_{dept}")

                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("🚀 Terapkan Penugasan, AI, & Harga"):
                    with st.spinner("Memproses sinkronisasi data..."):
                        # 1. Terapkan Staff
                        filtered_df['Staff_Penanggung_Jawab'] = filtered_df[dept_col].map(mapping_input).fillna("Belum Ditugaskan")
                        
                        # 2. Terapkan AI Instruksi Aksi
                        remark_col = next((col for col in filtered_df.columns if 'remark' in str(col).lower() or 'instruction' in str(col).lower()), None)
                        def smart_ai(row):
                            r_text = str(row[remark_col]).lower() if remark_col else ""
                            if 'expired' in r_text or 'lewat' in r_text: return "Ajukan WO Retur."
                            elif 'blue dot' in r_text: return "Pisahkan untuk GWP/Retur."
                            elif any(kw in r_text for kw in ['markdown', 'rtw', 'return', 'disetujui']): return str(row[remark_col])
                            else: return "Cek Fisik & Sesuaikan Display."
                        filtered_df['Instruksi_Aksi_AI'] = filtered_df.apply(smart_ai, axis=1)

                        # 3. Pencocokan Harga & Kalkulasi Diskon
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
                                filtered_df['Harga_Normal'] = filtered_df[art_col_rep].astype(str).str.strip().map(harga_dict).fillna(0)
                            else:
                                filtered_df['Harga_Normal'] = 0
                        else:
                            filtered_df['Harga_Normal'] = 0

                        try:
                            h_float = pd.to_numeric(filtered_df['Harga_Normal'], errors='coerce').fillna(0)
                            filtered_df['Diskon_50'] = h_float * 0.50
                            filtered_df['Diskon_75'] = h_float * 0.25
                            filtered_df['Diskon_90'] = h_float * 0.10
                        except:
                            filtered_df['Diskon_50'] = 0
                            filtered_df['Diskon_75'] = 0
                            filtered_df['Diskon_90'] = 0

                        # 4. Status Staff
                        filtered_df['Status_Pengerjaan'] = "Belum"
                        filtered_df['Feedback_Staff'] = "-"
                        filtered_df['Scan_Waktu'] = "-"

                        st.session_state['processed_data'] = filtered_df
                        st.success("✅ Proses Selesai! Silakan unduh file untuk dibagikan ke staf.")

            # AREA UNDUHAN (Muncul setelah diproses)
            if not st.session_state['processed_data'].empty:
                st.markdown("---")
                st.markdown("### 🖨️ Pusat Unduhan")
                
                final_df = st.session_state['processed_data']
                
                def convert_df_to_excel(df_in):
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_in.to_excel(writer, index=False, sheet_name='Laporan')
                    return output.getvalue()

                col_dl1, col_dl2 = st.columns(2)
                
                # DOWNLOAD 1: LAPORAN UTAMA (Format Lama + Kolom Tambahan)
                with col_dl1:
                    with st.container(border=True):
                        st.markdown("#### 1. Laporan Tugas Staf")
                        st.caption("File Excel dengan format lama utuh, ditambah kolom Penugasan Staf & AI di bagian akhir.")
                        # Sembunyikan kolom harga agar tidak membingungkan di file utama (opsional, tapi kita sertakan instruksi)
                        cols_to_hide_in_main = ['Harga_Normal', 'Diskon_50', 'Diskon_75', 'Diskon_90']
                        main_dl_df = final_df.drop(columns=[c for c in cols_to_hide_in_main if c in final_df.columns])
                        
                        st.download_button(
                            label="📥 Download Laporan Tugas Staf (.xlsx)",
                            data=convert_df_to_excel(main_dl_df),
                            file_name=f"Laporan_Tugas_{store_code_input}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                
                # DOWNLOAD 2: CLEARANCE LABEL
                with col_dl2:
                    with st.container(border=True):
                        st.markdown("#### 2. Label Clearance (Siap Cetak)")
                        st.caption("Format label 1.5x3, berisi Harga Terpotong (50%, 75%, 90%) dan Kode Toko.")
                        
                        # Filter hanya yang markdown
                        markdown_df = final_df[final_df.astype(str).apply(lambda x: x.str.contains("Markdown", case=False)).any(axis=1)].copy()
                        
                        if not markdown_df.empty:
                            label_df = pd.DataFrame()
                            label_df['No'] = range(1, len(markdown_df) + 1)
                            
                            art_c = next((c for c in markdown_df.columns if 'article' in str(c).lower()), 'Article')
                            desc_c = next((c for c in markdown_df.columns if 'desc' in str(c).lower()), 'Description')
                            
                            label_df['Artikel'] = markdown_df[art_c] if art_c in markdown_df.columns else "-"
                            label_df['Nama Barang'] = markdown_df[desc_c] if desc_c in markdown_df.columns else "-"
                            label_df['Harga Normal'] = markdown_df['Harga_Normal']
                            label_df['Diskon 50%'] = markdown_df['Diskon_50']
                            label_df['Diskon 75%'] = markdown_df['Diskon_75']
                            label_df['Diskon 90%'] = markdown_df['Diskon_90']
                            label_df['Toko Pemilik'] = store_code_input
                            label_df['Toko Titip Jual'] = store_code_input
                            
                            st.download_button(
                                label="📥 Download Label Clearance (.xlsx)",
                                data=convert_df_to_excel(label_df),
                                file_name=f"Label_Clearance_{store_code_input}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                        else:
                            st.info("Tidak ada item berstatus Markdown.")

# ==========================================
# HALAMAN 3: PORTAL STAFF (SCANNER)
# ==========================================
elif menu_selection == "3. Portal Staff (Scanner)":
    st.title("📱 3. Portal Kerja Staf")
    
    if st.session_state['processed_data'].empty:
        st.warning("Belum ada data tugas. SL/MOD harus memproses Laporan di Menu 2 terlebih dahulu.")
    else:
        st.info("Pilih Nama Anda, lalu Scan Barcode/Artikel yang sudah selesai dikerjakan untuk melapor ke SL/MOD.")
        
        df_current = st.session_state['processed_data']
        
        c_staff, c_scan = st.columns([1, 2])
        with c_staff:
            aktif_staff = st.selectbox("Nama Staf:", st.session_state['staff_list'])
        with c_scan:
            scan_input = st.text_input("Scan Barcode / Ketik Nomor Artikel:", placeholder="Klik disini lalu tekan scanner...")
            
        staff_note = st.text_input("Catatan / Feedback Kerja (Opsional):", placeholder="Contoh: Barang sudah dipindah ke rak depan...")
        
        if st.button("Kirim Laporan Selesai"):
            if scan_input:
                match_idx = df_current[df_current.astype(str).apply(lambda x: x.str.contains(scan_input, case=False)).any(axis=1)].index
                if not match_idx.empty:
                    df_current.loc[match_idx, 'Status_Pengerjaan'] = "Selesai"
                    df_current.loc[match_idx, 'Feedback_Staff'] = staff_note if staff_note else f"Dikerjakan oleh {aktif_staff}"
                    df_current.loc[match_idx, 'Scan_Waktu'] = pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")
                    st.session_state['processed_data'] = df_current
                    st.success(f"✅ Artikel {scan_input} berhasil diselesaikan oleh {aktif_staff}!")
                else:
                    st.error("❌ Artikel tidak ditemukan dalam daftar laporan.")
            else:
                st.warning("Masukkan nomor artikel terlebih dahulu.")
                
        # Tampilkan history scan staf ini
        st.markdown("#### Histori Pekerjaan Anda Hari Ini:")
        my_tasks = df_current[(df_current['Staff_Penanggung_Jawab'] == aktif_staff) & (df_current['Status_Pengerjaan'] == 'Selesai')]
        art_c = next((c for c in my_tasks.columns if 'article' in str(c).lower()), 'Article')
        desc_c = next((c for c in my_tasks.columns if 'desc' in str(c).lower()), 'Description')
        
        if not my_tasks.empty and art_c in my_tasks.columns:
            st.dataframe(my_tasks[[art_c, desc_c, 'Feedback_Staff', 'Scan_Waktu']], use_container_width=True)
        else:
            st.caption("Belum ada barang yang diselesaikan.")

# ==========================================
# HALAMAN 4: DASHBOARD SL & MOD (OTORISASI)
# ==========================================
elif menu_selection == "4. Dashboard SL & MOD":
    st.title("📊 4. Dashboard Monitoring Pimpinan")
    
    # SISTEM OTORISASI (PIN)
    if 'auth_mod' not in st.session_state:
        st.session_state['auth_mod'] = False
        
    if not st.session_state['auth_mod']:
        st.warning("🔒 Halaman Terkunci. Khusus Store Leader / MOD.")
        pin_input = st.text_input("Masukkan PIN Akses:", type="password")
        if st.button("Login"):
            # Anda bisa mengganti '123456' dengan PIN yang Anda inginkan
            if pin_input == "123456": 
                st.session_state['auth_mod'] = True
                st.rerun()
            else:
                st.error("PIN Salah!")
    else:
        if st.button("🔒 Logout"):
            st.session_state['auth_mod'] = False
            st.rerun()
            
        if st.session_state['processed_data'].empty:
            st.info("Belum ada data tugas yang diproses untuk dimonitor.")
        else:
            df_dash = st.session_state['processed_data']
            
            total_tugas = len(df_dash)
            selesai = len(df_dash[df_dash['Status_Pengerjaan'] == "Selesai"])
            pending = total_tugas - selesai
            
            st.markdown("### Ringkasan Pengerjaan Toko")
            c1, c2, c3 = st.columns(3)
            c1.metric("Total Tugas Item", total_tugas)
            c2.metric("Selesai (Scanned)", selesai, delta=f"{(selesai/total_tugas*100):.1f}%" if total_tugas>0 else "0%")
            c3.metric("Belum Dikerjakan", pending, delta_color="inverse")
            
            st.markdown("---")
            st.markdown("#### Detail Tabel Monitoring Staf")
            
            art_c = next((c for c in df_dash.columns if 'article' in str(c).lower()), 'Article')
            desc_c = next((c for c in df_dash.columns if 'desc' in str(c).lower()), 'Description')
            
            view_cols = [art_c, desc_c, 'Staff_Penanggung_Jawab', 'Status_Pengerjaan', 'Feedback_Staff', 'Scan_Waktu']
            valid_view_cols = [c for c in view_cols if c in df_dash.columns]
            
            st.dataframe(df_dash[valid_view_cols], use_container_width=True)
            
            # Download Final Report
            def convert_df_to_excel(df_in):
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_in.to_excel(writer, index=False, sheet_name='Monitoring_Selesai')
                return output.getvalue()
                
            st.markdown("<br>", unsafe_allow_html=True)
            st.download_button(
                "📥 Download Rekap Akhir (Format Lama + Status Selesai)", 
                data=convert_df_to_excel(df_dash), 
                file_name=f"Monitoring_Akhir_{st.session_state.get('store_code', 'Toko')}.xlsx", 
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
