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
# INISIALISASI SESSION STATE
# ==========================================
if 'step' not in st.session_state: st.session_state['step'] = 1
if 'staff_list' not in st.session_state: st.session_state['staff_list'] = []
if 'store_code' not in st.session_state: st.session_state['store_code'] = ""
if 'store_name_dynamic' not in st.session_state: st.session_state['store_name_dynamic'] = "Toko Retail"
if 'template_mapping_dict' not in st.session_state: st.session_state['template_mapping_dict'] = {}

# ==========================================
# SIDEBAR
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
    
    if st.session_state['step'] > 1:
        if st.button("🏠 Menu Utama / Reset"):
            st.session_state['step'] = 1
            st.rerun()

# ==========================================
# HEADER & PROGRESS BAR
# ==========================================
st.title("🛡️ E-Smart Guardian Pro (Store Operation)")
st.markdown("Sistem terintegrasi untuk pengecekan ED, pencocokan *Price Zone*, cetak label *Markdown*, dan *Dashboard* SL/MOD.")

progress_value = int((st.session_state['step'] / 5) * 100)
st.progress(st.session_state['step'] / 5, text=f"Tahap {st.session_state['step']} dari 5 ({progress_value}%)")
st.markdown("<br>", unsafe_allow_html=True)

# ==========================================
# TAHAP 1: UPLOAD LAPORAN & MASTER PRICE ZONE
# ==========================================
if st.session_state['step'] == 1:
    
    with st.expander("⚙ Pengaturan Daftar Staf & Pembagian Dept", expanded=False):
        st.info("Upload template pembagian departemen staf.")
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
                    if 'Dept' in df_mapping.columns and 'Nama Staff' in df_mapping.columns:
                        df_clean = df_mapping.dropna(subset=['Dept', 'Nama Staff'])
                        df_clean = df_clean[df_clean['Nama Staff'].astype(str).str.strip() != '']
                        st.session_state['staff_list'] = df_clean['Nama Staff'].astype(str).str.strip().unique().tolist()
                        st.session_state['template_mapping_dict'] = dict(zip(df_clean['Dept'].astype(str).str.strip(), df_clean['Nama Staff'].astype(str).str.strip()))
                        st.success("✅ Staf berhasil dimuat!")
                except Exception as e:
                    st.error(f"Error: {e}")

    with st.container(border=True):
        st.markdown("#### 📁 1. Unggah Laporan Pusat (ED & Markdown)")
        uploaded_files = st.file_uploader("Upload file laporan dari e-mail (Maksimal 5 file)", type=["csv", "xlsx", "xls", "xlsb"], accept_multiple_files=True)

        st.markdown("#### 💰 2. Unggah Master Harga & Price Zone (Excel)")
        st.markdown("*(Sheet 1: Master Harga Produk berdasarkan Article/PLU | Sheet 2: Pemetaan Kode Toko ke Price Zone)*")
        price_zone_file = st.file_uploader("Upload Master Price Zone (.xlsx)", type=["xlsx", "xls"])

        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("Lanjut ke Pengaturan Toko & Staf ➡"):
            if uploaded_files and price_zone_file:
                try:
                    # Baca File Laporan Utama
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
                    
                    # Baca Master Price Zone (Sheet 1 & Sheet 2)
                    xls = pd.ExcelFile(price_zone_file)
                    sheet_names = xls.sheet_names
                    df_master_harga = pd.read_excel(xls, sheet_name=0, dtype=str) # Sheet 1: Master Harga / Article
                    df_master_zone = pd.read_excel(xls, sheet_name=1, dtype=str)   # Sheet 2: Kode Toko & Price Zone
                    
                    st.session_state['raw_data'] = combined_df
                    st.session_state['master_harga'] = df_master_harga
                    st.session_state['master_zone'] = df_master_zone
                    st.session_state['step'] = 2
                    st.rerun()
                except Exception as e:
                    st.error(f"Gagal memproses file: {e}")
            else:
                st.warning("Harap unggah File Laporan dan Master Price Zone terlebih dahulu.")

# ==========================================
# TAHAP 2: INPUT KODE TOKO & MAPPING STAFF
# ==========================================
elif st.session_state['step'] == 2:
    with st.container(border=True):
        st.markdown("#### 🏢 Konfigurasi Toko & Penanggung Jawab")
        store_code_input = st.text_input("Masukkan 4 Digit Kode Toko Anda (Contoh: 6849):", value=st.session_state['store_code'])
        
        if store_code_input:
            st.session_state['store_code'] = store_code_input
            raw_df = st.session_state['raw_data']
            
            # Filter Toko
            code_col = next((col for col in raw_df.columns if 'code' in str(col).lower() or 'site' in str(col).lower()), None)
            if code_col:
                filtered_df = raw_df[raw_df[code_col].astype(str).str.contains(store_code_input, na=False)].copy()
            else:
                filtered_df = raw_df.copy()
            
            st.session_state['filtered_data'] = filtered_df
            st.metric(label="Total Baris Data Laporan Toko", value=len(filtered_df))

            if st.session_state['staff_list']:
                st.markdown("#### 👥 Mapping Dept ke Staf Toko")
                dept_col = next((col for col in raw_df.columns if str(col).upper() == 'DEPT' or 'dept' in str(col).lower()), None)
                
                if dept_col:
                    unique_depts = filtered_df[dept_col].dropna().unique()
                    mapping_input = {}
                    mapping_tpl = st.session_state.get('template_mapping_dict', {})
                    
                    cols = st.columns(3)
                    for i, dept in enumerate(unique_depts):
                        d_clean = str(dept).strip()
                        def_idx = 0
                        if d_clean in mapping_tpl and mapping_tpl[d_clean] in st.session_state['staff_list']:
                            def_idx = st.session_state['staff_list'].index(mapping_tpl[d_clean])
                        
                        with cols[i % 3]:
                            mapping_input[dept] = st.selectbox(f"**{dept}**", st.session_state['staff_list'], index=def_idx, key=f"map_{dept}")

                    st.markdown("<br>", unsafe_allow_html=True)
                    if st.button("🚀 Proses Pencocokan Harga & Terapkan AI"):
                        filtered_df['Staff_Penanggung_Jawab'] = filtered_df[dept_col].map(mapping_input).fillna("Belum Ditugaskan")
                        
                        # LOGIKA PENCARIAN & PENYESUAIAN PRICE ZONE
                        df_zone = st.session_state['master_zone']
                        df_harga = st.session_state['master_harga']
                        
                        # Cari Price Zone berdasarkan Kode Toko di Sheet 2
                        zone_row = df_zone[df_zone.astype(str).apply(lambda x: x.str.contains(store_code_input, case=False)).any(axis=1)]
                        current_zone = "Zone_1" # Default fallback
                        if not zone_row.empty:
                            # Asumsi kolom zone ada di sheet 2
                            for col in df_zone.columns:
                                if 'zone' in str(col).lower() or 'price' in str(col).lower():
                                    current_zone = str(zone_row.iloc[0][col])
                        
                        # Mapping harga berdasarkan Sheet 1 & Zone
                        art_col_rep = next((c for c in filtered_df.columns if 'article' in str(c).lower()), None)
                        art_col_har = next((c for c in df_harga.columns if 'article' in str(c).lower() or 'plu' in str(c).lower()), None)
                        zone_price_col = next((c for c in df_harga.columns if current_zone.lower() in str(c).lower() or 'price' in str(c).lower()), None)
                        
                        if art_col_rep and art_col_har and zone_price_col:
                            harga_dict = dict(zip(df_harga[art_col_har].astype(str).str.strip(), df_harga[zone_price_col]))
                            filtered_df['Harga_Master'] = filtered_df[art_col_rep].astype(str).str.strip().map(harga_dict).fillna(0)
                        else:
                            filtered_df['Harga_Master'] = 0

                        # Kalkulasi Diskon 50%, 75%, 90%
                        try:
                            h_float = pd.to_numeric(filtered_df['Harga_Master'], errors='coerce').fillna(0)
                            filtered_df['Harga_Diskon_50'] = h_float * 0.50
                            filtered_df['Harga_Diskon_75'] = h_float * 0.25
                            filtered_df['Harga_Diskon_90'] = h_float * 0.10
                        except:
                            filtered_df['Harga_Diskon_50'] = 0
                            filtered_df['Harga_Diskon_75'] = 0
                            filtered_df['Harga_Diskon_90'] = 0

                        # Inisialisasi Status Pengerjaan Staf & Feedback
                        filtered_df['Status_Pengerjaan'] = "Belum Dikerjakan"
                        filtered_df['Feedback_Staff'] = "-"
                        filtered_df['Scan_Timestamp'] = "-"

                        st.session_state['filtered_data'] = filtered_df
                        st.success("✅ Price Zone berhasil dicocokkan dan kalkulasi diskon selesai!")
                        st.session_state['step'] = 3
                        st.rerun()
            else:
                st.warning("⚠️ Silakan upload Template Staf di menu expander atas terlebih dahulu.")

# ==========================================
# TAHAP 3: DASHBOARD SL/MOD & SCANNER FEEDBACK STAF
# ==========================================
elif st.session_state['step'] == 3:
    st.markdown("### 📊 Dashboard Monitoring SL & MOD")
    
    df_current = st.session_state['filtered_data']
    
    # Metrik Utama Dashboard
    total_tugas = len(df_current)
    selesai = len(df_current[df_current['Status_Pengerjaan'] == "Selesai Dikerjakan"])
    pending = total_tugas - selesai
    
    c1, c2, c3 = st.columns(3)
    c1.metric("Total Tugas Item ED", total_tugas)
    c2.metric("Selesai Dikerjakan Staf", selesai, delta=f"{(selesai/total_tugas*100):.1f}%" if total_tugas>0 else "0%")
    c3.metric("Belum Dikerjakan", pending)
    
    st.markdown("---")
    st.markdown("#### 📱 Portal Staf: Scanner & Input Feedback Tugas")
    st.info("Staf dapat melakukan scan artikel / memasukkan nomor artikel untuk melaporkan penyelesaian tugas.")
    
    col_sc1, col_sc2 = st.columns(2)
    with col_sc1:
        scan_input = st.text_input("Scan / Masukkan Barcode Artikel:")
    with col_sc2:
        staff_note = st.text_input("Catatan / Feedback Pengerjaan:")
        
    if st.button("Kirim Laporan Pengerjaan Staf"):
        if scan_input:
            match_idx = df_current[df_current.astype(str).apply(lambda x: x.str.contains(scan_input, case=False)).any(axis=1)].index
            if not match_idx.empty:
                df_current.loc[match_idx, 'Status_Pengerjaan'] = "Selesai Dikerjakan"
                df_current.loc[match_idx, 'Feedback_Staff'] = staff_note if staff_note else "Selesai dicek & ditindaklanjuti"
                df_current.loc[match_idx, 'Scan_Timestamp'] = pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")
                st.session_state['filtered_data'] = df_current
                st.success(f"✅ Artikel {scan_input} berhasil diperbarui statusnya menjadi 'Selesai Dikerjakan'!")
            else:
                st.error("Artikel tidak ditemukan dalam daftar tugas.")
        else:
            st.warning("Masukkan nomor artikel terlebih dahulu.")

    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("Lanjut ke Pusat Cetak Label & Unduhan ➡"):
        st.session_state['step'] = 4
        st.rerun()

# ==========================================
# TAHAP 4: CETAK LABEL CLEARANCE SIAP CETAK & DOWNLOAD
# ==========================================
elif st.session_state['step'] == 4:
    st.markdown("### 🏷️ Pusat Cetak Label Clearance & Laporan")
    
    final_df = st.session_state['filtered_data']
    
    tab1, tab2 = st.tabs(["📊 Laporan Utama & Monitoring", "🖨️ Cetak Label Clearance (50%, 75%, 90%)"])
    
    with tab1:
        st.markdown("#### Tabel Monitoring Real-Time Staf")
        st.dataframe(final_df[['Staff_Penanggung_Jawab', 'Status_Pengerjaan', 'Feedback_Staff', 'Scan_Timestamp']], use_container_width=True)
        
        def convert_df_to_excel(df_in):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_in.to_excel(writer, index=False, sheet_name='Monitoring_ED')
            return output.getvalue()
            
        st.download_button("📥 Download Rekap Laporan & Monitoring SL/MOD", data=convert_df_to_excel(final_df), file_name=f"Monitoring_Guardian_{st.session_state['store_code']}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    with tab2:
        st.markdown("#### Format Label Siap Cetak (Clearance Diskon)")
        st.caption("Label otomatis menghitung harga sesuai Price Zone dan potongan diskon 50%, 75%, serta 90%.")
        
        markdown_df = final_df[final_df.astype(str).apply(lambda x: x.str.contains("Markdown", case=False)).any(axis=1)].copy()
        
        if not markdown_df.empty:
            # Tampilkan format siap cetak
            display_label_df = pd.DataFrame()
            display_label_df['No'] = range(1, len(markdown_df) + 1)
            
            art_col = next((c for c in markdown_df.columns if 'article' in str(c).lower()), 'Article')
            desc_col = next((c for c in markdown_df.columns if 'desc' in str(c).lower()), 'Description')
            
            display_label_df['Artikel'] = markdown_df[art_col] if art_col in markdown_df.columns else "-"
            display_label_df['Nama Barang'] = markdown_df[desc_col] if desc_col in markdown_df.columns else "-"
            display_label_df['Harga Normal (Zone)'] = markdown_df['Harga_Master']
            display_label_df['Diskon 50%'] = markdown_df['Harga_Diskon_50']
            display_label_df['Diskon 75%'] = markdown_df['Harga_Diskon_75']
            display_label_df['Diskon 90%'] = markdown_df['Harga_Diskon_90']
            display_label_df['Toko Pemilik'] = st.session_state['store_code']
            display_label_df['Toko Titip Jual'] = st.session_state['store_code']
            
            st.dataframe(display_label_df, use_container_width=True)
            
            st.download_button(
                "📥 Download File Siap Cetak Label Clearance (.xlsx)",
                data=convert_df_to_excel(display_label_df),
                file_name=f"Label_Clearance_Siap_Cetak_{st.session_state['store_code']}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.info("Tidak ada item berstatus Markdown untuk dicetak labelnya.")

    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("⬅ Kembali ke Dashboard"):
        st.session_state['step'] = 3
        st.rerun()
