import pandas as pd
import streamlit as st
from io import BytesIO

# ==========================================
# KONFIGURASI HALAMAN & CUSTOM CSS
# ==========================================
st.set_page_config(
    page_title="E-Smart Guardian",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
    <style>
    /* Mengurangi jarak kosong di bagian atas */
    .block-container { padding-top: 2rem; padding-bottom: 2rem; }
    
    /* Warna latar belakang dan font utama */
    .main { background-color: #f8fafc; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
    
    /* Warna header */
    h1, h2, h3, h4 { color: #1e293b; font-weight: 700; }
    
    /* Animasi dan Gaya Tombol Biru (Primary) */
    .stButton>button { 
        width: 100%; 
        background-color: #3b82f6; 
        color: white; 
        font-weight: 600; 
        border-radius: 8px; 
        height: 48px; 
        border: none;
        transition: all 0.3s ease;
    }
    .stButton>button:hover {
        background-color: #2563eb;
        transform: translateY(-2px);
        box-shadow: 0 4px 12px rgba(59, 130, 246, 0.3);
    }
    
    /* Animasi dan Gaya Tombol Hijau (Download) */
    .stDownloadButton>button { 
        width: 100%; 
        background-color: #10b981; 
        color: white; 
        font-weight: 600; 
        border-radius: 8px; 
        height: 48px; 
        border: none;
        transition: all 0.3s ease;
    }
    .stDownloadButton>button:hover {
        background-color: #059669;
        transform: translateY(-2px);
        box-shadow: 0 4px 12px rgba(16, 185, 129, 0.3);
    }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# INISIALISASI STATE
# ==========================================
if 'step' not in st.session_state:
    st.session_state['step'] = 1
if 'staff_list' not in st.session_state:
    st.session_state['staff_list'] = []
if 'store_name_dynamic' not in st.session_state:
    st.session_state['store_name_dynamic'] = "Toko Retail"
if 'ai_search_query' not in st.session_state:
    st.session_state['ai_search_query'] = ""
if 'template_mapping_dict' not in st.session_state:
    st.session_state['template_mapping_dict'] = {}

# ==========================================
# SIDEBAR
# ==========================================
with st.sidebar:
    # URL Gambar telah diubah menjadi ikon Perisai (Shield) yang lebih relevan
    st.image("https://cdn-icons-png.flaticon.com/512/1161/1161388.png", width=80) 
    st.markdown("### 🛡️ E-Smart Guardian")
    st.caption("Sistem Manajemen ED & Mitigasi Shrinkage Retail.")
    st.markdown("---")
    
    if st.session_state['step'] > 1:
        if st.button("🏠 Kembali ke Menu Utama"):
            st.session_state['step'] = 1
            st.rerun()

# ==========================================
# HEADER & PROGRESS BAR
# ==========================================
st.title("🛡️ E-Smart Guardian")
st.markdown("Membantu memproses data *Expired Date*, pemetaan tugas staf, dan *Clearance Label* secara otomatis.")

# Progress Bar Visual
progress_value = int((st.session_state['step'] / 4) * 100)
st.progress(st.session_state['step'] / 4, text=f"Tahap {st.session_state['step']} dari 4 ({progress_value}%)")
st.markdown("<br>", unsafe_allow_html=True)

# ==========================================
# TAHAP 1: TEMPLATE & UPLOAD DATA
# ==========================================
if st.session_state['step'] == 1:
    
    # KONTANER 1: PENGATURAN TEMPLATE (DISMBUNYIKAN DALAM EXPANDER AGAR RAPI)
    with st.expander("⚙️ Pengaturan Daftar Staf & Pembagian Dept (Klik untuk membuka)", expanded=False):
        st.info("Seluruh Departemen (Dept) sudah terisi otomatis! Anda hanya perlu mengisi Nama Staff di baris yang ditugaskan, lalu upload kembali.")
        
        col_tpl1, col_tpl2 = st.columns(2)
        
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
            
        with col_tpl1:
            st.markdown("**1. Download Template Kosong**")
            st.download_button(
                label="📥 Download Template Cek ED",
                data=generate_template(),
                file_name="Template_Pembagian_Cek_ED_By_Dept.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        with col_tpl2:
            st.markdown("**2. Upload Template Terisi**")
            uploaded_template = st.file_uploader("Format Excel (.xlsx, .xls)", type=["xlsx", "xls"], label_visibility="collapsed")
            
            if uploaded_template:
                try:
                    df_mapping = pd.read_excel(uploaded_template)
                    if 'Dept' in df_mapping.columns and 'Nama Staff' in df_mapping.columns:
                        df_mapping_clean = df_mapping.dropna(subset=['Dept', 'Nama Staff'])
                        df_mapping_clean = df_mapping_clean[df_mapping_clean['Nama Staff'].astype(str).str.strip() != '']
                        df_mapping_clean = df_mapping_clean[df_mapping_clean['Nama Staff'].astype(str).str.strip().str.lower() != 'nan']
                        
                        staff_from_tpl = df_mapping_clean['Nama Staff'].astype(str).str.strip().unique().tolist()
                        
                        st.session_state['staff_list'] = []
                        for s in staff_from_tpl:
                            if s and s not in st.session_state['staff_list']:
                                st.session_state['staff_list'].append(s)
                                
                        mapping_dict = dict(zip(df_mapping_clean['Dept'].astype(str).str.strip(), df_mapping_clean['Nama Staff'].astype(str).str.strip()))
                        st.session_state['template_mapping_dict'] = mapping_dict
                        st.success("✅ Daftar staf berhasil dimuat!")
                    else:
                        st.error("Format salah. Pastikan kolom 'Dept' dan 'Nama Staff' ada.")
                except Exception as e:
                    st.error(f"Gagal membaca: {e}")

    # KONTANER 2: UPLOAD DATA UTAMA (DALAM KOTAK RAPI)
    with st.container(border=True):
        st.markdown("#### 📁 Unggah Data Laporan (File Master)")
        st.markdown("Silakan pilih file laporan dari pusat untuk dianalisis. Mendukung format `.xlsx`, `.xlsb`, atau `.csv`.")
        
        uploaded_files = st.file_uploader(
            "Upload area", 
            type=["csv", "xlsx", "xls", "xlsb", "xlsm"], 
            accept_multiple_files=True,
            label_visibility="collapsed"
        )

        st.markdown("<br>", unsafe_allow_html=True)
        col_btn1, col_btn2, col_btn3 = st.columns([1, 1, 1])
        with col_btn3:
            if st.button("Lanjut ke Langkah 2 ➡"):
                if uploaded_files:
                    if len(uploaded_files) > 5:
                        st.error("Maksimal hanya 5 file sekaligus!")
                    else:
                        with st.spinner('Memproses file... mohon tunggu'):
                            all_data = []
                            for file in uploaded_files:
                                try:
                                    file_name_lower = file.name.lower()
                                    if file_name_lower.endswith('.csv'):
                                        df_temp = pd.read_csv(file, sep=';', header=1, on_bad_lines='skip', dtype=str)
                                    elif file_name_lower.endswith('.xlsb'):
                                        df_temp = pd.read_excel(file, header=1, engine='pyxlsb', dtype=str)
                                    else:
                                        df_temp = pd.read_excel(file, header=1, dtype=str)
                                    all_data.append(df_temp)
                                except Exception as e:
                                    st.error(f"Gagal membaca file {file.name}: {e}")
                            
                            if all_data:
                                combined_df = pd.concat(all_data, ignore_index=True)
                                combined_df.columns = [str(c).strip() for c in combined_df.columns]
                                combined_df = combined_df.loc[:, ~combined_df.columns.str.contains('^Unnamed|^None', case=False, na=False)]
                                combined_df = combined_df.dropna(how='all', axis=1)

                                st.session_state['raw_data'] = combined_df
                                st.session_state['step'] = 2
                                st.rerun()
                else:
                    st.warning("Harap unggah file laporan terlebih dahulu.")

# ==========================================
# TAHAP 2: KODE TOKO KOSONG & MAPPING STAFF
# ==========================================
elif st.session_state['step'] == 2:
    
    with st.container(border=True):
        st.markdown("#### 🏢 Identifikasi Toko")
        col_input1, col_input2 = st.columns([3, 1])
        
        with col_input1:
            store_code = st.text_input("Masukkan 4 Digit Kode Toko (Contoh: 6849):", value="")
            
    if store_code:
        raw_df = st.session_state['raw_data']
        code_col = next((col for col in raw_df.columns if 'code' in str(col).lower() or 'site' in str(col).lower()), None)
        site_desc_col = next((col for col in raw_df.columns if 'site desc' in str(col).lower() or 'store' in str(col).lower() or 'desc' in str(col).lower()), None)

        if code_col:
            filtered_df = raw_df[raw_df[code_col].astype(str).str.contains(store_code, na=False)].copy()
            if site_desc_col and not filtered_df.empty:
                st.session_state['store_name_dynamic'] = str(filtered_df[site_desc_col].iloc[0])
            else:
                st.session_state['store_name_dynamic'] = f"Toko Kode {store_code}"
        else:
            filtered_df = raw_df.copy()
            st.session_state['store_name_dynamic'] = f"Toko Kode {store_code}"

        # METRIC CARD UNTUK MENAMPILKAN INFO TOKO
        col_m1, col_m2 = st.columns(2)
        col_m1.metric(label="Nama Toko Aktif", value=st.session_state['store_name_dynamic'])
        col_m2.metric(label="Total Baris Data Ditemukan", value=len(filtered_df))

        st.markdown("<hr style='margin: 10px 0 25px 0'>", unsafe_allow_html=True)
        
        if not st.session_state['staff_list']:
            st.warning("⚠️ Belum ada nama staf yang terdeteksi. Silakan upload template di Langkah 1.")
            if st.button("⬅️ Kembali ke Langkah 1"):
                st.session_state['step'] = 1
                st.rerun()
        else:
            with st.container(border=True):
                st.markdown("#### 👥 Mapping Penanggung Jawab Berdasarkan Departemen")
                st.caption("Sistem otomatis memuat data staf dari template yang Anda unggah sebelumnya.")
                
                dept_col = next((col for col in raw_df.columns if str(col).upper() == 'DEPT' or 'dept' in str(col).lower()), None)
                
                if dept_col:
                    unique_depts = filtered_df[dept_col].dropna().unique()
                    mapping_input = {}
                    mapping_dict_tpl = st.session_state.get('template_mapping_dict', {})
                    
                    # Buat layout grid (3 kolom) agar kotak selectbox tidak memanjang ke bawah
                    cols = st.columns(3)
                    for i, dept in enumerate(unique_depts): 
                        dept_clean = str(dept).strip()
                        default_idx = 0
                        
                        if dept_clean in mapping_dict_tpl:
                            staff_name = mapping_dict_tpl[dept_clean]
                            if staff_name in st.session_state['staff_list']:
                                default_idx = st.session_state['staff_list'].index(staff_name)
                        
                        with cols[i % 3]:
                            mapping_input[dept] = st.selectbox(f"**{dept}**", st.session_state['staff_list'], index=default_idx, key=f"map_dept_{dept}")

                    st.markdown("<br>", unsafe_allow_html=True)
                    
                    col_act1, col_act2, col_act3 = st.columns([1, 1, 1])
                    with col_act1:
                        if st.button("⬅️ Kembali"):
                            st.session_state['step'] = 1
                            st.rerun()
                    with col_act2:
                        btn_apply = st.button("🚀 Terapkan & Analisis AI")
                    with col_act3:
                        btn_next_step = st.button("Lanjut ke Review ➡")

                    if btn_apply:
                        with st.spinner("Menerapkan strategi AI..."):
                            filtered_df['Staff_Penanggung_Jawab'] = filtered_df[dept_col].map(mapping_input).fillna("Belum Ditugaskan")
                            remark_col = next((col for col in filtered_df.columns if 'remark' in str(col).lower() or 'feedback' in str(col).lower() or 'instruction' in str(col).lower()), None)

                            def smart_ai_recommendation(row):
                                remark_text = str(row[remark_col]).lower() if remark_col else ""
                                if 'expired' in remark_text or 'lewat' in remark_text:
                                    return f"⚠️ [EXPIRED LEWAT]: Segera ajukan ajuan WO tambahan."
                                elif 'blue dot' in remark_text or 'not approved' in remark_text:
                                    return f"🚨 [BLUE DOT / HOLD]: Pisahkan fisik untuk Program GWP atau Retur ke DC."
                                elif any(kw in remark_text for kw in ['markdown', 'rtw', 'return', 'disetujui', 'approved']):
                                    return row[remark_col]
                                else:
                                    return f"💡 [STRATEGI TOKO]: Pindahkan ke Eye-Level Display & Bundling Silang."

                            filtered_df['Instruksi_Aksi'] = filtered_df.apply(smart_ai_recommendation, axis=1)
                            
                            if 'Feedback_Staff' not in filtered_df.columns:
                                filtered_df['Feedback_Staff'] = "Belum ada catatan"

                            st.session_state['filtered_data'] = filtered_df
                            st.success("✅ Selesai! Mapping Dept dan Analisis AI Berhasil Diterapkan. Silakan klik 'Lanjut ke Review'.")

                    if btn_next_step:
                        if 'filtered_data' in st.session_state:
                            st.session_state['step'] = 3
                            st.rerun()
                        else:
                            st.error("Klik tombol 'Terapkan & Analisis AI' terlebih dahulu.")
                else:
                    st.error("Kolom 'Dept' tidak ditemukan pada struktur file.")

# ==========================================
# TAHAP 3: REVIEW, SMART SEARCH & FEEDBACK
# ==========================================
elif st.session_state['step'] == 3:
    
    with st.container(border=True):
        st.markdown("#### 🔍 AI Smart Search & Filter Engine")
        ai_query = st.text_input(
            "Cari berdasarkan Article, Nama Barang, Brand, atau Remark:",
            value=st.session_state['ai_search_query'],
            placeholder="Ketik kata kunci di sini..."
        )
        st.session_state['ai_search_query'] = ai_query

    if 'filtered_data' in st.session_state:
        current_data = st.session_state['filtered_data']

        if ai_query:
            mask = current_data.apply(lambda row: row.astype(str).str.contains(ai_query, case=False).any(), axis=1)
            display_df = current_data[mask]
        else:
            display_df = current_data

        possible_cols = []
        for col in display_df.columns:
            c_low = str(col).lower()
            if any(k in c_low for k in ['code', 'site', 'desc', 'am', 'article', 'brand', 'dept', 'staff', 'instruksi', 'aksi', 'feedback', 'remark']):
                possible_cols.append(col)

        table_view_df = display_df[possible_cols] if possible_cols else display_df

        with st.container(border=True):
            st.markdown("#### 📝 Tabel Review & Catatan Koreksi Staf")
            st.caption("Tabel ini interaktif. Anda dapat mengedit langsung pada kolom **Feedback_Staff**.")

            column_config = {}
            for col in table_view_df.columns:
                if col != 'Feedback_Staff':
                    column_config[col] = st.column_config.Column(disabled=True)

            edited_table = st.data_editor(
                table_view_df,
                use_container_width=True,
                height=450,
                hide_index=True,
                column_config=column_config,
                key="editor_review_feedback"
            )

            for idx in edited_table.index:
                if idx in current_data.index:
                    current_data.at[idx, 'Feedback_Staff'] = edited_table.at[idx, 'Feedback_Staff']

            st.session_state['final_edited_data'] = current_data
            
        st.markdown("<br>", unsafe_allow_html=True)
        col_nav1, col_nav2, col_nav3 = st.columns([1, 1, 1])
        with col_nav1:
            if st.button("⬅️ Kembali"):
                st.session_state['step'] = 2
                st.rerun()
        with col_nav3:
            if st.button("Lanjut ke Unduhan ➡"):
                st.session_state['step'] = 4
                st.rerun()
    else:
        st.warning("Belum ada data toko yang diproses.")

# ==========================================
# TAHAP 4: DOWNLOAD LAPORAN TERESEDIKASI
# ==========================================
elif st.session_state['step'] == 4:
    
    with st.container(border=True):
        st.markdown("#### 📦 Pusat Unduhan Laporan Resmi")
        st.caption(f"Semua file siap diunduh untuk toko: **{st.session_state['store_name_dynamic']}**")
        
        if 'final_edited_data' in st.session_state:
            final_df = st.session_state['final_edited_data']

            # Logika filtering kolom
            cols_to_remove = ['city', 'region', 'dept', 'brand', 'dot', 'code.1', 'exp', 'cost', 'total value', 'sos']
            columns_to_keep = [col for col in final_df.columns if str(col).lower().strip() not in cols_to_remove and 'list markdown' not in str(col).lower().strip() and '3773603066' not in str(col).lower().strip()]
                    
            download_df = final_df[columns_to_keep]

            def convert_df_to_excel(df_in):
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_in.to_excel(writer, index=False, sheet_name='Laporan_Final')
                return output.getvalue()

            st.markdown("<br>", unsafe_allow_html=True)
            col_dl1, col_space, col_dl2 = st.columns([10, 1, 10])

            # BAGIAN 1: REKAP LENGKAP
            with col_dl1:
                with st.container(border=True):
                    st.markdown("### 📊 Rekap Keseluruhan")
                    st.markdown("Laporan utama berisi instruksi AI dan penugasan semua departemen.")
                    st.markdown("<br>", unsafe_allow_html=True)
                    
                    excel_all = convert_df_to_excel(download_df)
                    st.download_button(
                        label="📥 Download Excel Utama",
                        data=excel_all,
                        file_name=f"Laporan_Final_{st.session_state['store_name_dynamic'].replace(' ', '_')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

            # BAGIAN 2: FILE MARKDOWN
            with col_dl2:
                with st.container(border=True):
                    st.markdown("### 🏷️ Khusus Label Markdown")
                    st.markdown("Laporan siap tempel (*copy-paste*) khusus untuk aplikasi *Clearance Label*.")
                    st.markdown("<br>", unsafe_allow_html=True)
                    
                    markdown_source_df = final_df[final_df.astype(str).apply(lambda x: x.str.contains("Markdown", case=False)).any(axis=1)].copy()
                    target_cols = ['CN', 'Artikel', 'Desc', 'UOM', 'OUM2', 'Avg', 'SOH', 'SOL', 'OOQ', 'GIT', 'Requested', 'Remaks']
                    final_markdown_df = pd.DataFrame(index=markdown_source_df.index, columns=target_cols)
                    
                    if not markdown_source_df.empty:
                        final_markdown_df['CN'] = range(1, len(markdown_source_df) + 1)
                        
                        col_article = next((c for c in markdown_source_df.columns if str(c).strip().lower() == 'article'), None)
                        col_desc = next((c for c in markdown_source_df.columns if 'article desc' in str(c).lower() or 'desc' in str(c).lower() and 'site' not in str(c).lower()), None)
                        col_soh = next((c for c in markdown_source_df.columns if str(c).strip().lower() == 'soh' or 'soh' in str(c).lower()), None)
                        col_remark = next((c for c in markdown_source_df.columns if 'remark' in str(c).lower()), None)
                        col_month = next((c for c in markdown_source_df.columns if str(c).strip().lower() == 'month'), None)
                        col_year = next((c for c in markdown_source_df.columns if str(c).strip().lower() == 'year'), None)
                        col_date = next((c for c in markdown_source_df.columns if 'date' in str(c).lower() or 'tanggal' in str(c).lower() or 'exp' in str(c).lower()), None)
                        
                        if col_article: final_markdown_df['Artikel'] = markdown_source_df[col_article]
                        if col_desc: final_markdown_df['Desc'] = markdown_source_df[col_desc]
                        if col_soh: final_markdown_df['SOH'] = markdown_source_df[col_soh]
                        if col_remark: final_markdown_df['Remaks'] = markdown_source_df[col_remark]
                            
                        # Logika Format Kode Requested
                        if col_month and col_year:
                            bulan_bersih = markdown_source_df[col_month].fillna(0).astype(float).astype(int).astype(str)
                            tahun_bersih = markdown_source_df[col_year].fillna(0).astype(float).astype(int).astype(str)
                            bulan_str = bulan_bersih.str.zfill(2)
                            tahun_digit = tahun_bersih.str[-1]
                            final_markdown_df['Requested'] = bulan_str + tahun_digit + "01"
                            final_markdown_df.loc[bulan_bersih == '0', 'Requested'] = '-'
                        elif col_date:
                            final_markdown_df['Requested'] = markdown_source_df[col_date]
                    
                    final_markdown_df = final_markdown_df.fillna('-')
                    
                    excel_markdown = convert_df_to_excel(final_markdown_df)
                    st.download_button(
                        label="📥 Download Excel Markdown",
                        data=excel_markdown,
                        file_name=f"File_Markdown_{st.session_state['store_name_dynamic'].replace(' ', '_')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
        else:
            st.warning("Tidak ada data untuk diunduh.")
            
    st.markdown("<br>", unsafe_allow_html=True)
    col_nav1, _, _ = st.columns([1, 1, 1])
    with col_nav1:
        if st.button("⬅️️ Kembali ke Review"):
            st.session_state['step'] = 3
            st.rerun()
