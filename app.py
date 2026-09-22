import pandas as pd
import streamlit as st
from io import BytesIO

# Konfigurasi Halaman
st.set_page_config(
    page_title="E-Smart Guardian - Retail System",
    page_icon="🛡️",
    layout="wide"
)

# Custom CSS
st.markdown("""
    <style>
    .main { background-color: #f4f6f9; }
    .stButton>button { width: 100%; background-color: #007bff; color: white; font-weight: bold; border-radius: 6px; height: 45px; }
    .stDownloadButton>button { width: 100%; background-color: #28a745; color: white; font-weight: bold; border-radius: 6px; height: 45px; }
    </style>
""", unsafe_allow_html=True)

# Inisialisasi State Alur & Data
if 'step' not in st.session_state:
    st.session_state['step'] = 1

if 'staff_list' not in st.session_state:
    st.session_state['staff_list'] = []

if 'store_name_dynamic' not in st.session_state:
    st.session_state['store_name_dynamic'] = "Toko Retail"

if 'ai_search_query' not in st.session_state:
    st.session_state['ai_search_query'] = ""

# Tombol Kembali ke Home di Sidebar
if st.session_state['step'] > 1:
    if st.sidebar.button("🏠 Kembali ke Menu Utama (Home)"):
        st.session_state['step'] = 1
        st.rerun()

# Header Utama
st.title("🛡️ E-Smart Guardian: Manajemen ED & Mitigasi Shrinkage")
st.markdown("---")

# ==========================================
# TAHAP 1: INPUT/HAPUS STAFF & UPLOAD FILE
# ==========================================
if st.session_state['step'] == 1:
    st.markdown("### **Langkah 1 dari 4: Pengaturan Staff & Unggah Data Laporan**")
    
    with st.expander("👥 Kelola / Tambah & Hapus Nama Staff Toko"):
        st.info("Tambahkan atau hapus nama staff sesuai kebutuhan toko.")
        
        def add_staff_callback():
            val = st.session_state.get("input_staff_baru", "").strip()
            if val and val not in st.session_state['staff_list']:
                st.session_state['staff_list'].append(val)
            st.session_state["input_staff_baru"] = ""

        st.text_input("Nama Staff Baru:", key="input_staff_baru")
        st.button("➕ Tambahkan Staff", on_click=add_staff_callback)
        
        st.markdown("---")
        st.write("**Daftar Staff Saat Ini:**")
        if st.session_state['staff_list']:
            for idx, staff in enumerate(st.session_state['staff_list']):
                col_s1, col_s2 = st.columns([4, 1])
                col_s1.write(f"- {staff}")
                if col_s2.button("🗑️ Hapus", key=f"del_staff_{idx}"):
                    st.session_state['staff_list'].pop(idx)
                    st.rerun()
        else:
            st.warning("Belum ada staff yang ditambahkan.")

    st.markdown("---")
    uploaded_files = st.file_uploader(
        "Pilih file data laporan (Mendukung SEMUA format Excel .xlsx, .xlsb, .xls & CSV, maks 5 file):", 
        type=["csv", "xlsx", "xls", "xlsb", "xlsm"], 
        accept_multiple_files=True
    )

    col_b1, col_b2 = st.columns([4, 1])
    with col_b2:
        if st.button("Next ➡️"):
            if uploaded_files:
                if len(uploaded_files) > 5:
                    st.error("Maksimal hanya 5 file yang dapat diunggah sekaligus!")
                else:
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
                st.warning("Harap unggah file terlebih dahulu.")

# ==========================================
# TAHAP 2: KODE TOKO KOSONG & MAPPING STAFF
# ==========================================
elif st.session_state['step'] == 2:
    st.markdown("### **Langkah 2 dari 4: Masukkan Kode Toko & Mapping Tanggung Jawab Staff (By Category)**")
    
    # Tombol Back saja di atas
    col_nav1, _ = st.columns([1, 4])
    with col_nav1:
        if st.button("⬅️ Back"):
            st.session_state['step'] = 1
            st.rerun()

    store_code = st.text_input("Masukkan Kode Toko Anda (Contoh: 6849):", value="")
    
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

        st.success(f"🏢 Toko Terdeteksi: **{st.session_state['store_name_dynamic']}** (Total Data: {len(filtered_df)} baris)")

        st.markdown("---")
        st.markdown("#### **Mapping Penanggung Jawab Berdasarkan Kategori (Category)**")
        
        if not st.session_state['staff_list']:
            st.warning("⚠️ Belum ada nama staff yang dimasukkan. Harap kembali ke Langkah 1 dan tambahkan minimal 1 nama staff.")
        else:
            cat_col = next((col for col in raw_df.columns if str(col).upper() == 'CAT' or 'cat' in str(col).lower()), None)
            
            if cat_col:
                unique_cats = filtered_df[cat_col].dropna().unique()
                mapping_input = {}
                
                st.info(f"Ditemukan {len(unique_cats)} Kategori produk. Silakan tentukan penanggung jawabnya:")
                for cat in unique_cats: 
                    mapping_input[cat] = st.selectbox(f"Staff untuk Kategori: **{cat}**", st.session_state['staff_list'], key=f"map_cat_{cat}")

                # Tombol Aksi di Bawah (Terapkan Mapping & Next Berdampingan)
                col_act1, col_act2 = st.columns(2)
                with col_act1:
                    btn_apply = st.button("🚀 Terapkan Mapping Category & Jalankan AI")
                with col_act2:
                    btn_next_step = st.button("Next ➡️ (Lanjut ke Review)")

                if btn_apply:
                    filtered_df['Staff_Penanggung_Jawab'] = filtered_df[cat_col].map(mapping_input).fillna("Belum Ditugaskan")
                    
                    remark_col = next((col for col in filtered_df.columns if 'remark' in str(col).lower() or 'feedback' in str(col).lower() or 'instruction' in str(col).lower()), None)
                    brand_col = next((col for col in raw_df.columns if 'brand' in str(col).lower()), None)

                    def smart_ai_recommendation(row):
                        remark_text = ""
                        if remark_col:
                            remark_text = str(row[remark_col]).lower()
                        
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
                    st.success("Mapping dan Analisis AI Berhasil Diterapkan!")

                if btn_next_step:
                    if 'filtered_data' in st.session_state:
                        st.session_state['step'] = 3
                        st.rerun()
                    else:
                        st.warning("Harap klik tombol 'Terapkan Mapping Category & Jalankan AI' terlebih dahulu.")
            else:
                st.error("Kolom 'Cat' (Category) tidak ditemukan pada struktur file.")
    else:
        st.info("👆 Silakan masukkan Kode Toko di atas untuk memunculkan data dan opsi mapping.")

# ==========================================
# TAHAP 3: REVIEW, SMART SEARCH & FEEDBACK
# ==========================================
elif st.session_state['step'] == 3:
    st.markdown(f"### **Langkah 3 dari 4: Review Data & Koreksi Staf ({st.session_state['store_name_dynamic']})**")
    
    st.markdown("#### **🔍 AI Smart Search & Filter Engine**")
    st.info("Ketik kata kunci untuk menyaring data secara instan.")
    
    ai_query = st.text_input(
        "Ketik kata kunci pencarian:",
        value=st.session_state['ai_search_query']
    )
    st.session_state['ai_search_query'] = ai_query

    col_nav1, col_nav2 = st.columns(2)
    with col_nav1:
        if st.button("⬅️ Back"):
            st.session_state['step'] = 2
            st.rerun()
    with col_nav2:
        if st.button("Next (Lanjut ke Download) ➡️"):
            st.session_state['step'] = 4
            st.rerun()

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
            if any(k in c_low for k in ['code', 'site', 'desc', 'am', 'article', 'brand', 'cat', 'staff', 'instruksi', 'aksi', 'feedback', 'remark']):
                possible_cols.append(col)

        table_view_df = display_df[possible_cols] if possible_cols else display_df

        st.markdown("#### **📝 Tabel Review & Catatan Koreksi Staf**")
        st.info("💡 Anda hanya dapat mengedit kolom **Feedback_Staff** di tabel bawah. Kolom lainnya terkunci agar aman.")

        column_config = {}
        for col in table_view_df.columns:
            if col != 'Feedback_Staff':
                column_config[col] = st.column_config.Column(disabled=True)

        edited_table = st.data_editor(
            table_view_df,
            use_container_width=True,
            height=400,
            hide_index=True,
            column_config=column_config,
            key="editor_review_feedback"
        )

        for idx in edited_table.index:
            if idx in current_data.index:
                current_data.at[idx, 'Feedback_Staff'] = edited_table.at[idx, 'Feedback_Staff']

        st.session_state['final_edited_data'] = current_data
    else:
        st.warning("Belum ada data toko yang diproses. Silakan kembali ke Langkah 2.")

# ==========================================
# TAHAP 4: DOWNLOAD LAPORAN TERESEDIKASI
# ==========================================
elif st.session_state['step'] == 4:
    st.markdown(f"### **Langkah 4 dari 4: Unduh Laporan Resmi Toko ({st.session_state['store_name_dynamic']})**")
    
    if st.button("⬅️ Back ke Menu Review"):
        st.session_state['step'] = 3
        st.rerun()

    st.markdown("---")
    if 'final_edited_data' in st.session_state:
        final_df = st.session_state['final_edited_data']

        col_dl1, col_dl2 = st.columns(2)

        def convert_df_to_excel(df_in):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_in.to_excel(writer, index=False, sheet_name='Laporan_Final')
            return output.getvalue()

        with col_dl1:
            st.markdown("#### **📥 Rekap Laporan Toko Keseluruhan**")
            excel_all = convert_df_to_excel(final_df)
            st.download_button(
                label="Download Rekap Lengkap (Excel)",
                data=excel_all,
                file_name=f"Laporan_Final_{st.session_state['store_name_dynamic'].replace(' ', '_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        with col_dl2:
            st.markdown("#### **📦 Khusus File Barang Markdown**")
            markdown_df = final_df[final_df.astype(str).apply(lambda x: x.str.contains("Markdown", case=False)).any(axis=1)]
            excel_markdown = convert_df_to_excel(markdown_df)
            st.download_button(
                label="Download File Markdown Saja (Excel)",
                data=excel_markdown,
                file_name=f"File_Markdown_{st.session_state['store_name_dynamic'].replace(' ', '_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.warning("Tidak ada data untuk diunduh.")