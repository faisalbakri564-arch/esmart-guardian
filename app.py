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
    # URL Gambar menggunakan logo Guardian dari internet
    st.image("https://res.cloudinary.com/malls-id/image/upload/v1584069814/malls/brands/guardian.png", width=180) 
    st.markdown("### E-Smart Guardian")
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
    with st.expander("⚙️️ Pengaturan Daftar Staf & Pembagian Dept (Klik untuk membuka)", expanded=False):
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
                        with st.spinner('Memproses file
