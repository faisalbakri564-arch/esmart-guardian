import pandas as pd
import streamlit as st
from io import BytesIO
import os
from datetime import datetime

# ==========================================
# KONFIGURASI HALAMAN
# ==========================================
st.set_page_config(page_title="E-Smart Guardian", page_icon="🛡️", layout="wide")

# ==========================================
# FUNGSI DATABASE LOKAL (SIMULASI GOOGLE SHEETS)
# ==========================================
DB_FILE = "db_tugas_staff.csv"

def load_db():
    if os.path.exists(DB_FILE):
        return pd.read_csv(DB_FILE)
    else:
        # Jika belum ada, buat struktur tabelnya (mirip kolom di Google Sheets nanti)
        return pd.DataFrame(columns=["Waktu_Dibuat", "Kode_Toko", "Dept", "Nama_Staff", "Article", "Nama_Barang", "Instruksi", "Status", "Waktu_Selesai"])

def save_to_db(df):
    df.to_csv(DB_FILE, index=False)

# ==========================================
# FUNGSI BACA PRICE ZONE
# ==========================================
@st.cache_data
def load_price_zone():
    try:
        # Baca Sheet 1 (Harga) dan Sheet 2 (Mapping Toko)
        df_harga = pd.read_excel('FILE PRICE ZONE.xlsx', sheet_name=0)
        df_toko = pd.read_excel('FILE PRICE ZONE.xlsx', sheet_name=1)
        
        # Buat dictionary untuk pencarian cepat
        store_to_pz = dict(zip(df_toko['Title'].astype(str), df_toko['PZ'].astype(str)))
        
        # Buat dictionary harga per artikel
        article_prices = {}
        for _, row in df_harga.iterrows():
            article = str(row['Article'])
            article_prices[article] = row.to_dict()
            
        return store_to_pz, article_prices
    except Exception as e:
        return {}, {}

store_to_pz, article_prices = load_price_zone()

def get_original_price(store_code, article):
    pz_code = store_to_pz.get(str(store_code), "P37") # Default P37
    article = str(article)
    if article in article_prices:
        # Kembalikan harga sesuai PZ, jika tidak ada, ambil 0
        return article_prices[article].get(pz_code, 0)
    return 0

# ==========================================
# FUNGSI CETAK LABEL HTML
# ==========================================
def generate_label_html(df_markdown, store_code, discount_rate):
    html_content = """
    <html><head><style>
        body { font-family: Arial, sans-serif; margin: 0; padding: 10px; }
        .grid { display: grid; grid-template-columns: repeat(5, 1fr); gap: 2px; }
        .label { border: 1px dashed #000; width: 150px; height: 75px; position: relative; padding: 5px; background: #fff; display: inline-block; margin-right: 5px; margin-bottom: 5px;}
        .disc { font-size: 11px; font-weight: bold; position: absolute; top: 5px; left: 5px;}
        .new-price { font-size: 18px; font-weight: 900; position: absolute; top: 5px; right: 5px; }
        .old-price { font-size: 10px; text-decoration: line-through; position: absolute; top: 25px; right: 5px; }
        .barcode { font-size: 10px; position: absolute; bottom: 5px; left: 5px; }
        .codes { font-size: 9px; position: absolute; bottom: 5px; right: 5px; }
        @media print { button { display: none; } }
    </style></head><body>
    <button onclick="window.print()" style="padding:10px; margin-bottom:15px; cursor:pointer;">🖨️ Cetak Label Sekarang</button>
    <div class="grid">
    """
    
    # Titip jual statis (contoh)
    kode_titip = "899X" 
    
    for _, row in df_markdown.iterrows():
        article = str(row.get('Article', row.get('Artikel', '')))
        harga_asli = get_original_price(store_code, article)
        
        # Hitung potongan diskon (50, 75, atau 90)
        harga_diskon = int(harga_asli - (harga_asli * (discount_rate / 100.0)))
        
        # Format Rupiah
        str_harga_asli = f"{harga_asli:,.0f}".replace(',', '.')
        str_harga_diskon = f"{harga_diskon:,.0f}".replace(',', '.')
        
        html_content += f"""
        <div class="label">
            <div class="disc">Disc.<br>{discount_rate}%</div>
            <div class="new-price">{str_harga_diskon}</div>
            <div class="old-price">{str_harga_asli}</div>
            <div class="barcode">{article}</div>
            <div class="codes">{store_code}<br>{kode_titip}</div>
        </div>
        """
        
    html_content += "</div></body></html>"
    return html_content

# ==========================================
# SIDEBAR: PEMILIHAN PERAN (ROLE)
# ==========================================
with st.sidebar:
    if os.path.exists("logo.png"):
        st.image("logo.png", width=180)
    st.markdown("### 🛡️ E-Smart Guardian")
    
    # Pilih Menu
    menu = st.selectbox("Masuk Sebagai:", ["SL / MOD (Dashboard)", "Staff (Scanner)"])
    st.markdown("---")

# ==========================================
# MENU 1: STAFF (SCANNER & FEEDBACK)
# ==========================================
if menu == "Staff (Scanner)":
    st.title("📱 Menu Scanner Staff")
    st.markdown("Scan barcode atau ketik Article barang yang sudah Anda proses/display.")
    
    db = load_db()
    
    if db.empty:
        st.info("Belum ada tugas yang dibagikan oleh Store Leader.")
    else:
        # Ambil daftar staff dari database
        staff_list = db['Nama_Staff'].unique().tolist()
        nama_staff = st.selectbox("Pilih Nama Anda:", ["-- Pilih Nama --"] + staff_list)
        
        if nama_staff != "-- Pilih Nama --":
            st.subheader(f"Halo, {nama_staff}!")
            
            # Tampilkan tugas yang BELUM selesai
            tugas_saya = db[(db['Nama_Staff'] == nama_staff) & (db['Status'] == 'Belum')]
            
            col1, col2 = st.columns([1, 2])
            with col1:
                st.metric("Sisa Tugas Anda", len(tugas_saya))
            
            with col2:
                # Form Scanner
                scanned_article = st.text_input("Scan/Ketik Article (PLU):", placeholder="Contoh: 123456")
                if st.button("✅ Tandai Selesai"):
                    if scanned_article:
                        # Cek apakah article ada di tugas staff ini
                        idx = db[(db['Nama_Staff'] == nama_staff) & (db['Article'] == scanned_article) & (db['Status'] == 'Belum')].index
                        if len(idx) > 0:
                            db.loc[idx, 'Status'] = 'Selesai'
                            db.loc[idx, 'Waktu_Selesai'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                            save_to_db(db)
                            st.success(f"Tugas untuk Article {scanned_article} berhasil diselesaikan!")
                            st.rerun()
                        else:
                            st.error(f"Article {scanned_article} tidak ditemukan dalam tugas Anda atau sudah selesai.")
                    else:
                        st.warning("Masukkan Article terlebih dahulu.")
            
            st.markdown("**Daftar Tugas Anda yang Belum Selesai:**")
            st.dataframe(tugas_saya[['Article', 'Nama_Barang', 'Instruksi']])

# ==========================================
# MENU 2: SL / MOD (DASHBOARD & GENERATOR)
# ==========================================
elif menu == "SL / MOD (Dashboard)":
    st.title("📊 Dashboard SL & MOD")
    
    tab1, tab2, tab3 = st.tabs(["1. Upload & Mapping", "2. Cetak Label Markdown", "3. Monitoring Staff"])
    
    # TAB 1: UPLOAD DATA & MAPPING
    with tab1:
        st.markdown("Unggah data laporan ED dari pusat dan distribusikan ke Staff.")
        
        # Sederhanakan mapping staff (Contoh: ketik manual nama staff per dept)
        col_toko, col_file = st.columns(2)
        with col_toko:
            kode_toko = st.text_input("Masukkan Kode Toko:", "6809")
        with col_file:
            uploaded_file = st.file_uploader("Upload Laporan ED (.xlsx / .csv)", type=["xlsx", "csv"])
            
        if uploaded_file:
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file, sep=';')
            else:
                df = pd.read_excel(uploaded_file)
                
            st.write("Preview Data:", df.head(3))
            
            # Asumsi kolom di data pusat: 'Dept', 'Article', 'Desc', 'Remark'
            # (Pastikan nama kolom disesuaikan dengan file asli Anda)
            dept_col = next((c for c in df.columns if 'dept' in str(c).lower()), None)
            
            if dept_col:
                unique_depts = df[dept_col].dropna().unique()
                st.markdown("#### Mapping Staff Penanggung Jawab")
                
                mapping_staff = {}
                cols = st.columns(4)
                for i, dept in enumerate(unique_depts):
                    with cols[i % 4]:
                        mapping_staff[dept] = st.text_input(f"Staff {dept}:", key=f"staff_{dept}")
                
                if st.button("🚀 Proses & Bagikan Tugas ke Database"):
                    db_baru = []
                    waktu_skrg = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    
                    for _, row in df.iterrows():
                        dept_name = row[dept_col]
                        staff_name = mapping_staff.get(dept_name, "Belum Ditugaskan")
                        if staff_name.strip() == "":
                            staff_name = "Belum Ditugaskan"
                            
                        # Identifikasi kolom yang relevan
                        art_col = next((c for c in df.columns if 'article' in str(c).lower()), df.columns[0])
                        desc_col = next((c for c in df.columns if 'desc' in str(c).lower()), df.columns[1])
                        remark_col = next((c for c in df.columns if 'remark' in str(c).lower()), None)
                        
                        instruksi = "Periksa / Pindahkan ke Eye-Level"
                        if remark_col and 'markdown' in str(row[remark_col]).lower():
                            instruksi = "Siapkan Label Clearance (Markdown)"
                            
                        db_baru.append({
                            "Waktu_Dibuat": waktu_skrg,
                            "Kode_Toko": kode_toko,
                            "Dept": dept_name,
                            "Nama_Staff": staff_name,
                            "Article": str(row[art_col]),
                            "Nama_Barang": str(row[desc_col]),
                            "Instruksi": instruksi,
                            "Status": "Belum",
                            "Waktu_Selesai": ""
                        })
                    
                    # Simpan ke Database Lokal
                    df_db_baru = pd.DataFrame(db_baru)
                    db_lama = load_db()
                    db_gabungan = pd.concat([db_lama, df_db_baru], ignore_index=True)
                    save_to_db(db_gabungan)
                    st.success("Tugas berhasil didistribusikan ke Scanner Staff!")
            else:
                st.error("Kolom 'Dept' tidak ditemukan di file laporan Anda.")
                
    # TAB 2: CETAK LABEL MARKDOWN
    with tab2:
        st.markdown("Generate File HTML siap cetak khusus untuk barang Clearance (Markdown).")
        
        if uploaded_file and 'df' in locals():
            st.info("Sistem otomatis memfilter barang Markdown dan mencocokkan harga berdasarkan `FILE PRICE ZONE.xlsx`")
            
            # Filter baris yang mengandung kata 'markdown'
            remark_col = next((c for c in df.columns if 'remark' in str(c).lower()), None)
            if remark_col:
                df_markdown = df[df[remark_col].astype(str).str.contains("markdown", case=False, na=False)]
                
                if not df_markdown.empty:
                    col_disc, col_gen = st.columns([1, 2])
                    with col_disc:
                        pilihan_diskon = st.selectbox("Pilih Persentase Diskon:", [50, 75, 90])
                    with col_gen:
                        st.markdown("<br>", unsafe_allow_html=True)
                        if st.button("📄 Generate Label HTML"):
                            html_label = generate_label_html(df_markdown, kode_toko, pilihan_diskon)
                            
                            st.download_button(
                                label="📥 Download File Label (.html)",
                                data=html_label,
                                file_name=f"Label_Clearance_{pilihan_diskon}Persen.html",
                                mime="text/html"
                            )
                            st.caption("Setelah diunduh, klik 2x file HTML-nya, lalu tekan Ctrl + P untuk Print!")
                else:
                    st.warning("Tidak ada barang Markdown di laporan ini.")
            else:
                st.error("Kolom Remark tidak ditemukan.")
        else:
            st.warning("Upload data laporan ED di Tab 1 terlebih dahulu.")

    # TAB 3: MONITORING STAFF
    with tab3:
        st.markdown("Pantau progres kerja harian staff secara Real-Time.")
        db = load_db()
        
        if not db.empty:
            st.write("Database Tugas Keseluruhan:")
            
            # Hitung progress
            total_tugas = len(db)
            tugas_selesai = len(db[db['Status'] == 'Selesai'])
            progress = int((tugas_selesai / total_tugas) * 100) if total_tugas > 0 else 0
            
            st.progress(progress / 100.0, text=f"Progress Keseluruhan: {progress}% ({tugas_selesai} dari {total_tugas} Selesai)")
            
            # Tabel Ringkasan
            st.markdown("#### Detail Status")
            # Tampilkan dengan warna (opsional)
            st.dataframe(db[['Nama_Staff', 'Article', 'Nama_Barang', 'Status', 'Waktu_Selesai']].sort_values(by='Status'))
            
            if st.button("Hapus Semua Riwayat Tugas (Reset)"):
                os.remove(DB_FILE)
                st.rerun()
        else:
            st.info("Belum ada data tugas. Silakan bagikan tugas di Tab 1.")
