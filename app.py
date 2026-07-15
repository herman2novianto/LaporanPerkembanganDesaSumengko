import streamlit as st
import pandas as pd
from fpdf import FPDF
from datetime import date
import uuidd
from streamlit_gsheets import GSheetsConnection

# ==========================================
# 1. KONFIGURASI KEAMANAN (SECURITY FILTER)
# ==========================================
USERS = {
    "admin.desa@gmail.com": "ALL",
    "email.kades@gmail.com": "ALL",
    "lugaskhalidmaulana@gmail.com": "1",
    "sarihpuspita@gmail.com": "2",
    "assegafsetiawan2018@gmail.com": "3",
    "ramlinuraziza@gmail.com": "4",
    "galuhmawati@gmail.com": "5"
}

import os # Tambahkan ini di bagian paling atas (bersama import pandas dll)

import os

# ==========================================
# 2. DATABASE (Terhubung ke CSV Asli)
# ==========================================
DB_FILE = "data_spp.csv" # Untuk menyimpan input realisasi
MASTER_FILE = "Master_DPA.csv" # File referensi Anda

# Mapping PPKD berdasarkan digit pertama Kode Rekening
# Ini berfungsi menyisipkan Nama_PPKD yang tidak ada di CSV
PEMETAAN_PPKD = {
    "1": "Lugas Khalid Maulana",
    "2": "Puspita Sari Handayani, SE",
    "3": "Agus Setiawan, Spd",
    "4": "Ramli Nur Aziza",
    "5": "Galuh Rismawati"
}

# ==========================================
# FUNGSI PEMBERSIH ANGKA (Letakkan di luar init_db)
# ==========================================
def bersihkan_rupiah(teks):
    """Mengubah format '31.452.000,00' menjadi angka float 31452000.0"""
    try:
        return float(str(teks).replace(".", "").replace(",", "."))
    except:
        return 0.0

# ==========================================
# 2. DATABASE (Terhubung ke CSV Asli)
# ==========================================
conn = st.connection("gsheets", type=GSheetsConnection)

def init_db():
    # 1. BACA FILE MASTER DPA CSV
    if 'master_dpa' not in st.session_state:
        try:
def init_db():
    # 1. BACA FILE MASTER DPA DARI GOOGLE SHEETS
    if 'master_dpa' not in st.session_state:
        # Mengambil sheet Master_DPA
        df_dpa = conn.read(worksheet="Master_DPA", ttl=0)
        
        # Bersihkan data jika ada kolom kosong dari Excel
        df_dpa = df_dpa.dropna(subset=['Kode_Rekening'])
        df_dpa['Kode_Rekening'] = df_dpa['Kode_Rekening'].astype(str)
        
        def get_ppkd_name(kode):
            digit_awal = str(kode)[0]
            return PEMETAAN_PPKD.get(digit_awal, "Tidak Diketahui")
            
        df_dpa['Nama_PPKD'] = df_dpa['Kode_Rekening'].apply(get_ppkd_name)
        st.session_state.master_dpa = df_dpa
            
    # 2. BACA FILE DATA REALISASI (SPP) DARI GOOGLE SHEETS
    if 'realisasi_spp' not in st.session_state:
        df_spp = conn.read(worksheet="Data_SPP", ttl=0)
        
        # Jika sheet masih kosong melompong, buatkan strukturnya
        if df_spp.empty:
            df_spp = pd.DataFrame(columns=[
                "ID", "Tanggal_SPP", "No_SPP", "Kode_Rekening", "Nominal", "Vol_Realisasi", "User_Email"
            ])
            
        st.session_state.realisasi_spp = df_spp

init_db()
            df_dpa.columns = df_dpa.columns.str.strip()
            df_dpa['Kode_Rekening'] = df_dpa['Kode_Rekening'].astype(str)
            
            def get_ppkd_name(kode):
                digit_awal = str(kode)[0]
                return PEMETAAN_PPKD.get(digit_awal, "Tidak Diketahui")
                
            df_dpa['Nama_PPKD'] = df_dpa['Kode_Rekening'].apply(get_ppkd_name)
            st.session_state.master_dpa = df_dpa
            
        except FileNotFoundError:
            st.error(f"File '{MASTER_FILE}' tidak ditemukan.")
            st.stop()
            
    # 2. BACA FILE DATA REALISASI (SPP)
    if os.path.exists(DB_FILE):
        st.session_state.realisasi_spp = pd.read_csv(DB_FILE)
        # Tambahkan kolom ini jika CSV lama belum memilikinya
        if 'Vol_Realisasi' not in st.session_state.realisasi_spp.columns:
            st.session_state.realisasi_spp['Vol_Realisasi'] = 0.0
    else:
        st.session_state.realisasi_spp = pd.DataFrame(columns=[
            "ID", "Tanggal_SPP", "No_SPP", "Kode_Rekening", "Nominal", "Vol_Realisasi", "User_Email"
        ])
        # Jika menyimpan transaksi SPP (Ganti DB_FILE):
conn.update(worksheet="Data_SPP", data=st.session_state.realisasi_spp)
st.cache_data.clear() # Bersihkan cache agar langsung update

init_db()
# ==========================================
# 3. FUNGSI GENERATE PDF (DIPERBARUI)
# ==========================================
class PDF(FPDF):
    def __init__(self, orientation='P', unit='mm', format='A4', judul=""):
        super().__init__(orientation, unit, format)
        self.judul = judul  # Menyimpan judul yang dikirim

    def header(self):
        self.set_font("Arial", "B", 11)
        self.cell(0, 5, self.judul, align="C", ln=1)
        self.ln(2)

def generate_laporan_pdf(df_laporan, bulan, tahun, desa, kecamatan, kabupaten, provinsi, nama_ppkd, judul_laporan):
    pdf = PDF(orientation="L", judul=judul_laporan) # Mengirim judul ke header
    pdf.add_page()
    
    # KOP INFO DAERAH
    pdf.set_font("Arial", "B", 10)
    pdf.cell(0, 5, f"Bulan : {bulan}                    Tahun : {tahun}", align="C", ln=1)
    pdf.ln(3)
    
    pdf.set_font("Arial", "", 9)
    pdf.cell(40, 5, "DESA", ln=0); pdf.cell(100, 5, f": {desa}", ln=1)
    pdf.cell(40, 5, "KECAMATAN", ln=0); pdf.cell(100, 5, f": {kecamatan}", ln=1)
    pdf.cell(40, 5, "KABUPATEN", ln=0); pdf.cell(100, 5, f": {kabupaten}", ln=1)
    pdf.cell(40, 5, "PROVINSI", ln=0); pdf.cell(100, 5, f": {provinsi}", ln=1)
    pdf.ln(3)
    
    # ================= HEADER TABEL BERTINGKAT =================
    pdf.set_font("Arial", "B", 7)
    
    # BARIS 1
    pdf.cell(28, 5, "", border="LTR", align="C")
    pdf.cell(65, 5, "", border="LTR", align="C")
    pdf.cell(40 + 58, 5, "OUTPUT", border=1, align="C")
    pdf.cell(68, 5, "SUMBER DANA", border=1, align="C", ln=1)
    
    # BARIS 2
    pdf.cell(28, 5, "KODE REKENING", border="LR", align="L")
    pdf.cell(65, 5, "URAIAN", border="LR", align="C")
    pdf.cell(40, 5, "Rencana", border=1, align="C")
    pdf.cell(58, 5, "Realisasi Sampai Saat ini", border=1, align="C")
    pdf.cell(17, 5, "Dana Desa", border="LTR", align="C")
    pdf.cell(17, 5, "Alokasi Dana", border="LTR", align="C")
    pdf.cell(17, 5, "Lain-Lain", border="LTR", align="C")
    pdf.cell(17, 5, "Bentuk Lain", border="LTR", align="C", ln=1)
    
    # BARIS 3
    pdf.cell(28, 5, "", border="LBR", align="C")
    pdf.cell(65, 5, "", border="LBR", align="C")
    
    # Rencana
    pdf.cell(10, 5, "Vol", border=1, align="C")
    pdf.cell(10, 5, "Sat", border=1, align="C")
    pdf.cell(20, 5, "Anggaran", border=1, align="C")
    
    # Realisasi
    pdf.cell(10, 5, "Vol", border=1, align="C")
    pdf.cell(10, 5, "Sat", border=1, align="C")
    pdf.cell(20, 5, "Anggaran", border=1, align="C")
    pdf.cell(18, 5, "Capaian (%)", border=1, align="C")
    
    # Sumber Dana
    pdf.cell(17, 5, "(Rp)", border="LBR", align="C")
    pdf.cell(17, 5, "Desa (Rp)", border="LBR", align="C")
    pdf.cell(17, 5, "(Rp)", border="LBR", align="C")
    pdf.cell(17, 5, "", border="LBR", align="C", ln=1)
    
    # BARIS NOMOR KOLOM (1 - 14)
    pdf.set_font("Arial", "", 6)
    widths = [28, 65, 10, 10, 20, 10, 10, 20, 18, 17, 17, 17, 17]
    for i, w in enumerate(widths):
        pdf.cell(w, 4, str(i+1), border=1, align="C")
    pdf.ln()
    
    # ================= ISI TABEL =================
    pdf.set_font("Arial", "", 7)
    total_angg_ren = 0; total_angg_real = 0
    total_dd = 0; total_add = 0; total_lain = 0
    
    for idx, row in df_laporan.iterrows():
        # Parsing data master dan SPP
        uraian = str(row['Uraian_Kegiatan'])
        if len(uraian) > 35: uraian = uraian[:55] + "..."
            
        angg_rencana = bersihkan_rupiah(row['Anggaran_Rencana'])
        angg_real = float(row['Nominal']) # Didapat dari penjumlahan SPP
        
        # Mengambil Volume Realisasi
        vol_real = float(row['Vol_Realisasi']) if pd.notna(row['Vol_Realisasi']) else 0
        
        capaian = (angg_real / angg_rencana * 100) if angg_rencana > 0 else 0
        sumber = str(row['Sumber_Dana']).strip().upper()
        
        # PERBAIKAN LOGIKA: Sekarang mengambil angka dari Realisasi (angg_real)
        dd = angg_real if sumber == "DDS" else 0
        add = angg_real if sumber == "ADD" else 0
        lain = angg_real if sumber not in ["DDS", "ADD", ""] else 0
        
        # Cetak ke PDF
        pdf.cell(widths[0], 6, str(row['Kode_Rekening']), border=1, align="C")
        pdf.cell(widths[1], 6, uraian, border=1)
        
        # Rencana
        pdf.cell(widths[2], 6, str(row['Vol_Rencana']), border=1, align="C")
        pdf.cell(widths[3], 6, str(row['Satuan_Rencana']), border=1, align="C")
        pdf.cell(widths[4], 6, f"{angg_rencana:,.0f}", border=1, align="R")
        
        # Realisasi (Volume sekarang akan muncul)
        pdf.cell(widths[5], 6, f"{vol_real:g}" if vol_real > 0 else "-", border=1, align="C") 
        pdf.cell(widths[6], 6, str(row['Satuan_Rencana']), border=1, align="C")
        pdf.cell(widths[7], 6, f"{angg_real:,.0f}", border=1, align="R")
        pdf.cell(widths[8], 6, f"{capaian:.1f}%", border=1, align="C")
        
        # Sumber Dana (Memunculkan nilai Realisasi)
        pdf.cell(widths[9], 6, f"{dd:,.0f}" if dd>0 else "-", border=1, align="R")
        pdf.cell(widths[10], 6, f"{add:,.0f}" if add>0 else "-", border=1, align="R")
        pdf.cell(widths[11], 6, f"{lain:,.0f}" if lain>0 else "-", border=1, align="R")
        pdf.cell(widths[12], 6, str(row['Sumber_Dana']), border=1, align="C") # Bentuk Lain
        pdf.ln()
        
        # Akumulasi
        total_angg_ren += angg_rencana; total_angg_real += angg_real
        total_dd += dd; total_add += add; total_lain += lain
        
    # FOOTER JUMLAH
    pdf.set_font("Arial", "B", 7)
    pdf.cell(widths[0]+widths[1], 6, "Jumlah", border=1, align="R")
    pdf.cell(widths[2]+widths[3], 6, "", border=1)
    pdf.cell(widths[4], 6, f"{total_angg_ren:,.0f}", border=1, align="R")
    pdf.cell(widths[5]+widths[6], 6, "", border=1)
    pdf.cell(widths[7], 6, f"{total_angg_real:,.0f}", border=1, align="R")
    pdf.cell(widths[8], 6, "", border=1)
    pdf.cell(widths[9], 6, f"{total_dd:,.0f}", border=1, align="R")
    pdf.cell(widths[10], 6, f"{total_add:,.0f}", border=1, align="R")
    pdf.cell(widths[11], 6, f"{total_lain:,.0f}", border=1, align="R")
    pdf.cell(widths[12], 6, "", border=1)
    pdf.ln(15)
    
    # TANDA TANGAN
    pdf.set_x(180)
    pdf.cell(60, 5, "Kaur/Kasi", align="C", ln=1)
    pdf.ln(15)
    pdf.set_x(180)
    pdf.cell(60, 5, f"( {nama_ppkd} )", align="C", ln=1)
    
    filename = f"Laporan_Perkembangan_{tgl_mulai.strftime('%d%m%Y')}_{tgl_akhir.strftime('%d%m%Y')}.pdf"
    pdf.output(filename)
    with open(filename, "rb") as f:
        return f.read(), filename
# ==========================================
# 4. ANTARMUKA PENGGUNA (UI)
# ==========================================
st.set_page_config(page_title="Keuangan Desa", layout="wide")
st.title("Sistem Tata Kelola Keuangan Desa")

# SIMULASI LOGIN (Pengganti USEREMAIL() di AppSheet)
user_email = st.sidebar.selectbox("Simulasi Login (Security Filter):", list(USERS.keys()))
hak_akses = USERS[user_email]

if hak_akses == "ALL":
    st.sidebar.info("Akses: Admin / Kades (Melihat Semua Data)")
else:
    st.sidebar.success(f"Akses: Bidang {hak_akses}")

# Relasi Data (Ref / Lookup) Master DPA
df_dpa = st.session_state.master_dpa
if hak_akses != "ALL":
    # Sama persis dengan logika: LEFT([Kode_Rekening], 1) = Kunci Akses
    df_dpa = df_dpa[df_dpa['Kode_Rekening'].str.startswith(hak_akses)]

# --- TAMBAHKAN MENU NAVIGASI SIDEBAR ---
st.sidebar.markdown("---")
if hak_akses == "ALL":
    # Hanya Admin yang bisa melihat menu Kelola Master DPA
    menu_utama = st.sidebar.radio("📂 MENU UTAMA", ["📝 Transaksi & Laporan", "⚙️ Kelola Master DPA"])
else:
    menu_utama = "📝 Transaksi & Laporan"
st.sidebar.markdown("---")

# Membuat kamus/peta untuk menggabungkan Kode dan Uraian
uraian_dict = dict(zip(df_dpa['Kode_Rekening'], df_dpa['Uraian_Kegiatan']))
sdana_dict = dict(zip(df_dpa['Kode_Rekening'], df_dpa['Sumber_Dana']))

def format_dropdown(kode):
    # Logika baru untuk merespon pilihan "ALL"
    if kode == "ALL":
        return "☑️ ALL - Cetak Semua Kegiatan"
    # Menggabungkan Kode_Rekening dengan Uraian_Kegiatan
    uraian = uraian_dict.get(kode, "")
    sdana = sdana_dict.get(kode, "")
    return f"{kode} - {uraian} - {sdana}"


# ==========================================
# PERCABANGAN MENU
# ==========================================
if menu_utama == "⚙️ Kelola Master DPA":
    st.subheader("⚙️ Kelola Data Master DPA")
    st.info("💡 **Petunjuk:** Klik 2x pada sel tabel untuk mengedit teks/angka. Gunakan tombol '➕ Add Row' di bagian bawah tabel untuk menambah kegiatan baru. Untuk menghapus, centang kotak di ujung kiri baris, lalu tekan ikon '🗑️' (Delete).")
    
    # Menampilkan Data Editor Interaktif
    edited_dpa = st.data_editor(
        st.session_state.master_dpa, 
        num_rows="dynamic", # Memungkinkan penambahan & penghapusan baris
        use_container_width=True,
        key="dpa_editor"
    )
    
    # Tombol Simpan Perubahan
    if st.button("💾 Simpan Perubahan ke CSV", type="primary"):
        try:
            # Update session_state dengan data yang baru diedit
            st.session_state.master_dpa = edited_dpa
            # Simpan secara permanen ke file CSV
            # Jika menyimpan transaksi SPP (Ganti DB_FILE):
conn.update(worksheet="Data_SPP", data=st.session_state.realisasi_spp)
st.cache_data.clear() # Bersihkan cache agar langsung update
            st.success("✅ Data Master DPA berhasil diperbarui dan disimpan secara permanen!")
            st.rerun() # Muat ulang agar kamus dictionary (uraian_dict) ikut ter-update
        except PermissionError:
            st.error(f"⚠️ Gagal Menyimpan: File '{MASTER_FILE}' sedang terbuka di aplikasi lain (seperti Excel). Silakan tutup dulu file tersebut.")

elif menu_utama == "📝 Transaksi & Laporan":
    
    tab1, tab2, tab3 , tab4 = st.tabs(["📝 Input Realisasi SPP", "🗂️ Data SPP Sudah Input", "🖨️ Print Lap Perkembangan", "🖨️ Print Lap Akhir"])

    # --- TAB 1: INPUT DATA ---
    with tab1:
        st.subheader("Input Data Realisasi SPP Baru")
        
        if df_dpa.empty:
            st.warning("Tidak ada data kegiatan untuk bidang Anda.")
        else:
            kode = st.selectbox(
                "Pilih Kode Rekening", 
                df_dpa['Kode_Rekening'].tolist(),
                format_func=format_dropdown
            )
            
            detail_kegiatan = df_dpa[df_dpa['Kode_Rekening'] == kode].iloc[0]
            vol = detail_kegiatan['Vol_Rencana']
            satuan = detail_kegiatan['Satuan_Rencana']
            anggaran = detail_kegiatan['Anggaran_Rencana']
            sumber = detail_kegiatan['Sumber_Dana']
            
            st.info(f"📌 **Detail Perencanaan:** Volume {vol} {satuan} | Pagu Anggaran: Rp {anggaran} | Sumber Dana: {sumber}")
            
            with st.form("form_input", clear_on_submit=True):
                col1, col2 = st.columns(2)
                tgl = col1.date_input("Tanggal SPP")
                no_spp = col2.text_input("Nomor SPP (Misal: 001/SPP/2026)")
                
                col_vol, col_nom = st.columns(2)
                vol_realisasi = col_vol.number_input("Volume Realisasi", min_value=0.0, step=1.0)
                nominal = col_nom.number_input("Nominal (Rp)", min_value=0, step=50000)
                
                if st.form_submit_button("Simpan Data (Save)"):
                    if nominal > 0:
                        data_baru = pd.DataFrame([{
                            "ID": str(uuid.uuid4())[:8],
                            "Tanggal_SPP": tgl, "No_SPP": no_spp, 
                            "Kode_Rekening": kode, "Nominal": nominal, 
                            "Vol_Realisasi": vol_realisasi, 
                            "User_Email": user_email
                        }])
                        
                        st.session_state.realisasi_spp = pd.concat(
                            [st.session_state.realisasi_spp, data_baru], 
                            ignore_index=True
                        )
                        st.session_state.realisasi_spp.to_csv(DB_FILE, index=False)
                        st.success(f"Data SPP untuk {kode} berhasil disimpan permanen!")
                    else:
                        st.error("Nominal tidak boleh 0.")

    # --- TAB 2: DATA & CETAK SATUAN ---
    with tab2:
        st.subheader("Database Realisasi SPP")
        df_spp = st.session_state.realisasi_spp
        
        if hak_akses != "ALL":
            df_spp = df_spp[df_spp['Kode_Rekening'].str.startswith(hak_akses)]
        
        if df_spp.empty:
            st.info("Belum ada data SPP yang diinput.")
        else:
            st.write("Daftar transaksi yang telah diinput:")
            
            for idx, row in df_spp.iterrows():
                uraian = uraian_dict.get(row['Kode_Rekening'], "Uraian tidak diketahui")
                vol = row['Vol_Realisasi'] if pd.notna(row['Vol_Realisasi']) else 0
                
                with st.container():
                    col1, col2, col3 = st.columns([1, 4, 1])
                    col1.write(f"📅 **{row['Tanggal_SPP']}**")
                    
                    teks_detail = f"**{row['No_SPP']}** | {row['Kode_Rekening']}\n\n"
                    teks_detail += f"*{uraian}*\n\n"
                    teks_detail += f"📦 **Vol:** {vol:g} &nbsp;&nbsp;|&nbsp;&nbsp; 💰 **Nominal:** Rp {int(row['Nominal']):,}"
                    col2.write(teks_detail)
                    
                    if col3.button("🗑️ Hapus", key=f"del_{row['ID']}", type="secondary"):
                        st.session_state.realisasi_spp = st.session_state.realisasi_spp[st.session_state.realisasi_spp['ID'] != row['ID']]
                        st.session_state.realisasi_spp.to_csv(DB_FILE, index=False)
                        st.rerun()
                st.divider()

    # --- TAB 3: PRINT LAP PERKEMBANGAN SESUAI FORMAT GAMBAR ---
    with tab3:
        st.subheader("Cetak Laporan Perkembangan Pelaksanaan Kegiatan")
        
        if 'master_dpa' not in st.session_state or st.session_state.master_dpa.empty:
            st.warning("⚠️ Data Master DPA belum tersedia. Pastikan file Master_DPA.csv sudah benar.")
        else:
            with st.expander("📝 Pengaturan Header & Waktu", expanded=True):
                rc1, rc2 = st.columns(2)
                pilihan_bulan = rc1.selectbox("Pilih Bulan", ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"])
                pilihan_tahun = rc2.number_input("Tahun", value=2026, step=1)
                tgl_mulai = rc1.date_input("Mulai Tanggal (SPP)", date(2026, 1, 1))
                tgl_akhir = rc2.date_input("Sampai Tanggal (SPP)", date(2026, 1, 31))
                
                c1, c2, c3, c4 = st.columns(4)
                nama_desa = c1.text_input("Nama Desa", "Sumengko")
                nama_kec = c2.text_input("Kecamatan", "Kwadungan")
                nama_kab = c3.text_input("Kabupaten", "Ngawi")
                nama_prov = c4.text_input("Provinsi", "Jawa Timur")
            
            pilihan_kegiatan = ["ALL"] + df_dpa['Kode_Rekening'].tolist()
            kode_trigger = st.multiselect(
                "Pilih Kegiatan yang akan dicetak (Bisa pilih ALL atau tambah satu per satu):", 
                options=pilihan_kegiatan,
                default=["ALL"],
                format_func=format_dropdown
            )
            
            if st.button("🖨️ Generate Laporan PDF (Format Perbup)"):
                if not kode_trigger:
                    st.warning("Silakan pilih minimal satu kegiatan untuk dicetak!")
                else:
                    if "ALL" in kode_trigger:
                        df_laporan_bidang = df_dpa.copy()
                    else:
                        df_laporan_bidang = df_dpa[df_dpa['Kode_Rekening'].isin(kode_trigger)].copy() 
                    
                    nama_ppkd_cetak = df_laporan_bidang['Nama_PPKD'].iloc[0] if not df_laporan_bidang.empty else "Ttd"
                
                    spp_all = st.session_state.realisasi_spp.copy()
                    
                    if not spp_all.empty:
                        spp_all['Tanggal_SPP'] = pd.to_datetime(spp_all['Tanggal_SPP']).dt.date
                        spp_filter_waktu = spp_all[(spp_all['Tanggal_SPP'] >= tgl_mulai) & (spp_all['Tanggal_SPP'] <= tgl_akhir)]
                        spp_grouped = spp_filter_waktu.groupby("Kode_Rekening")[["Nominal", "Vol_Realisasi"]].sum().reset_index()
                    else:
                        spp_grouped = pd.DataFrame(columns=["Kode_Rekening", "Nominal", "Vol_Realisasi"])
                    
                    df_final = pd.merge(df_laporan_bidang, spp_grouped, on="Kode_Rekening", how="left")
                    df_final['Nominal'] = df_final['Nominal'].fillna(0)
                    df_final['Vol_Realisasi'] = df_final['Vol_Realisasi'].fillna(0)
                    
                    if df_final.empty:
                        st.warning("Tidak ada data DPA untuk bidang tersebut.")
                    else:
                        pdf_bytes, nama_file = generate_laporan_pdf(
                            df_final, pilihan_bulan, pilihan_tahun, 
                            nama_desa, nama_kec, nama_kab, nama_prov, nama_ppkd_cetak, "LAPORAN PERKEMBANGAN PELAKSANAAN KEGIATAN DAN ANGGARAN"
                        )
                        
                        st.download_button(
                            label=f"⬇️ Download Laporan PDF",
                            data=pdf_bytes,
                            file_name=nama_file,
                            mime="application/pdf",
                            type="primary"
                        )
                        st.success(f"Tabel Laporan dari {tgl_mulai.strftime('%d-%m-%Y')} s/d {tgl_akhir.strftime('%d-%m-%Y')} berhasil di-generate!")

    # --- TAB 4: PRINT LAP AKHIR SESUAI FORMAT GAMBAR ---
    with tab4:
        st.subheader("Cetak Laporan Akhir Pelaksanaan Kegiatan")
        
        if 'master_dpa' not in st.session_state or st.session_state.master_dpa.empty:
            st.warning("⚠️ Data Master DPA belum tersedia. Pastikan file Master_DPA.csv sudah benar.")
        else:
            with st.expander("📝 Pengaturan Header & Waktu", expanded=True):
                rc1, rc2 = st.columns(2)
                pilihan_bulan_akhir = rc1.selectbox("Pilih Bulan", ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"], key="bulan_tab4")
                pilihan_tahun_akhir = rc2.number_input("Tahun", value=2026, step=1, key="tahun_tab4")
                tgl_mulai_akhir = rc1.date_input("Mulai Tanggal (SPP)", date(2026, 1, 1), key="tgl_mulai_tab4")
                tgl_akhir_akhir = rc2.date_input("Sampai Tanggal (SPP)", date(2026, 1, 31), key="tgl_akhir_tab4")
                
                c1, c2, c3, c4 = st.columns(4)
                nama_desa_akhir = c1.text_input("Nama Desa", "Sumengko", key="desa_tab4")
                nama_kec_akhir = c2.text_input("Kecamatan", "Kwadungan", key="kec_tab4")
                nama_kab_akhir = c3.text_input("Kabupaten", "Ngawi", key="kab_tab4")
                nama_prov_akhir = c4.text_input("Provinsi", "Jawa Timur", key="prov_tab4")
            
            pilihan_kegiatan_akhir = ["ALL"] + df_dpa['Kode_Rekening'].tolist()
            kode_trigger_akhir = st.multiselect(
                "Pilih Kegiatan yang akan dicetak (Bisa pilih ALL atau tambah satu per satu):", 
                options=pilihan_kegiatan_akhir,
                default=["ALL"],
                format_func=format_dropdown,
                key="bidang_tab4"
            )
            
            if st.button("🖨️ Generate Laporan Akhir PDF", key="btn_gen_tab4"):
                if not kode_trigger_akhir:
                    st.warning("Silakan pilih minimal satu kegiatan untuk dicetak!")
                else:
                    if "ALL" in kode_trigger_akhir:
                        df_laporan_bidang = df_dpa.copy()
                    else:
                        df_laporan_bidang = df_dpa[df_dpa['Kode_Rekening'].isin(kode_trigger_akhir)].copy()
                    
                    nama_ppkd_cetak = df_laporan_bidang['Nama_PPKD'].iloc[0] if not df_laporan_bidang.empty else "Ttd"
                
                    spp_all = st.session_state.realisasi_spp.copy()
                    
                    if not spp_all.empty:
                        spp_all['Tanggal_SPP'] = pd.to_datetime(spp_all['Tanggal_SPP']).dt.date
                        spp_filter_waktu = spp_all[(spp_all['Tanggal_SPP'] >= tgl_mulai_akhir) & (spp_all['Tanggal_SPP'] <= tgl_akhir_akhir)]
                        spp_grouped = spp_filter_waktu.groupby("Kode_Rekening")[["Nominal", "Vol_Realisasi"]].sum().reset_index()
                    else:
                        spp_grouped = pd.DataFrame(columns=["Kode_Rekening", "Nominal", "Vol_Realisasi"])
                    
                    df_final = pd.merge(df_laporan_bidang, spp_grouped, on="Kode_Rekening", how="left")
                    df_final['Nominal'] = df_final['Nominal'].fillna(0)
                    df_final['Vol_Realisasi'] = df_final['Vol_Realisasi'].fillna(0)
                    
                    if df_final.empty:
                        st.warning("Tidak ada data DPA untuk bidang tersebut.")
                    else:
                        pdf_bytes, nama_file = generate_laporan_pdf(
                            df_final, pilihan_bulan_akhir, pilihan_tahun_akhir, 
                            nama_desa_akhir, nama_kec_akhir, nama_kab_akhir, nama_prov_akhir, nama_ppkd_cetak, "LAPORAN AKHIR PELAKSANAAN KEGIATAN DAN ANGGARAN"
                        )
                        
                        st.download_button(
                            label=f"⬇️ Download Laporan Akhir PDF",
                            data=pdf_bytes,
                            file_name="Akhir_" + nama_file,
                            mime="application/pdf",
                            type="primary",
                            key="btn_dl_tab4"
                        )
                        st.success(f"Tabel Laporan Akhir dari {tgl_mulai_akhir.strftime('%d-%m-%Y')} s/d {tgl_akhir_akhir.strftime('%d-%m-%Y')} berhasil di-generate!")