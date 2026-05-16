import streamlit as st
import google.generativeai as genai
from docx import Document
from io import BytesIO

# Konfigurasi Halaman Web
st.set_page_config(page_title="Asisten Administrasi Guru", layout="wide")

# Setup API Key di Sidebar
st.sidebar.title("🔑 Pengaturan")
api_key = st.sidebar.text_input("Masukkan Google Gemini API Key:", type="password")

if api_key:
    genai.configure(api_key=api_key)
    # Menggunakan model Gemini terbaru yang stabil dan gratis untuk teks
    model = genai.GenerativeModel('gemini-1.5-flash')
else:
    st.sidebar.warning("Silakan masukkan API Key di sini untuk mengaktifkan AI.")

st.title("🍎 Ruang Kerja Digital Guru (Khusus Pribadi)")

# Fungsi Konversi Teks ke File Word (.docx)
def buat_file_word(teks_konten):
    doc = Document()
    for baris in teks_konten.split('\n'):
        doc.add_paragraph(baris)
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# Fungsi Panggil AI Gemini
def generate_konten(prompt):
    if not api_key:
        return None
    try:
        response = model.generate_content(prompt)
        return response.text
    except Exception as e:
        st.error(f"Terjadi kesalahan: {str(e)}")
        return None

# Menu Fitur Utama
menu = st.sidebar.radio("Pilih Dokumen Admin:", [
    "📝 Modul Ajar (Kurikulum Merdeka)", 
    "📊 Kisi-Kisi & Soal Evaluasi", 
    "📋 Rubrik Penilaian"
])

# --- 1. MODUL AJAR ---
if menu == "📝 Modul Ajar (Kurikulum Merdeka)":
    st.header("Generator Modul Ajar")
    c1, c2 = st.columns(2)
    with c1:
        mapel = st.text_input("Mata Pelajaran:", value="IPA")
        fase = st.selectbox("Fase/Kelas:", ["Fase A (Kelas 1-2)", "Fase B (Kelas 3-4)", "Fase C (Kelas 5-6)", "Fase D (Kelas 7-9)", "Fase E (Kelas 10)", "Fase F (Kelas 11-12)"])
    with c2:
        materi = st.text_input("Topik Pembelajaran:", value="Sistem Pencernaan Manusia")
        waktu = st.text_input("Alokasi Waktu:", value="2 x 45 Menit")

    if st.button("Susun Modul Ajar"):
        prompt = f"""Buatkan Modul Ajar Kurikulum Merdeka lengkap untuk mapel {mapel}, {fase}, materi {materi}, alokasi waktu {waktu}. 
        Struktur wajib: 
        I. INFORMASI UMUM (Identitas, Kompetensi Awal, Profil Pelajar Pancasila, Sarana & Prasarana, Target Peserta Didik).
        II. KOMPONEN INTI (Tujuan Pembelajaran, Pemahaman Bermakna, Pertanyaan Pemantik, Kegiatan Pembelajaran Lengkap: Pendahuluan, Inti, Penutup, Asesmen).
        III. LAMPIRAN (Lembar Kerja Peserta Didik/LKPD singkat, Bahan Bacaan Guru & Peserta Didik)."""
        
        with st.spinner("AI sedang menulis modul ajar Anda..."):
            hasil = generate_konten(prompt)
            if hasil:
                st.markdown(hasil)
                file_word = buat_file_word(hasil)
                st.download_button(
                    label="📥 Unduh Modul Ajar (.docx)",
                    data=file_word,
                    file_name=f"Modul_Ajar_{materi.replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

# --- 2. KISI-KISI & SOAL ---
elif menu == "📊 Kisi-Kisi & Soal Evaluasi":
    st.header("Generator Kisi-Kisi & Soal")
    c1, c2 = st.columns(2)
    with c1:
        mapel = st.text_input("Mata Pelajaran:")
        materi = st.text_input("Cakupan Materi:")
    with c2:
        bentuk = st.selectbox("Bentuk Soal:", ["Pilihan Ganda (5 Opsi)", "Esai/Isian"])
        jumlah = st.slider("Jumlah Soal:", 5, 20, 10)

    if st.button("Buat Kisi-Kisi & Soal"):
        prompt = f"""Buatkan paket evaluasi untuk mapel {mapel} materi {materi} sejumlah {jumlah} soal berbentuk {bentuk}.
        Output harus berisi dua bagian:
        Bagian 1: Tabel Kisi-Kisi Soal (No, Alur Tujuan Pembelajaran, Indikator Soal, Level Kognitif, Nomor Soal).
        Bagian 2: Daftar Soal dan Opsi Jawabannya.
        Bagian 3: Kunci Jawaban beserta Pembahasan singkat di bagian paling akhir."""
        
        with st.spinner("AI sedang merumuskan soal..."):
            hasil = generate_konten(prompt)
            if hasil:
                st.markdown(hasil)
                file_word = buat_file_word(hasil)
                st.download_button(
                    label="📥 Unduh Kisi-Kisi & Soal (.docx)",
                    data=file_word,
                    file_name=f"Soal_{materi.replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

# --- 3. RUBRIK PENILAIAN ---
elif menu == "📋 Rubrik Penilaian":
    st.header("Generator Rubrik Penilaian Tugas")
    tugas = st.text_input("Nama Tugas/Aktivitas:", value="Presentasi Kelompok")
    kriteria = st.text_area("Kriteria yang dinilai:", value="Penguasaan materi, Penyampaian/komunikasi, Kerja sama tim")

    if st.button("Buat Rubrik Penilaian"):
        prompt = f"""Buatkan tabel rubrik penilaian analitik untuk tugas: {tugas}. 
        Kriteria penilaian meliputi: {kriteria}. 
        Format harus berbentuk tabel dengan skala skor 4 (Sangat Baik), 3 (Baik), 2 (Cukup), dan 1 (Perlu Bimbingan) lengkap dengan deskripsi indikator pencapaian di setiap skornya."""
        
        with st.spinner("AI sedang menyusun rubrik..."):
            hasil = generate_konten(prompt)
            if hasil:
                st.markdown(hasil)
                file_word = buat_file_word(hasil)
                st.download_button(
                    label="📥 Unduh Rubrik Penilaian (.docx)",
                    data=file_word,
                    file_name=f"Rubrik_{tugas.replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
