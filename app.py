import streamlit as st
import google.generativeai as genai
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls
from io import BytesIO
import time
import re

# --- 1. KONFIGURASI HALAMAN WEB STREAMLIT ---
st.set_page_config(page_title="Sistem Perangkat Pembelajaran MAN 2 Kota Makassar", layout="wide")

# --- 2. SETUP API KEY & NAVIGASI DI SIDEBAR ---
st.sidebar.title("🔑 Pengaturan & Navigasi")
api_key = st.sidebar.text_input("Masukkan Google Gemini API Key:", type="password")

if api_key:
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel('gemini-2.5-flash')
else:
    st.sidebar.warning("Silakan masukkan API Key di sini untuk mengaktifkan AI.")

st.sidebar.markdown("---")
st.sidebar.subheader("📌 Pilih Modul Kerja")

# Navigasi Menu Utama di Sidebar
menu_utama = st.sidebar.radio(
    "Silakan pilih layanan:",
    ["📝 Menyusun RPP", "📊 Menyusun Kisi-Kisi", "❓ Membuat Soal"]
)

st.sidebar.markdown("---")
st.sidebar.subheader("📄 Identitas Dokumen")
nama_guru = st.sidebar.text_input("Nama Guru:", value="Kaharuddin, S.Pd")
mapel = st.sidebar.text_input("Mata Pelajaran:", value="PPKn")
topik_pokok = st.sidebar.text_input("Topik Pokok:", value="Pancasila dalam Kehidupan Global")

# --- 3. JUDUL UTAMA APLIKASI ---
st.title("🍎 Sistem Perencanaan & Evaluasi Pembelajaran (KBC)")
st.caption("Edisi Eco-Friendly (Hemat Kertas): MAN 2 Kota Makassar")

# --- 4. INISIALISASI SESSION STATE DATA ---
if 'rpp_data' not in st.session_state:
    st.session_state.rpp_data = {
        "bagian_awal": "", "kegiatan_inti": "", "asesmen": "", "lkpd": "", "bahan_bacaan": "", "rubrik": ""
    }
if 'kisi_data' not in st.session_state:
    st.session_state.kisi_data = {"konten": ""}

if 'soal_data' not in st.session_state:
    st.session_state.soal_data = {"konten": ""}


# --- 5. FUNGSI CORE AI ---
def panggil_ai(prompt):
    if not api_key:
        st.error("Masukkan API Key terlebih dahulu di sidebar kiri!")
        return ""
    try:
        response = model.generate_content(prompt)
        return response.text
    except Exception as e:
        if "429" in str(e) or "quota" in str(e).lower():
            st.warning("⚠️ Kuota menit ini habis. Mohon tunggu 60 detik...")
            time.sleep(5)
        else:
            st.error(f"Terjadi kesalahan: {str(e)}")
        return ""

# --- 6. FUNGSI FORMAT WORD EKSPOR ---
def set_cell_background(cell, color_hex):
    shading_xml = f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>'
    cell._tc.get_or_add_tcPr().append(parse_xml(shading_xml))

def format_cell_margins(cell, top=60, bottom=60, left=100, right=100):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for margin, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{margin}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def pasang_garis_pembatas_tabel(table, color_hex="D1D5DB"):
    tblPr = table._tbl.tblPr
    borders_xml = f"""
    <w:tblBorders {nsdecls("w")}>
        <w:top w:val="single" w:sz="4" w:space="0" w:color="{color_hex}"/>
        <w:bottom w:val="single" w:sz="6" w:space="0" w:color="{color_hex}"/>
        <w:insideH w:val="single" w:sz="4" w:space="0" w:color="{color_hex}"/>
        <w:left w:val="none"/>
        <w:right w:val="none"/>
        <w:insideV w:val="none"/>
    </w:tblBorders>
    """
    tblPr.append(parse_xml(borders_xml))

def buat_dokumen_word(judul_dokumen, komponen_list):
    doc = Document()
    
    for section in doc.sections:
        section.page_width = Inches(8.27)
        section.page_height = Inches(11.69)
        section.top_margin = Inches(0.6)
        section.bottom_margin = Inches(0.6)
        section.left_margin = Inches(0.6)
        section.right_margin = Inches(0.6)

    header_p = doc.add_paragraph()
    header_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h_run = header_p.add_run(f"{judul_dokumen}\n")
    h_run.font.name = 'Georgia'
    h_run.font.size = Pt(13)
    h_run.font.bold = True
    h_run.font.color.rgb = RGBColor(26, 54, 93)
    
    sub_run = header_p.add_run("MAN 2 Kota Makassar — Kurikulum Berbasis Cinta (KBC)\n")
    sub_run.font.name = 'Arial'
    sub_run.font.size = Pt(9.5)
    sub_run.font.italic = True
    sub_run.font.color.rgb = RGBColor(100, 110, 125)

    meta_table = doc.add_table(rows=2, cols=4)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    pasang_garis_pembatas_tabel(meta_table, "A0AEC0")
    
    identitas_data = [
        ("Nama Guru", f": {nama_guru}", "Mata Pelajaran", f": {mapel}"),
        ("Fase/Kelas", f": F / XII", "Topik Pokok", f": {topik_pokok}")
    ]
    for i, (k1, v1, k2, v2) in enumerate(identitas_data):
        row = meta_table.rows[i]
        row.cells[0].text = k1
        row.cells[1].text = v1
        row.cells[2].text = k2
        row.cells[3].text = v2
        for cell in row.cells:
            format_cell_margins(cell, top=40, bottom=40, left=80, right=80)
            for p in cell.paragraphs:
                for r in p.runs:
                    r.font.name = 'Arial'
                    r.font.size = Pt(9.5)
                    r.font.color.rgb = RGBColor(45, 55, 72)

    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    for judul, isi in komponen_list:
        if isi:
            if judul:
                h = doc.add_heading(level=1)
                hrun = h.add_run(judul)
                hrun.font.name = 'Georgia'
                hrun.font.size = Pt(11.5)
                hrun.font.bold = True
                hrun.font.color.rgb = RGBColor(26, 54, 93)
                h.paragraph_format.space_before = Pt(8)
                h.paragraph_format.space_after = Pt(4)
            
            baris_list = isi.split('\n')
            di_dalam_tabel = False
            tabel_obj = None
            
            for baris in baris_list:
                baris_bersih = baris.strip()
                if not baris_bersih:
                    continue
                
                if baris_bersih.startswith('|'):
                    kolom_data = [k.strip() for k in baris_bersih.split('|')[1:-1]]
                    if not kolom_data or all(c == '' or c.startswith('-') for c in kolom_data):
                        continue 
                    
                    if not di_dalam_tabel:
                        di_dalam_tabel = True
                        tabel_obj = doc.add_table(rows=0, cols=len(kolom_data))
                        tabel_obj.alignment = WD_TABLE_ALIGNMENT.CENTER
                        pasang_garis_pembatas_tabel(tabel_obj)
                        
                        row = tabel_obj.add_row()
                        for idx, teks in enumerate(kolom_data):
                            cell = row.cells[idx]
                            cell.text = teks
                            set_cell_background(cell, "1A365D")
                            format_cell_margins(cell, top=70, bottom=70, left=90, right=90)
                            for p_cell in cell.paragraphs:
                                p_cell.alignment = WD_ALIGN_PARAGRAPH.CENTER
                                if p_cell.runs:
                                    p_cell.runs[0].font.bold = True
                                    p_cell.runs[0].font.color.rgb = RGBColor(255, 255, 255)
                                    p_cell.runs[0].font.name = 'Arial'
                        continue
                    
                    row = tabel_obj.add_row()
                    for idx, teks in enumerate(kolom_data):
                        cell = row.cells[idx]
                        cell.text = teks
                        format_cell_margins(cell, top=70, bottom=70, left=90, right=90)
                else:
                    di_dalam_tabel = False
                    p = doc.add_paragraph()
                    p.paragraph_format.space_after = Pt(2)
                    p.paragraph_format.line_spacing = 1.05
                    run = p.add_run(baris_bersih)
                    run.font.name = 'Arial'
                    run.font.size = Pt(10)
                    run.font.color.rgb = RGBColor(45, 55, 72)
                    
    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output


# ==========================================
# HALAMAN UTAMA BERDASARKAN PILIHAN MENU
# ==========================================

# --- MENU 1: MENYUSUN RPP ---
if menu_utama == "📝 Menyusun RPP":
    st.header("Modul Pembuatan RPP Padat & Mendalam")
    
    # 1. MENU INPUT (Dropdown Komponen RPP seperti semula)
    pilihan_bagian = st.selectbox(
        "Pilih Komponen RPP yang ingin dikerjakan:",
        ["Bagian Awal & Kompetensi", "Skenario Kegiatan Inti", "Asesmen", "LKPD", "Bahan Bacaan", "Rubrik Penilaian"]
    )
    
    map_key_rpp = {
        "Bagian Awal & Kompetensi": "bagian_awal", "Skenario Kegiatan Inti": "kegiatan_inti",
        "Asesmen": "asesmen", "LKPD": "lkpd", "Bahan Bacaan": "bahan_bacaan", "Rubrik Penilaian": "rubrik"
    }
    key_rpp = map_key_rpp[pilihan_bagian]
    
    # Kolom Input Tambahan untuk Instruksi Khusus Guru (Custom Prompt)
    input_permintaan_rpp = st.text_input(
        f"Masukkan instruksi tambahan / sub-materi khusus untuk {pilihan_bagian} (Opsional):",
        placeholder="Contoh: Fokuskan pada sila ke-3 Pancasila dan hubungkan dengan aksi pelestarian lingkungan mangrove."
    )
    
    # Tampilan Output Utama Teks Area RPP
    st.session_state.rpp_data[key_rpp] = st.text_area(
        f"Hasil Konten Teks {pilihan_bagian}:", value=st.session_state.rpp_data[key_rpp], height=300
    )
    
    if st.button(f"✨ Generate {pilihan_bagian} dengan Gemini"):
        with st.spinner("AI sedang merumuskan materi RPP..."):
            prompt_dasar = f"Buatlah isi untuk modul {pilihan_bagian} tentang materi '{topik_pokok}' mata pelajaran {mapel} untuk tingkat kelas XII berdasarkan prinsip Kurikulum Berbasis Cinta (KBC). Buat dalam format padat hemat kertas cetak."
            if input_permintaan_rpp:
                prompt_dasar += f" Harap sesuaikan juga dengan permintaan guru berikut: {input_permintaan_rpp}"
            
            hasil = panggil_ai(prompt_dasar)
            if hasil:
                st.session_state.rpp_data[key_rpp] = hasil
                st.success("Berhasil di-generate! Data otomatis ter-update di kolom teks atas.")
                st.rerun()
                
    st.markdown("---")
    st.subheader("📥 Cetak Dokumen RPP")
    rpp_components = [
        ("I. BAGIAN AWAL & ANALISIS KOMPETENSI", st.session_state.rpp_data["bagian_awal"]),
        ("II. SKENARIO KEGIATAN INTI", st.session_state.rpp_data["kegiatan_inti"]),
        ("III. ASESMEN PEMBELAJARAN MENDALAM", st.session_state.rpp_data["asesmen"]),
        ("IV. LEMBAR KERJA PESERTA DIDIK (LKPD)", st.session_state.rpp_data["lkpd"]),
        ("V. BAHAN BACAAN GURU & SISWA", st.session_state.rpp_data["bahan_bacaan"]),
        ("VI. RUBRIK PENILAIAN OTENTIK SISWA", st.session_state.rpp_data["rubrik"])
    ]
    
    if any(v for k, v in st.session_state.rpp_data.items() if v):
        file_doc = buat_dokumen_word("PERENCANAAN PEMBELAJARAN MENDALAM (RPP / MODUL AJAR)", rpp_components)
        st.download_button(
            label="📥 Unduh File RPP (.docx)", 
            data=file_doc, 
            file_name=f"RPP_Eco_{topik_pokok.replace(' ', '_')}.docx"
        )
    else:
        st.info("Silakan generate isi RPP terlebih dahulu agar tombol unduh aktif.")


# --- MENU 2: MENYUSUN KISI-KISI ---
elif menu_utama == "📊 Menyusun Kisi-Kisi":
    st.header("Modul Penyusunan Kisi-Kisi Instrumen Soal")
    
    # 2. MENU INPUT KHUSUS KISI-KISI
    input_permintaan_kisi = st.text_area(
        "Masukkan rincian materi / indikator kompetensi yang ingin dimasukkan ke dalam kisi-kisi:",
        placeholder="Contoh: Buat 5 nomor soal. Materi tentang Tantangan Global Pancasila, sebaran level kognitif dominan C4 dan C5.",
        height=100
    )
    
    # Tampilan Output Utama Teks Area Kisi-kisi
    st.session_state.kisi_data["konten"] = st.text_area(
        "Hasil Matriks Kisi-Kisi Soal (Format Tabel):", 
        value=st.session_state.kisi_data["konten"], 
        height=300
    )
    
    if st.button("✨ Generate Kisi-Kisi Soal dengan Gemini"):
        with st.spinner("AI sedang merancang matriks kisi-kisi..."):
            prompt_dasar = f"Buatlah tabel kisi-kisi soal evaluasi yang padat untuk materi '{topik_pokok}' mata pelajaran {mapel} kelas XII. Gunakan format tabel Markdown dengan kolom: No, Kompetensi Dasar/Indikator, Indikator Soal, Level Kognitif (C1-C6), Bentuk Soal, Nomor Soal. Buat sepadat mungkin."
            if input_permintaan_kisi:
                prompt_dasar += f" Sesuaikan isi tabel matriksnya dengan instruksi khusus ini: {input_permintaan_kisi}"
            
            hasil = panggil_ai(prompt_dasar)
            if hasil:
                st.session_state.kisi_data["konten"] = hasil
                st.success("Kisi-kisi berhasil diperbarui!")
                st.rerun()
                
    st.markdown("---")
    st.subheader("📥 Cetak Kisi-Kisi Soal")
    if st.session_state.kisi_data["konten"]:
        kisi_components = [("", st.session_state.kisi_data["konten"])]
        file_doc = buat_dokumen_word("MATRIKS KISI-KISI INSTRUMEN PENILAIAN", kisi_components)
        st.download_button(
            label="📥 Unduh File Kisi-Kisi (.docx)", 
            data=file_doc, 
            file_name=f"Kisi_Kisi_{topik_pokok.replace(' ', '_')}.docx"
        )
    else:
        st.info("Silakan masukkan input dan klik generate terlebih dahulu untuk mengunduh.")


# --- MENU 3: MEMBUAT SOAL ---
elif menu_utama == "❓ Membuat Soal":
    st.header("Modul Pembuatan Lembar Soal & Kunci Jawaban")
    
    # 3. MENU INPUT KHUSUS SOAL EVALUASI
    input_permintaan_soal = st.text_area(
        "Masukkan jenis soal, jumlah, atau tipe materi ujian yang Anda inginkan:",
        placeholder="Contoh: Buat 10 soal pilihan ganda tingkat kesulitan sedang dan 3 soal esai HOTS analisis studi kasus di lingkungan madrasah.",
        height=100
    )
    
    # Tampilan Output Utama Teks Area Soal
    st.session_state.soal_data["konten"] = st.text_area(
        "Hasil Naskah Lembar Soal & Kunci:", 
        value=st.session_state.soal_data["konten"], 
        height=300
    )
    
    if st.button("✨ Generate Lembar Soal dengan Gemini"):
        with st.spinner("AI sedang memformulasikan daftar soal evaluasi..."):
            prompt_dasar = f"Buatlah paket lembar soal evaluasi berdasarkan Kurikulum Berbasis Cinta (KBC) dan materi '{topik_pokok}' untuk {mapel} kelas XII. Pilihan ganda A,B,C,D,E disusun horizontal menyamping agar hemat kertas. Sertakan Kunci Jawaban singkat di bawah."
            if input_permintaan_soal:
                prompt_dasar += f" Modifikasi paket soalnya dengan aturan khusus dari guru ini: {input_permintaan_soal}"
            else:
                prompt_dasar += " Terdiri dari default: 5 soal pilihan ganda dan 2 soal esai HOTS."
                
            hasil = panggil_ai(prompt_dasar)
            if hasil:
                st.session_state.soal_data["konten"] = hasil
                st.success("Bank soal evaluasi berhasil disimpan!")
                st.rerun()
                
    st.markdown("---")
    st.subheader("📥 Cetak Lembar Soal")
    if st.session_state.soal_data["konten"]:
        soal_components = [("", st.session_state.soal_data["konten"])]
        file_doc = buat_dokumen_word("LEMBAR EVALUASI SISWA & KUNCI JAWABAN", soal_components)
        st.download_button(
            label="📥 Unduh File Lembar Soal (.docx)", 
            data=file_doc, 
            file_name=f"Soal_Evaluasi_{topik_pokok.replace(' ', '_')}.docx"
        )
    else:
        st.info("Silakan masukkan input dan klik generate terlebih dahulu untuk mengunduh.")
