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

# --- 2. SETUP API KEY & IDENTITAS DI SIDEBAR ---
st.sidebar.title("🔑 Pengaturan & Identitas")
api_key = st.sidebar.text_input("Masukkan Google Gemini API Key:", type="password")

st.sidebar.markdown("---")
st.sidebar.subheader("📄 Identitas Dokumen")
nama_guru = st.sidebar.text_input("Nama Guru:", value="Kaharuddin, S.Pd")
mapel = st.sidebar.text_input("Mata Pelajaran:", value="PPKn")
topik_pokok = st.sidebar.text_input("Topik Pokok:", value="Pancasila dalam Kehidupan Global")

if api_key:
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel('gemini-2.5-flash')
else:
    st.sidebar.warning("Silakan masukkan API Key di sini untuk mengaktifkan AI.")

# --- 3. JUDUL UTAMA APLIKASI ---
st.title("🍎 Sistem Perencanaan & Evaluasi Pembelajaran (KBC)")
st.caption("Edisi Eco-Friendly (Hemat Kertas): MAN 2 Kota Makassar")

# --- 4. INISIALISASI SESSION STATE ---
if 'rpp_data' not in st.session_state:
    st.session_state.rpp_data = {
        # Modul 1: RPP
        "bagian_awal": "", "kegiatan_inti": "", "asesmen": "", "lkpd": "", "bahan_bacaan": "", "rubrik": "",
        # Modul 2: Kisi-Kisi
        "kisi_kisi": "",
        # Modul 3: Soal
        "soal_evaluasi": ""
    }

# Simpan metadata ke session state
st.session_state.rpp_data['nama_guru'] = nama_guru
st.session_state.rpp_data['mapel'] = mapel
st.session_state.rpp_data['topik'] = topik_pokok


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

# --- 6. FUNGSI EKSPOR COCOK UNTUK MASING-MASING MODUL ---
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
    
    # Margin super hemat kertas
    for section in doc.sections:
        section.page_width = Inches(8.27)
        section.page_height = Inches(11.69)
        section.top_margin = Inches(0.6)
        section.bottom_margin = Inches(0.6)
        section.left_margin = Inches(0.6)
        section.right_margin = Inches(0.6)

    header_p = doc.add_paragraph()
    header_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h_run = doc.add_paragraph().add_run(f"{judul_dokumen}\n")
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
        ("Nama Guru", f": {st.session_state.rpp_data.get('nama_guru')}", "Mata Pelajaran", f": {st.session_state.rpp_data.get('mapel')}"),
        ("Fase/Kelas", f": F / XII", "Topik Pokok", f": {st.session_state.rpp_data.get('topik')}")
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
            if judul: # Jika ada subjudul bab
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


# --- 7. SISTEM NAVIGASI UTAMA (3 MENU TERPISAH) ---
tab_rpp, tab_kisi, tab_soal = st.tabs([
    "📝 1. Membuat RPP", 
    "📊 2. Menyusun Kisi-Kisi Soal", 
    "❓ 3. Membuat Soal Evaluasi"
])


# ==========================================
# --- MENU 1: MEMBUAT RPP ---
# ==========================================
with tab_rpp:
    st.header("Modul 1: Pembuatan RPP Padat & Mendalam")
    
    pilihan_bagian = st.selectbox(
        "Pilih Komponen RPP:",
        ["Bagian Awal & Kompetensi", "Skenario Kegiatan Inti", "Asesmen", "LKPD", "Bahan Bacaan", "Rubrik Penilaian"],
        key="sb_rpp"
    )
    
    map_key_rpp = {
        "Bagian Awal & Kompetensi": "bagian_awal", "Skenario Kegiatan Inti": "kegiatan_inti",
        "Asesmen": "asesmen", "LKPD": "lkpd", "Bahan Bacaan": "bahan_bacaan", "Rubrik Penilaian": "rubrik"
    }
    key_rpp = map_key_rpp[pilihan_bagian]
    
    st.session_state.rpp_data[key_rpp] = st.text_area(
        f"Konten untuk {pilihan_bagian}:", value=st.session_state.rpp_data[key_rpp], height=300, key=f"ta_{key_rpp}"
    )
    
    if st.button(f"✨ Generate {pilihan_bagian}", key=f"btn_{key_rpp}"):
        with st.spinner("AI sedang menyusun RPP..."):
            prompt = f"Buatlah isi untuk modul {pilihan_bagian} tentang materi '{topik_pokok}' mata pelajaran {mapel} untuk tingkat kelas XII. Sesuai dengan prinsip Kurikulum Berbasis Cinta (KBC) dan hemat kertas cetak."
            hasil = panggil_ai(prompt)
            if hasil:
                st.session_state.rpp_data[key_rpp] = hasil
                st.rerun()
                
    st.markdown("---")
    if st.button("📥 Unduh Dokumen RPP (.docx)", key="dl_rpp"):
        rpp_components = [
            ("I. BAGIAN AWAL & ANALISIS KOMPETENSI", st.session_state.rpp_data["bagian_awal"]),
            ("II. SKENARIO KEGIATAN INTI", st.session_state.rpp_data["kegiatan_inti"]),
            ("III. ASESMEN PEMBELAJARAN MENDALAM", st.session_state.rpp_data["asesmen"]),
            ("IV. LEMBAR KERJA PESERTA DIDIK (LKPD)", st.session_state.rpp_data["lkpd"]),
            ("V. BAHAN BACAAN GURU & SISWA", st.session_state.rpp_data["bahan_bacaan"]),
            ("VI. RUBRIK PENILAIAN OTENTIK SISWA", st.session_state.rpp_data["rubrik"])
        ]
        file_doc = buat_dokumen_word("PERENCANAAN PEMBELAJARAN MENDALAM (RPP)", rpp_components)
        st.download_button(label="Klik di sini untuk mendownload", data=file_doc, file_name=f"RPP_{topik_pokok.replace(' ', '_')}.docx")


# ==========================================
# --- MENU 2: MENYUSUN KISI-KISI SOAL ---
# ==========================================
with tab_kisi:
    st.header("Modul 2: Penyusunan Kisi-Kisi Instrumen Soal")
    st.markdown("Gunakan AI untuk merancang matriks keterkaitan antara indikator, level kognitif, dan nomor soal.")
    
    st.session_state.rpp_data["kisi_kisi"] = st.text_area(
        "Matriks Kisi-Kisi Soal (Format Tabel):", value=st.session_state.rpp_data["kisi_kisi"], height=350, key="ta_kisi"
    )
    
    if st.button("✨ Generate Kisi-Kisi Soal dengan Gemini", key="btn_kisi"):
        with st.spinner("AI sedang merancang tabel kisi-kisi..."):
            prompt = f"Buatlah tabel kisi-kisi soal evaluasi yang padat untuk materi '{topik_pokok}' mata pelajaran {mapel} kelas XII. Gunakan format tabel Markdown dengan kolom: No, Kompetensi Dasar/Indikator, Indikator Soal, Level Kognitif (C1-C6), Bentuk Soal, Nomor Soal. Buat sepadat mungkin demi hemat kertas."
            hasil = panggil_ai(prompt)
            if hasil:
                st.session_state.rpp_data["kisi_kisi"] = hasil
                st.rerun()
                
    st.markdown("---")
    if st.button("📥 Unduh Kisi-Kisi Soal (.docx)", key="dl_kisi"):
        kisi_components = [("MATRIKS KISI-KISI INSTRUMEN PENILAIAN", st.session_state.rpp_data["kisi_kisi"])]
        file_doc = buat_dokumen_word("KISI-KISI SOAL EVALUASI", kisi_components)
        st.download_button(label="Klik di sini untuk mendownload", data=file_doc, file_name=f"Kisi_Kisi_{topik_pokok.replace(' ', '_')}.docx")


# ==========================================
# --- MENU 3: MEMBUAT SOAL ---
# ==========================================
with tab_soal:
    st.header("Modul 3: Pembuatan Lembar Soal & Kunci Jawaban")
    st.markdown("Membuat bank soal pilihan ganda hemat kertas dan esai HOTS yang mendalam.")
    
    st.session_state.rpp_data["soal_evaluasi"] = st.text_area(
        "Naskah Soal & Kunci Jawaban:", value=st.session_state.rpp_data["soal_evaluasi"], height=350, key="ta_soal"
    )
    
    if st.button("✨ Generate Paket Soal dengan Gemini", key="btn_soal"):
        with st.spinner("AI sedang menyusun daftar soal..."):
            prompt = f"Buatlah paket lembar soal evaluasi berdasarkan Kurikulum Berbasis Cinta (KBC) dan materi '{topik_pokok}' untuk {mapel} kelas XII. Terdiri dari 5 soal pilihan ganda (beserta pilihan A, B, C, D, E disusun horizontal agar hemat kertas) dan 2 soal esai HOTS mendalam yang mengintegrasikan nilai ekoteologi. Sertakan Kunci Jawaban singkat di bagian paling bawah."
            hasil = panggil_ai(prompt)
            if hasil:
                st.session_state.rpp_data["soal_evaluasi"] = hasil
                st.rerun()
                
    st.markdown("---")
    if st.button("📥 Unduh Lembar Soal (.docx)", key="dl_soal"):
        soal_components = [("LEMBAR EVALUASI SISWA & KUNCI JAWABAN", st.session_state.rpp_data["soal_evaluasi"])]
        file_doc = buat_dokumen_word("LEMBAR SOAL EVALUASI PEMBELAJARAN", soal_components)
        st.download_button(label="Klik di sini untuk mendownload", data=file_doc, file_name=f"Soal_Evaluasi_{topik_pokok.replace(' ', '_')}.docx")
