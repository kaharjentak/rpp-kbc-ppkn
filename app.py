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

# Konfigurasi Halaman Web Streamlit
st.set_page_config(page_title="Generator RPP Hemat Kertas MAN 2 Kota Makassar", layout="wide")

# Setup API Key di Sidebar
st.sidebar.title("🔑 Pengaturan")
api_key = st.sidebar.text_input("Masukkan Google Gemini API Key:", type="password")

if api_key:
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel('gemini-2.5-flash')
else:
    st.sidebar.warning("Silakan masukkan API Key di sini untuk mengaktifkan AI.")

st.title("🍎 Sistem Perencanaan Pembelajaran Mendalam & KBC")
st.caption("Edisi Eco-Friendly (Hemat Kertas): Pak Kaharuddin, S.Pd - MAN 2 Kota Makassar")

if 'rpp_data' not in st.session_state:
    st.session_state.rpp_data = {
        "bagian_awal": "", "kegiatan_inti": "", "asesmen": "", "lkpd": "", "bahan_bacaan": "", "rubrik": ""
    }

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

# --- FUNGSI FORMAT TINGKAT TINGGI ---
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

def cetak_teks_berformat(paragraph, teks_mentah):
    parts = re.split(r'(\*\*.*?\*\*)', teks_mentah)
    for part in parts:
        if part.startswith('**') and part.endswith('**'):
            teks_bersih = part[2:-2]
            run = paragraph.add_run(teks_bersih)
            run.font.bold = True
        else:
            run = paragraph.add_run(part)
        run.font.name = 'Arial'
        run.font.size = Pt(10)
        run.font.color.rgb = RGBColor(45, 55, 72)

def buat_file_word_eco():
    doc = Document()
    
    # MARGIN SUPER SEMPIT (1.5 cm / 0.6 Inci)
    for section in doc.sections:
        section.page_width = Inches(8.27)
        section.page_height = Inches(11.69)
        section.top_margin = Inches(0.6)
        section.bottom_margin = Inches(0.6)
        section.left_margin = Inches(0.6)
        section.right_margin = Inches(0.6)

    # HEADER IDENTITAS ATAS
    header_p = doc.add_paragraph()
    header_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h_run = header_p.add_run("PERENCANAAN PEMBELAJARAN MENDALAM (RPP / MODUL AJAR)\n")
    h_run.font.name = 'Georgia'
    h_run.font.size = Pt(13)
    h_run.font.bold = True
    h_run.font.color.rgb = RGBColor(26, 54, 93)
    
    sub_run = header_p.add_run("Integrasi Ekoteologi dalam Kurikulum Berbasis Cinta (KBC) — MAN 2 Kota Makassar\n")
    sub_run.font.name = 'Arial'
    sub_run.font.size = Pt(9.5)
    sub_run.font.italic = True
    sub_run.font.color.rgb = RGBColor(100, 110, 125)

    # Tabel Ringkas Identitas Atas
    meta_table = doc.add_table(rows=2, cols=4)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    pasang_garis_pembatas_tabel(meta_table, "A0AEC0")
    
    identitas_data = [
        ("Nama Guru", f": {st.session_state.rpp_data.get('nama_guru', 'Kaharuddin, S.Pd')}", "Mata Pelajaran", f": {st.session_state.rpp_data.get('mapel', 'PPKn')}"),
        ("Fase/Kelas", f": F / XII", "Topik Pokok", f": {st.session_state.rpp_data.get('topik', 'Pancasila dalam Kehidupan Global')}")
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

    # GABUNGKAN KONTEN SECARA PADAT
    komponen = [
        ("I. BAGIAN AWAL & ANALISIS KOMPETENSI", st.session_state.rpp_data["bagian_awal"]),
        ("II. SKENARIO KEGIATAN INTI (PERTEMUAN 1 & 2)", st.session_state.rpp_data["kegiatan_inti"]),
        ("III. ASESMEN PEMBELAJARAN MENDALAM", st.session_state.rpp_data["asesmen"]),
        ("IV. LEMBAR KERJA PESERTA DIDIK (LKPD)", st.session_state.rpp_data["lkpd"]),
        ("V. BAHAN BACAAN GURU & SISWA", st.session_state.rpp_data["bahan_bacaan"]),
        ("VI. RUBRIK PENILAIAN OTENTIK SISWA", st.session_state.rpp_data["rubrik"])
    ]

    for judul, isi in komponen:
        if isi:
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
                            # PERBAIKAN: Mengatur perataan teks header tabel yang benar
                            for p_cell in cell.paragraphs:
                                p_cell.alignment = WD_ALIGN_PARAGRAPH.CENTER
                                if p_cell.runs:
                                    p_cell.runs[0].font.bold = True
                                    p_cell.runs[0].font.color.rgb = RGBColor(255, 255, 255)
                                    p_cell.runs[0].font.name = 'Arial'
