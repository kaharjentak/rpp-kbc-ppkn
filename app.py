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
                                    p_cell.runs[0].font.size = Pt(9)
                    else:
                        row = tabel_obj.add_row()
                        for idx, teks in enumerate(kolom_data):
                            if idx < len(row.cells):
                                cell = row.cells[idx]
                                cell.text = "" 
                                p_cell = cell.paragraphs[0]
                                format_cell_margins(cell, top=50, bottom=50, left=90, right=90)
                                set_cell_background(cell, "F9FAFB" if len(tabel_obj.rows) % 2 == 0 else "FFFFFF")
                                cetak_teks_berformat(p_cell, teks)
                    continue
                else:
                    di_dalam_tabel = False

                p = doc.add_paragraph()
                if baris_bersih.startswith(('###', 'A.', 'B.', 'C.', 'D.', 'E.', 'F.', 'G.', 'H.', 'I.', 'J.', 'K.')):
                    teks_judul_sub = baris_bersih.replace('###', '').strip()
                    sub_run = p.add_run(teks_judul_sub)
                    sub_run.font.name = 'Georgia'
                    sub_run.font.size = Pt(10.5)
                    sub_run.font.bold = True
                    sub_run.font.color.rgb = RGBColor(44, 82, 130)
                    p.paragraph_format.space_before = Pt(4)
                    p.paragraph_format.space_after = Pt(2)
                else:
                    if baris_bersih.startswith(('*', '-', '1.', '2.', '3.', '4.', '5.')):
                        p.paragraph_format.left_indent = Inches(0.2)
                    p.paragraph_format.line_spacing = 1.05
                    p.paragraph_format.space_after = Pt(2)
                    cetak_teks_berformat(p, baris_bersih)

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- PANEL DATA INPUT UTAMA ---
st.subheader("📝 Data Utama Modul Ajar")
col1, col2, col3 = st.columns(3)
with col1:
    nama_madrasah = st.text_input("Nama Madrasah:", value="MAN 2 KOTA MAKASSAR")
    nama_guru = st.text_input("Nama Guru:", value="Kaharuddin, S.Pd")
    mapel = st.text_input("Mata Pelajaran:", value="PPKn")
with col2:
    fase_kelas = st.text_input("Fase / Kelas / Smt:", value="F / XII / ganjil")
    topik = st.text_input("Topik Pembelajaran:", value="Pancasila dalam Kehidupan Global")
    sub_topik = st.text_input("Sub Topik Pembelajaran:", value="Hambatan dan tantangan Pancasila dalam kehidupan Global")
with col3:
    waktu = st.text_input("Alokasi Waktu:", value="4 JP ( 2 x Pertemuan )")

st.session_state.rpp_data["nama_madrasah"] = nama_madrasah
st.session_state.rpp_data["nama_guru"] = nama_guru
st.session_state.rpp_data["mapel"] = mapel
st.session_state.rpp_data["topik"] = topik

st.markdown("---")
st.subheader("🚀 Alur Kerja Penyusunan Modul")

tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "1. Bagian Awal", "2. Kegiatan Inti", "3. Asesmen", "4. LKPD", "5. Bahan Bacaan", "6. Rubrik Penilaian"
])

konteks_identitas = f"\nNama Madrasah: {nama_madrasah}\nNama Guru: {nama_guru}\nMapel: {mapel}\nFase/Kelas/Smt: {fase_kelas}\nTopik Pembelajaran: {topik}\nSub Topik Pembelajaran: {sub_topik}\nAlokasi Waktu: {waktu}\nGuru adalah pengajar PPKN jenjang SMA/MA.\n"

with tab1:
    st.info("Tahap 1: Menyusun Informasi Identitas, Target Kompetensi, Profil Lulusan, dan Eksplorasi Ekoteologi KBC.")
    if st.button("Generate Tahap 1: Bagian Awal RPP"):
        prompt1 = konteks_identitas + """
        Tujuan pembelajaran umum adalah Menganalisis Hambatan dan tantangan Pancasila dalam kehidupan Global. Buat bagian awal rencana pembelajaran dengan spesifikasi berikut:
        A. Identifikasi Peserta Didik: Tuliskan identifikasi kesiapan peserta didik sebelum belajar, seperti pengetahuan awal, minat, latar belakang, dan kebutuhan belajar. Tulis dalam bentuk poin.
        B. Identifikasi Materi Pembelajaran: Tuliskan analisis materi pelajaran seperti jenis pengetahuan yang akan dicapai, relevansi dengan kehidupan nyata peserta didik, tingkat kesulitan, struktur materi, serta integrasi nilai and karakter, and lainnya. Tulis dalam bentuk poin.
        C. Dimensi Profil Lulusan: Tuliskan dimensi profil lulusan yang akan dicapai (Pilih 4 dari 8 dimensi: Keimanan and Ketakwaan, Kewargaan, Penalaran Kritis, Kreativitas, Kolaborasi, Kemandirian, Kesehatan, and Komunikasi).
        D. Topik Panca Cinta: Tuliskan topik panca cinta yang sesuai dengan materi pembelajaran (Pilih 3 dari 5 topik: Cinta Allah and Rasul-Nya, Cinta Ilmu, Cinta Lingkungan, Cinta Diri and Sesama Manusia, and Cinta Tanah Air).
        E. Materi Integrasi KBC: Tuliskan materi integrasi ekoteologi dalam Kurikulum Berbasis Cinta (KBC) (Panca Cinta) yang akan dikembangkan and relevan dengan materi pembelajaran. Tuliskan dalam bentuk paragraf yang mendalam.
        F. Topik Pembelajaran: Tuliskan topik pembelajaran yang relevan dengan capaian and tujuan pembelajaran.
        G. Tujuan Pembelajaran: Tuliskan tujuan pembelajaran umum, kemudian tuliskan tujuan pembelajaran yang rinci per pertemuan yang mencakup aspek utama, yaitu subjek belajar, pengetahuan keterampilan atau sikap yang harus dikuasi dengan kata kerja operasional yang terukur (Sediakan rincian jika ada 2 pertemuan).
        H. Praktik Pedagogis: Tuliskan Model/Strategi/Metode yang ditentukan oleh guru. Pertemuan 1 menggunakan blended learning, model pembelajaran STAD and berbasis masalah.
        I. Kemitraan Pembelajaran: Tuliskan Mitra kerjasama untuk berkolaborasi (lingkungan sekolah, luar sekolah, masyarakat).
        J. Lingkungan Pembelajaran: Tuliskan Lingkungan pembelajaran yang mengintegrasikan antara ruang fisik, ruang virtual, and budaya belajar.
        K. Pemanfaatan Digital: Tuliskan Pemanfaatan teknologi digital menciptakan pembelajaran yang lebih interaktif, kolaboratif, and kontekstual.
        Gunakan format Kode Tabel Markdown murni dengan penanda karakter '|' di setiap batas kolom agar terbaca sistem otomatis.
        """
        with st.spinner("Sedang memproses Tahap 1..."):
            st.session_state.rpp_data["bagian_awal"] = panggil_ai(prompt1)
    if st.session_state.rpp_data["bagian_awal"]:
        st.markdown(st.session_state.rpp_data["bagian_awal"])

with tab2:
    st.info("Tahap 2: Menyusun skenario 2x Pertemuan berbasis Pembelajaran Mendalam (Mindful, Meaningful, Joyful).")
    if st.button("Generate Tahap 2: Kegiatan Inti RPP"):
        if not st.session_state.rpp_data["bagian_awal"]:
            st.error("Silakan lakukan Tahap 1 terlebih dahulu!")
        else:
            prompt2 = f"""Berdasarkan hasil analisis Bagian Awal RPP berikut:\n{st.session_state.rpp_data['bagian_awal']}\n
            Buat langkah pembelajaran yang memuat prinsip pembelajaran mendalam (berkesadaran/mindful, bermakna/meaningful, menggembirakan/joyful) buatkan RPP dengan 2 x pertemuan, dengan spesifikasi berikut:
            Langkah-Langkah Pembelajaran:
            AWAL: Pembuka pembelajaran bertujuan mempersiapkan peserta didik meliputi salam, mengecek kehadiran, mengkondisikan peserta didik siap belajar, orientasi bermakna, apersepsi kontekstual, penyampaian tujuan pembelajaran, and motivasi menggembirakan.
            INTI:
            1. Memahami: Kognitif mendalam kesadaran murid.
            2. Mengaplikasi: Pemecahan masalah atau pengambilan keputusan secara individu/kolaboratif.
            3. Merefleksi: Mengevaluasi esensi proses nyata.
            PENUTUP: Umpan balik reflektif konseptual.
            Gunakan format Kode Tabel Markdown murni dengan penanda karakter '|' di setiap batas kolom jika menyajikan data berkolom.
            """
            with st.spinner("Sedang memproses Tahap 2..."):
                st.session_state.rpp_data["kegiatan_inti"] = panggil_ai(prompt2)
    if st.session_state.rpp_data["kegiatan_inti"]:
        st.markdown(st.session_state.rpp_data["kegiatan_inti"])

with tab3:
    st.info("Tahap 3: Menyusun skema Asesmen Komprehensif (As, For, dan Of Learning).")
    if st.button("Generate Tahap 3: Asesmen"):
        if not st.session_state.rpp_data["kegiatan_inti"]:
            st.error("Silakan lakukan Tahap 2 terlebih dahulu!")
        else:
            prompt3 = f"""Berdasarkan tujuan and semua langkah pembelajaran berikut:\n{st.session_state.rpp_data['kegiatan_inti']}\n
            Buat secara rinci and lengkap asesmen dalam pembelajaran mendalam disesuaikan dengan assessment as learning, assessment for learning, and assessment of learning. Tentukan metode secara komprehensif untuk mengukur pencapaian kompetensi peserta didik dengan spesifikasi:
            1. Asesmen awal Pembelajaran
            2. Asesmen proses Pembelajaran
            3. Asesmen akhir Pembelajaran
            Gunakan format Kode Tabel Markdown murni dengan penanda karakter '|' di setiap batas kolom.
            """
            with st.spinner("Sedang memproses Tahap 3..."):
                st.session_state.rpp_data["asesmen"] = panggil_ai(prompt3)
    if st.session_state.rpp_data["asesmen"]:
        st.markdown(st.session_state.rpp_data["asesmen"])

with tab4:
    st.info("Tahap 4: Membuat Lembar Kerja Peserta Didik (LKPD) Kontekstual.")
    if st.button("Generate Tahap 4: LKPD"):
        if not st.session_state.rpp_data["kegiatan_inti"]:
            st.error("Silakan lakukan Tahap 2 terlebih dahulu!")
        else:
            prompt4 = f"Berdasarkan tujuan pembelajaran and langkah kegiatan pembelajaran pada pertemuan 1 and 2 berikut:\n{st.session_state.rpp_data['kegiatan_inti']}\nBuat Lembar Kerja Peserta Didik (LKPD) yang sesuai and siap digunakan murid pada pembelajaran tersebut. Gunakan format Kode Tabel Markdown murni dengan penanda karakter '|' untuk membuat kolom lembar isian siswa."
            with st.spinner("Sedang memproses Tahap 4..."):
                st.session_state.rpp_data["lkpd"] = panggil_ai(prompt4)
    if st.session_state.rpp_data["lkpd"]:
        st.markdown(st.session_state.rpp_data["lkpd"])

with tab5:
    st.info("Tahap 5: Membuat Bahan Bacaan Guru & Siswa dilengkapi contoh soal latihan.")
    if st.button("Generate Tahap 5: Bahan Bacaan"):
        if not st.session_state.rpp_data["kegiatan_inti"]:
            st.error("Silakan lakukan Tahap 2 terlebih dahulu!")
        else:
            prompt5 = f"Berdasarkan tujuan pembelajaran and langkah kegiatan pembelajaran pada pertemuan 1 and 2 berikut:\n{st.session_state.rpp_data['kegiatan_inti']}\nBuat materi pembelajaran atau bahan bacaan bagi murid mengenai Hambatan & Tantangan Pancasila di era Global yang dilengkapi dengan contoh soal analisis and soal latihan mandiri."
            with st.spinner("Sedang memproses Tahap 5..."):
                st.session_state.rpp_data["bahan_bacaan"] = panggil_ai(prompt5)
    if st.session_state.rpp_data["bahan_bacaan"]:
        st.markdown(st.session_state.rpp_data["bahan_bacaan"])

with tab6:
    st.info("Tahap 6: Menyusun Rubrik Penilaian Deskriptif berdasarkan Asesmen Tahap 3.")
    if st.button("Generate Tahap 6: Rubrik Penilaian"):
        if not st.session_state.rpp_data["asesmen"]:
            st.error("Silakan lakukan Tahap 3 terlebih dahulu!")
        else:
            prompt6 = f"Berdasarkan instrumen Asesmen berikut:\n{st.session_state.rpp_data['asesmen']}\nBuat paket Rubrik Penilaian lengkap berbentuk TABEL MARKDOWN MURNI (wajib menggunakan format karakter '|') dengan skor kriteria 1 sampai 4."
            with st.spinner("Sedang memproses Tahap 6..."):
                st.session_state.rpp_data["rubrik"] = panggil_ai(prompt6)
    if st.session_state.rpp_data["rubrik"]:
        st.markdown(st.session_state.rpp_data["rubrik"])

# --- BAGIAN EKSPOR TOTAL (DESAIN SUPER HEMAT KERTAS) ---
st.markdown("---")
st.subheader("💾 Unduh Bundel Dokumen Hasil Kerja (Desain Eco-Friendly)")
if st.button("🖨️ Ambil Berkas RPP / Modul Lengkap (.docx)"):
    if not st.session_state.rpp_data["bagian_awal"]:
        st.warning("Anda belum membuat komponen apa pun. Silakan generate setidaknya Tahap 1.")
    else:
        file_final = buat_file_word_eco()
        st.download_button(
            label="📥 Klik di Sini untuk Mengunduh File Hasil Cetak Hemat Kertas",
            data=file_final,
            file_name=f"RPP_Eco_Mendalam_KBC_Pak_Kaharuddin.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
