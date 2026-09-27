import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def create_docx():
    doc = docx.Document()
    
    # Page setup (A4, 0.4 inch margins)
    section = doc.sections[0]
    section.page_width = Inches(8.27)
    section.page_height = Inches(11.69)
    section.top_margin = Inches(0.3)
    section.bottom_margin = Inches(0.3)
    section.left_margin = Inches(0.4)
    section.right_margin = Inches(0.4)

    # Header Title Table
    table_header = doc.add_table(rows=1, cols=2)
    table_header.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_header.autofit = False
    table_header.columns[0].width = Inches(4.7)
    table_header.columns[1].width = Inches(2.7)

    cell_left = table_header.cell(0, 0)
    cell_right = table_header.cell(0, 1)

    set_cell_background(cell_left, "EEF2FF")
    set_cell_background(cell_right, "EEF2FF")
    set_cell_margins(cell_left, 100, 100, 150, 150)
    set_cell_margins(cell_right, 100, 100, 150, 150)

    p1 = cell_left.paragraphs[0]
    p1.paragraph_format.space_before = Pt(0)
    p1.paragraph_format.space_after = Pt(2)
    r1 = p1.add_run("🌈 LEMBAR AKTIVITAS PELANGI")
    r1.font.name = "Arial"
    r1.font.size = Pt(14)
    r1.font.bold = True
    r1.font.color.rgb = RGBColor(67, 56, 202)

    p2 = cell_left.add_paragraph()
    p2.paragraph_format.space_before = Pt(0)
    p2.paragraph_format.space_after = Pt(0)
    r2 = p2.add_run("Menebali Huruf SAMAR & Mengenal Hewan (3–5 Tahun)")
    r2.font.name = "Arial"
    r2.font.size = Pt(9.5)
    r2.font.bold = True
    r2.font.color.rgb = RGBColor(71, 85, 105)

    p_meta = cell_right.paragraphs[0]
    p_meta.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_meta.paragraph_format.space_before = Pt(0)
    p_meta.paragraph_format.space_after = Pt(2)
    
    r3 = p_meta.add_run("Nama: ")
    r3.font.size = Pt(9)
    r3.font.bold = True
    r3_val = p_meta.add_run("Pelangi Lotusia Rabbani\n")
    r3_val.font.size = Pt(9)
    r3_val.font.bold = True
    r3_val.font.color.rgb = RGBColor(67, 56, 202)

    r4 = p_meta.add_run("Hari / Tgl: ")
    r4.font.size = Pt(9)
    r4.font.bold = True
    r4_val = p_meta.add_run("............................\n")
    r4_val.font.size = Pt(9)

    r5 = p_meta.add_run("Waktu: ")
    r5.font.size = Pt(9)
    r5.font.bold = True
    r5_val = p_meta.add_run("⏱️ 3 – 5 Menit")
    r5_val.font.size = Pt(9)
    r5_val.font.bold = True
    r5_val.font.color.rgb = RGBColor(217, 119, 6)

    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    # Instruction Table
    t_inst = doc.add_table(rows=1, cols=1)
    t_inst.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell_inst = t_inst.cell(0, 0)
    cell_inst.width = Inches(7.4)
    set_cell_background(cell_inst, "FEF3C7")
    set_cell_margins(cell_inst, 80, 80, 150, 150)
    
    p_inst = cell_inst.paragraphs[0]
    p_inst.paragraph_format.space_before = Pt(0)
    p_inst.paragraph_format.space_after = Pt(0)
    r_inst_t = p_inst.add_run("✏️ PETUNJUK (HURUF SAMAR UNTUK DITEBALI): ")
    r_inst_t.font.bold = True
    r_inst_t.font.size = Pt(9.5)
    r_inst_t.font.color.rgb = RGBColor(120, 53, 15)
    
    r_inst_c = p_inst.add_run("Ayo Pelangi, tebali huruf-huruf SAMAR ABU-ABU di bawah ini menggunakan krayon/pensil warna favoritmu!")
    r_inst_c.font.size = Pt(9.5)
    r_inst_c.font.color.rgb = RGBColor(120, 53, 15)

    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    # 4 Items Data - FAINT LIGHT GRAY (RGB 209, 213, 219) for true tracing!
    items = [
        {
            "num": "1",
            "title": "JERAPAH / GIRAFFE 🦒",
            "bg": "FFFBEB",
            "id": "J - E - R - A - P - A - H",
            "en": "G - I - R - A - F - F - E",
            "id_size": 24,
            "en_size": 24
        },
        {
            "num": "2",
            "title": "LUMBA-LUMBA / DOLPHIN 🐬",
            "bg": "ECFEFF",
            "id": "L - U - M - B - A - L - U - M - B - A",
            "en": "D - O - L - P - H - I - N",
            "id_size": 18, # Scaled to fit 100% cleanly
            "en_size": 24
        },
        {
            "num": "3",
            "title": "PAUS / WHALE 🐳",
            "bg": "EFF6FF",
            "id": "P - A - U - S",
            "en": "W - H - A - L - E",
            "id_size": 26,
            "en_size": 26
        },
        {
            "num": "4",
            "title": "MOMMY ❤️",
            "bg": "FDF2F8",
            "id": "",
            "en": "M - O - M - M - Y",
            "id_size": 0,
            "en_size": 32
        }
    ]

    for item in items:
        t_item = doc.add_table(rows=1, cols=1)
        t_item.alignment = WD_TABLE_ALIGNMENT.CENTER
        c_item = t_item.cell(0, 0)
        c_item.width = Inches(7.4)
        set_cell_background(c_item, item["bg"])
        set_cell_margins(c_item, 80, 80, 120, 120)

        p_item = c_item.paragraphs[0]
        p_item.paragraph_format.space_before = Pt(0)
        p_item.paragraph_format.space_after = Pt(2)
        
        r_num = p_item.add_run(f"{item['num']}. {item['title']}\n")
        r_num.font.bold = True
        r_num.font.size = Pt(11)
        r_num.font.color.rgb = RGBColor(30, 41, 59)

        if item["id"]:
            r_id_tag = p_item.add_run("   🇮🇩 ")
            r_id_tag.font.size = Pt(9.5)
            r_id = p_item.add_run(f"{item['id']}\n")
            r_id.font.size = Pt(item["id_size"])
            r_id.font.bold = True
            # VERY LIGHT FAINT GRAY FOR REAL TRACING (RGB 209, 213, 219)
            r_id.font.color.rgb = RGBColor(209, 213, 219)

        if item["en"]:
            r_en_tag = p_item.add_run("   🇬🇧 " if item["id"] else "   ❤️ ")
            r_en_tag.font.size = Pt(9.5)
            r_en = p_item.add_run(f"{item['en']}")
            r_en.font.size = Pt(item["en_size"])
            r_en.font.bold = True
            # FAINT SOFT PINK / LIGHT GRAY
            r_en.font.color.rgb = RGBColor(244, 114, 182) if not item["id"] else RGBColor(209, 213, 219)

        doc.add_paragraph().paragraph_format.space_after = Pt(3)

    # Footer Notes Table
    t_foot = doc.add_table(rows=1, cols=2)
    t_foot.alignment = WD_TABLE_ALIGNMENT.CENTER
    c_f1 = t_foot.cell(0, 0)
    c_f2 = t_foot.cell(0, 1)
    c_f1.width = Inches(4.7)
    c_f2.width = Inches(2.7)

    set_cell_background(c_f1, "F8FAFC")
    set_cell_background(c_f2, "FFFFFF")
    set_cell_margins(c_f1, 80, 80, 100, 100)
    set_cell_margins(c_f2, 80, 80, 100, 100)

    p_f1 = c_f1.paragraphs[0]
    p_f1.paragraph_format.space_before = Pt(0)
    p_f1.paragraph_format.space_after = Pt(2)
    r_f1_t = p_f1.add_run("📌 CATATAN UNTUK AYAH / IBU:\n")
    r_f1_t.font.bold = True
    r_f1_t.font.size = Pt(9)

    r_f1_c1 = p_f1.add_run("• Puji Pelafalannya: Ajak Pelangi mengucapkan kata dalam bahasa Inggrisnya.\n")
    r_f1_c1.font.size = Pt(8)
    r_f1_c2 = p_f1.add_run("• Tebali huruf samar abu-abu di atas dengan krayon/pensil warna.")
    r_f1_c2.font.size = Pt(8)

    p_f2 = c_f2.paragraphs[0]
    p_f2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_f2.paragraph_format.space_before = Pt(0)
    p_f2.paragraph_format.space_after = Pt(2)
    r_f2_t = p_f2.add_run("⭐ Nilai Pelangi ⭐\n")
    r_f2_t.font.bold = True
    r_f2_t.font.size = Pt(9)
    r_stars = p_f2.add_run("★ ★ ★ ★ ★\n\n")
    r_stars.font.size = Pt(13)
    r_stars.font.color.rgb = RGBColor(203, 213, 225)
    r_paraf = p_f2.add_run("Paraf Ortu: ____________")
    r_paraf.font.size = Pt(8)

    doc.save("lembar_menebali_huruf_hewan_pelangi.docx")
    print("DOCX with FAINT LIGHT GRAY tracing letters created successfully!")

if __name__ == "__main__":
    create_docx()
