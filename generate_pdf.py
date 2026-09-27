from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.pdfgen import canvas

def draw_giraffe(canvas, x, y):
    canvas.setFillColor(colors.HexColor("#F59E0B"))
    canvas.rect(x + 14, y + 2, 5, 26, stroke=0, fill=1)
    canvas.rect(x + 28, y + 2, 5, 26, stroke=0, fill=1)
    canvas.setFillColor(colors.HexColor("#FBBF24"))
    canvas.ellipse(x + 8, y + 24, x + 40, y + 48, stroke=0, fill=1)
    canvas.rect(x + 26, y + 38, 10, 36, stroke=0, fill=1)
    canvas.setFillColor(colors.HexColor("#B45309"))
    canvas.circle(x + 28, y + 50, 2.5, stroke=0, fill=1)
    canvas.circle(x + 32, y + 62, 3, stroke=0, fill=1)
    canvas.circle(x + 20, y + 34, 2.5, stroke=0, fill=1)
    canvas.setFillColor(colors.HexColor("#FBBF24"))
    canvas.ellipse(x + 24, y + 68, x + 44, y + 82, stroke=0, fill=1)
    canvas.setStrokeColor(colors.HexColor("#B45309"))
    canvas.setLineWidth(2)
    canvas.line(x + 29, y + 80, x + 27, y + 88)
    canvas.line(x + 39, y + 80, x + 41, y + 88)
    canvas.setFillColor(colors.HexColor("#B45309"))
    canvas.circle(x + 27, y + 88, 2, stroke=0, fill=1)
    canvas.circle(x + 41, y + 88, 2, stroke=0, fill=1)
    canvas.setFillColor(colors.HexColor("#1E293B"))
    canvas.circle(x + 30, y + 76, 1.5, stroke=0, fill=1)
    canvas.circle(x + 38, y + 76, 1.5, stroke=0, fill=1)

def draw_dolphin(canvas, x, y):
    canvas.setStrokeColor(colors.HexColor("#38BDF8"))
    canvas.setLineWidth(2.5)
    canvas.line(x + 2, y + 10, x + 46, y + 10)
    canvas.setFillColor(colors.HexColor("#06B6D4"))
    p = canvas.beginPath()
    p.moveTo(x + 5, y + 30)
    p.curveTo(x + 18, y + 65, x + 38, y + 65, x + 45, y + 42)
    p.curveTo(x + 32, y + 24, x + 18, y + 22, x + 5, y + 30)
    canvas.drawPath(p, fill=1, stroke=0)
    canvas.setFillColor(colors.HexColor("#CFFAFF"))
    p2 = canvas.beginPath()
    p2.moveTo(x + 14, y + 26)
    p2.curveTo(x + 26, y + 42, x + 38, y + 42, x + 42, y + 38)
    p2.curveTo(x + 30, y + 24, x + 20, y + 24, x + 14, y + 26)
    canvas.drawPath(p2, fill=1, stroke=0)
    canvas.setFillColor(colors.HexColor("#1E293B"))
    canvas.circle(x + 38, y + 44, 1.5, stroke=0, fill=1)

def draw_whale(canvas, x, y):
    canvas.setFillColor(colors.HexColor("#3B82F6"))
    canvas.ellipse(x + 5, y + 18, x + 43, y + 55, stroke=0, fill=1)
    canvas.setFillColor(colors.HexColor("#DBEAFE"))
    canvas.ellipse(x + 10, y + 14, x + 38, y + 30, stroke=0, fill=1)
    canvas.setFillColor(colors.HexColor("#2563EB"))
    p = canvas.beginPath()
    p.moveTo(x + 7, y + 36)
    p.lineTo(x + 0, y + 46)
    p.lineTo(x + 0, y + 26)
    canvas.drawPath(p, fill=1, stroke=0)
    canvas.setStrokeColor(colors.HexColor("#60A5FA"))
    canvas.setLineWidth(2)
    canvas.line(x + 24, y + 54, x + 18, y + 66)
    canvas.line(x + 26, y + 54, x + 32, y + 66)
    canvas.setFillColor(colors.HexColor("#1E293B"))
    canvas.circle(x + 34, y + 40, 1.5, stroke=0, fill=1)

def draw_mommy(canvas, x, y):
    canvas.setFillColor(colors.HexColor("#FBCFE8"))
    canvas.circle(x + 18, y + 52, 12, stroke=0, fill=1)
    canvas.setFillColor(colors.HexColor("#831843"))
    canvas.ellipse(x + 6, y + 52, x + 30, y + 70, stroke=0, fill=1)
    canvas.setFillColor(colors.HexColor("#FED7AA"))
    canvas.circle(x + 32, y + 34, 10, stroke=0, fill=1)
    canvas.setFillColor(colors.HexColor("#7C2D12"))
    canvas.ellipse(x + 22, y + 34, x + 42, y + 48, stroke=0, fill=1)
    canvas.setFillColor(colors.HexColor("#EC4899"))
    canvas.circle(x + 23, y + 74, 4, stroke=0, fill=1)

def create_pdf(filename="lembar_menebali_huruf_hewan_pelangi.pdf"):
    c = canvas.Canvas(filename, pagesize=A4)
    width, height = A4 # 595.27 x 841.89

    mx = 18
    pw = width - 2 * mx

    # Header Border & Background
    c.setStrokeColor(colors.HexColor("#6366F1"))
    c.setLineWidth(2)
    c.setFillColor(colors.HexColor("#EEF2FF"))
    c.roundRect(mx, height - 75, pw, 58, 8, fill=1, stroke=1)

    # Title
    c.setFillColor(colors.HexColor("#4338CA"))
    c.setFont("Helvetica-Bold", 15)
    c.drawString(mx + 12, height - 38, "🌈 LEMBAR AKTIVITAS PELANGI")

    c.setFillColor(colors.HexColor("#475569"))
    c.setFont("Helvetica-Bold", 9)
    c.drawString(mx + 12, height - 54, "Menebali Huruf Samar & Mengenal Hewan (Untuk Anak Usia 3-5 Tahun)")

    # Meta Info
    c.setFillColor(colors.HexColor("#1E293B"))
    c.setFont("Helvetica-Bold", 9)
    c.drawRightString(mx + pw - 12, height - 34, "Nama: Pelangi Lotusia Rabbani")
    c.drawRightString(mx + pw - 12, height - 48, "Hari / Tgl: ........................................")
    
    c.setFillColor(colors.HexColor("#D97706"))
    c.drawRightString(mx + pw - 12, height - 62, "Waktu: ⏱️ 3 – 5 Menit")

    # Instruction Banner
    c.setStrokeColor(colors.HexColor("#F59E0B"))
    c.setLineWidth(1.5)
    c.setFillColor(colors.HexColor("#FEF3C7"))
    c.roundRect(mx, height - 112, pw, 32, 6, fill=1, stroke=1)

    c.setFillColor(colors.HexColor("#78350F"))
    c.setFont("Helvetica-Bold", 9)
    c.drawString(mx + 10, height - 98, "✏️ PETUNJUK BELAJAR (HURUF SAMAR / TRACING):")
    c.setFont("Helvetica", 8.5)
    c.drawString(mx + 10, height - 108, "Ayo Pelangi, tebali huruf-huruf SAMAR ABU-ABU di bawah ini dengan krayon/pensil warna favoritmu!")

    # Items Data - Very Light Gray Color (#CBD5E1) for true Tracing!
    items = [
        {
            "num": "1",
            "name_id": "JERAPAH",
            "name_en": "GIRAFFE",
            "id_letters": "J - E - R - A - P - A - H",
            "en_letters": "G - I - R - A - F - F - E",
            "bg": "#FFFBEB",
            "border": "#F59E0B",
            "draw": draw_giraffe,
            "id_size": 24,
            "en_size": 24
        },
        {
            "num": "2",
            "name_id": "LUMBA-LUMBA",
            "name_en": "DOLPHIN",
            "id_letters": "L - U - M - B - A - L - U - M - B - A",
            "en_letters": "D - O - L - P - H - I - N",
            "bg": "#ECFEFF",
            "border": "#06B6D4",
            "draw": draw_dolphin,
            "id_size": 17, # Scaled so all 19 chars fit 100% cleanly without overflow!
            "en_size": 24
        },
        {
            "num": "3",
            "name_id": "PAUS",
            "name_en": "WHALE",
            "id_letters": "P - A - U - S",
            "en_letters": "W - H - A - L - E",
            "bg": "#EFF6FF",
            "border": "#3B82F6",
            "draw": draw_whale,
            "id_size": 26,
            "en_size": 26
        },
        {
            "num": "4",
            "name_id": "MOMMY",
            "name_en": "",
            "id_letters": "",
            "en_letters": "M - O - M - M - Y",
            "bg": "#FDF2F8",
            "border": "#EC4899",
            "draw": draw_mommy,
            "id_size": 0,
            "en_size": 32
        }
    ]

    card_h = 142
    start_y = height - 120 - card_h

    for item in items:
        cy = start_y
        
        c.setStrokeColor(colors.HexColor(item["border"]))
        c.setLineWidth(2)
        c.setFillColor(colors.HexColor(item["bg"]))
        c.roundRect(mx, cy, pw, card_h, 8, fill=1, stroke=1)

        # Number Badge
        c.setFillColor(colors.HexColor("#1E293B"))
        c.circle(mx + 16, cy + card_h - 16, 11, stroke=0, fill=1)
        c.setFillColor(colors.white)
        c.setFont("Helvetica-Bold", 11)
        c.drawCentredString(mx + 16, cy + card_h - 20, item["num"])

        # Animal Illustration Box
        c.setFillColor(colors.white)
        c.setStrokeColor(colors.HexColor("#CBD5E1"))
        c.setLineWidth(1)
        c.roundRect(mx + 33, cy + 10, 80, 122, 6, fill=1, stroke=1)

        item["draw"](c, mx + 48, cy + 28)

        c.setFillColor(colors.HexColor(item["border"]))
        c.setFont("Helvetica-Bold", 8.5)
        c.drawCentredString(mx + 73, cy + 16, item["name_id"])

        rx = mx + 120
        rw = pw - 128

        if item["name_id"] != "MOMMY":
            # Row 1: Indonesian
            c.setFillColor(colors.HexColor("#FEE2E2"))
            c.setStrokeColor(colors.HexColor("#FCA5A5"))
            c.roundRect(rx, cy + 76, 56, 52, 4, fill=1, stroke=1)
            c.setFillColor(colors.HexColor("#991B1B"))
            c.setFont("Helvetica-Bold", 9)
            c.drawCentredString(rx + 28, cy + 102, "🇮🇩 ID")
            c.setFont("Helvetica-Bold", 7.5)
            c.drawCentredString(rx + 28, cy + 86, item["name_id"][:6])

            box_x = rx + 62
            box_w = rw - 62
            c.setFillColor(colors.white)
            c.setStrokeColor(colors.HexColor("#CBD5E1"))
            c.roundRect(box_x, cy + 76, box_w, 52, 4, fill=1, stroke=1)

            # Guidelines
            c.setStrokeColor(colors.HexColor("#E2E8F0"))
            c.setLineWidth(0.6)
            c.line(box_x, cy + 118, box_x + box_w, cy + 118)
            c.line(box_x, cy + 86, box_x + box_w, cy + 86)
            c.setStrokeColor(colors.HexColor("#CBD5E1"))
            c.setDash([3, 3], 0)
            c.line(box_x, cy + 102, box_x + box_w, cy + 102)
            c.setDash([], 0)

            # VERY LIGHT FAINT GRAY FOR REAL TRACING (Hex #CBD5E1 / Light Gray)
            c.setFont("Helvetica-Bold", item["id_size"])
            c.setFillColor(colors.HexColor("#CBD5E1")) # Light gray so child can trace over!
            c.drawString(box_x + 6, cy + 92, item["id_letters"])

            # Row 2: English
            c.setFillColor(colors.HexColor("#DBEAFE"))
            c.setStrokeColor(colors.HexColor("#93C5FD"))
            c.roundRect(rx, cy + 14, 56, 52, 4, fill=1, stroke=1)
            c.setFillColor(colors.HexColor("#1E40AF"))
            c.setFont("Helvetica-Bold", 9)
            c.drawCentredString(rx + 28, cy + 40, "🇬🇧 EN")
            c.setFont("Helvetica-Bold", 7.5)
            c.drawCentredString(rx + 28, cy + 24, item["name_en"][:7])

            box_x = rx + 62
            box_w = rw - 62
            c.setFillColor(colors.white)
            c.setStrokeColor(colors.HexColor("#CBD5E1"))
            c.roundRect(box_x, cy + 14, box_w, 52, 4, fill=1, stroke=1)

            c.setStrokeColor(colors.HexColor("#E2E8F0"))
            c.setLineWidth(0.6)
            c.line(box_x, cy + 56, box_x + box_w, cy + 56)
            c.line(box_x, cy + 24, box_x + box_w, cy + 24)
            c.setStrokeColor(colors.HexColor("#CBD5E1"))
            c.setDash([3, 3], 0)
            c.line(box_x, cy + 40, box_x + box_w, cy + 40)
            c.setDash([], 0)

            # VERY LIGHT FAINT GRAY FOR REAL TRACING (#CBD5E1)
            c.setFont("Helvetica-Bold", item["en_size"])
            c.setFillColor(colors.HexColor("#CBD5E1"))
            c.drawString(box_x + 6, cy + 30, item["en_letters"])

        else:
            # MOMMY SPECIAL ROW
            c.setFillColor(colors.HexColor("#FCE7F3"))
            c.setStrokeColor(colors.HexColor("#FBCFE8"))
            c.roundRect(rx, cy + 14, 56, 114, 4, fill=1, stroke=1)
            c.setFillColor(colors.HexColor("#9D174D"))
            c.setFont("Helvetica-Bold", 12)
            c.drawCentredString(rx + 28, cy + 78, "❤️")
            c.setFont("Helvetica-Bold", 9)
            c.drawCentredString(rx + 28, cy + 54, "MOMMY")

            box_x = rx + 62
            box_w = rw - 62
            c.setFillColor(colors.white)
            c.setStrokeColor(colors.HexColor("#CBD5E1"))
            c.roundRect(box_x, cy + 14, box_w, 114, 4, fill=1, stroke=1)

            c.setStrokeColor(colors.HexColor("#E2E8F0"))
            c.setLineWidth(0.8)
            c.line(box_x, cy + 104, box_x + box_w, cy + 104)
            c.line(box_x, cy + 34, box_x + box_w, cy + 34)
            c.setStrokeColor(colors.HexColor("#CBD5E1"))
            c.setDash([4, 4], 0)
            c.line(box_x, cy + 69, box_x + box_w, cy + 69)
            c.setDash([], 0)

            # LIGHT PINK-GRAY FOR MOMMY TRACING (#F472B6 / Light Pink)
            c.setFont("Helvetica-Bold", 32)
            c.setFillColor(colors.HexColor("#F472B6")) # Faint soft pink for tracing over!
            c.drawString(box_x + 12, cy + 50, item["en_letters"])

        start_y -= (card_h + 8)

    # Footer Section
    fy = 20
    fh = 65

    c.setFillColor(colors.HexColor("#F8FAFC"))
    c.setStrokeColor(colors.HexColor("#94A3B8"))
    c.setLineWidth(1)
    c.setDash([3, 3], 0)
    c.roundRect(mx, fy, pw * 0.62, fh, 6, fill=1, stroke=1)
    c.setDash([], 0)

    c.setFillColor(colors.HexColor("#334155"))
    c.setFont("Helvetica-Bold", 8.5)
    c.drawString(mx + 8, fy + fh - 14, "📌 CATATAN UNTUK AYAH / IBU:")
    c.setFont("Helvetica", 8)
    c.drawString(mx + 8, fy + fh - 28, "• Puji Pelafalannya: Ajak Pelangi mengucapkan kata bahasa Inggrisnya saat menebali.")
    c.drawString(mx + 8, fy + fh - 42, "• Sesi Singkat (Micro-learning): Cukup 3–5 menit agar Pelangi tidak merasa tertekan.")
    c.drawString(mx + 8, fy + fh - 54, "• Tebali huruf samar abu-abu di atas menggunakan krayon / pensil warna!")

    rx_f = mx + pw * 0.65
    rw_f = pw * 0.35
    c.setFillColor(colors.white)
    c.setStrokeColor(colors.HexColor("#CBD5E1"))
    c.roundRect(rx_f, fy, rw_f, fh, 6, fill=1, stroke=1)

    c.setFillColor(colors.HexColor("#1E293B"))
    c.setFont("Helvetica-Bold", 8.5)
    c.drawCentredString(rx_f + rw_f/2, fy + fh - 14, "⭐ Nilai Pelangi ⭐")

    c.setFillColor(colors.HexColor("#CBD5E1"))
    c.setFont("Helvetica-Bold", 13)
    c.drawCentredString(rx_f + rw_f/2, fy + fh - 34, "★ ★ ★ ★ ★")

    c.setFillColor(colors.HexColor("#64748B"))
    c.setFont("Helvetica", 7.5)
    c.drawCentredString(rx_f + rw_f/2, fy + 10, "Paraf Ortu: ____________")

    c.showPage()
    c.save()
    print("PDF with FAINT LIGHT GRAY tracing letters generated cleanly!")

if __name__ == "__main__":
    create_pdf()
