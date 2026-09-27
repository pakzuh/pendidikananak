import fitz

def render_preview():
    doc = fitz.open("lembar_menebali_huruf_hewan_pelangi.pdf")
    page = doc[0]
    pix = page.get_pixmap(dpi=150)
    pix.save("preview_lembar_pelangi.png")
    print("PNG preview generated successfully!")

if __name__ == "__main__":
    render_preview()
