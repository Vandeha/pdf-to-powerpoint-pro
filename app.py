import os, io, fitz
from flask import Flask, request, send_file, send_from_directory
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from flask_cors import CORS

app = Flask(__name__, static_folder='.')
CORS(app)

def srgb_to_rgb(srgb):
    """Converts PDF integer color to (R, G, B) tuple."""
    if srgb is None: return (0, 0, 0)
    return (srgb >> 16 & 255, srgb >> 8 & 255, srgb & 255)

@app.route('/')
def index():
    return send_from_directory('.', 'index.html')

@app.route("/convert", methods=["POST"])
def convert():
    file = request.files.get("file")
    if not file: return "No file", 400

    doc = fitz.open(stream=file.read(), filetype="pdf")
    prs = Presentation()
    
    # Standard PDF Size (8.5 x 11 inches)
    prs.slide_width = Inches(8.5)
    prs.slide_height = Inches(11)

    for page in doc:
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        
        # 1. Process Text with Color and Size
        dict_data = page.get_text("dict")
        for block in dict_data["blocks"]:
            if "lines" in block:
                for line in block["lines"]:
                    for span in line["spans"]:
                        l, t, r, b = span["bbox"]
                        # Create Textbox
                        txt_box = slide.shapes.add_textbox(Inches(l/72), Inches(t/72), Inches((r-l)/72), Inches((b-t)/72))
                        tf = txt_box.text_frame
                        tf.word_wrap = True
                        
                        p = tf.paragraphs[0]
                        p.text = span["text"]
                        
                        # Apply Font Size and Color
                        font = p.font
                        font.size = Pt(span["size"])
                        r_val, g_val, b_val = srgb_to_rgb(span["color"])
                        font.color.rgb = RGBColor(r_val, g_val, b_val)

        # 2. Process Standalone Images
        for img in page.get_images(full=True):
            xref = img[0]
            base_image = doc.extract_image(xref)
            img_rects = page.get_image_rects(xref)
            if img_rects:
                rect = img_rects[0]
                slide.shapes.add_picture(
                    io.BytesIO(base_image["image"]), 
                    Inches(rect.x0/72), 
                    Inches(rect.y0/72), 
                    width=Inches(rect.width/72)
                )

    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    return send_file(output, mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation", as_attachment=True, download_name="Converted_Pro.pptx")

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)