import os
import io
import re
import sys
import telebot
from flask import Flask, request, jsonify

# डॉक्यूमेंट लाइब्रेरीज़
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

from pptx import Presentation
from pptx.util import Inches as PptInches, Pt as PptPt
from pptx.dml.color import RGBColor as PptRGBColor
from pptx.enum.text import PP_ALIGN

from weasyprint import HTML

# ==========================================
# 1. कॉन्फ़िगरेशन
# ==========================================
BOT_TOKEN = os.getenv("BOT_TOKEN")
WEBHOOK_URL = os.getenv("WEBHOOK_URL", "https://mcq-converter-bot.onrender.com")

if not BOT_TOKEN:
    sys.exit("❌ Error: Render Environment Variables में 'BOT_TOKEN' सेट करें!")

bot = telebot.TeleBot(BOT_TOKEN, threaded=False)
app = Flask(__name__)


# ==========================================
# 2. टेक्स्ट और MCQ पार्सिंग
# ==========================================
def extract_text_from_docx(file_bytes):
    doc = Document(io.BytesIO(file_bytes))
    full_text = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    for table in doc.tables:
        for row in table.rows:
            row_data = [cell.text.strip() for cell in row.cells if cell.text.strip()]
            if row_data:
                full_text.append(" | ".join(row_data))
    return "\n".join(full_text)


def parse_mcqs(text):
    raw_blocks = re.split(r'\n(?=Question:)', text.strip(), flags=re.IGNORECASE)
    parsed = []

    for block in raw_blocks:
        if not block.strip():
            continue

        q_match = re.search(r'Question:\s*(.*?)(?=\([a-d]\)|Answer:|$)', block, re.DOTALL | re.IGNORECASE)
        ans_match = re.search(r'Answer:\s*([a-d])', block, re.IGNORECASE)
        sol_match = re.search(r'Solution:\s*(.*?)(?=Key Points:|Positive Marks:|$)', block, re.DOTALL | re.IGNORECASE)
        kp_match = re.search(r'Key Points:\s*(.*?)(?=Positive Marks:|Question:|$)', block, re.DOTALL | re.IGNORECASE)
        pos_match = re.search(r'Positive Marks:\s*(\d+)', block, re.IGNORECASE)
        neg_match = re.search(r'Negative Marks:\s*(\d+)', block, re.IGNORECASE)

        options = re.findall(r'\(([a-d])\)\s*([^(\n]+)', block, re.IGNORECASE)

        if q_match:
            parsed.append({
                "question": q_match.group(1).strip(),
                "options": options,
                "answer": ans_match.group(1).strip() if ans_match else "",
                "solution": sol_match.group(1).strip() if sol_match else "",
                "key_points": kp_match.group(1).strip() if kp_match else "",
                "pos_marks": pos_match.group(1) if pos_match else "1",
                "neg_marks": neg_match.group(1) if neg_match else "0"
            })
    return parsed


# ==========================================
# 3. PDF जनरेटर (इमेज लेआउट जैसा 2-कॉलम)
# ==========================================
def generate_pdf(mcqs, title="MCQ Test"):
    cards_html = []

    for i, item in enumerate(mcqs, 1):
        opts_html = "".join([f'<div class="opt"><b>({lbl})</b> {txt.strip()}</div>' for lbl, txt in item['options']])

        kp_bullets = ""
        if item['key_points']:
            lines = [ln.strip('• ').strip() for ln in item['key_points'].split('\n') if ln.strip()]
            bullet_items = " • ".join(lines)
            kp_bullets = f'<div class="key-points-box"><b>Key Points:</b> • {bullet_items}</div>'

        cards_html.append(f"""
        <div class="question-container">
            <div class="q-header">
                <div class="q-title"><span class="red-bar"></span> Question {i} / प्रश्न {i}</div>
                <div class="q-marks">Marks: +{item['pos_marks']}, -{item['neg_marks']}</div>
            </div>

            <div class="columns-grid">
                <div class="col">
                    <div class="question-text">{item['question']}</div>
                    <div class="options-group">{opts_html}</div>
                </div>
                <div class="col">
                    <div class="question-text">{item['question']}</div>
                    <div class="options-group">{opts_html}</div>
                </div>
            </div>

            <div class="solution-card">
                <div class="ans-line"><b>Answer: ({item['answer']})</b></div>
                <div class="sol-text"><b>Solution:</b> {item['solution']}</div>
                {kp_bullets}
            </div>
        </div>
        """)

    body_content = "\n".join(cards_html)

    html_template = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <style>
            @page {{
                size: A4;
                margin: 12mm 15mm 15mm 15mm;
                @top-center {{
                    content: "{title}";
                    font-size: 11px;
                    color: #888;
                    border-bottom: 1px solid #ddd;
                }}
            }}
            body {{
                font-family: 'Helvetica Neue', Helvetica, 'Noto Sans Devanagari', Arial, sans-serif;
                font-size: 11.5px;
                color: #111;
                line-height: 1.45;
            }}
            .question-container {{
                margin-bottom: 22px;
                page-break-inside: avoid;
            }}
            .q-header {{
                display: flex;
                justify-content: space-between;
                align-items: center;
                background-color: #fcf6f4;
                padding: 4px 8px;
                border-top: 1px solid #ebd3c8;
                border-bottom: 1px solid #ebd3c8;
                margin-bottom: 8px;
            }}
            .q-title {{
                font-size: 12.5px;
                font-weight: bold;
                color: #8b1e0f;
            }}
            .red-bar {{
                display: inline-block;
                width: 4px;
                height: 12px;
                background-color: #a82414;
                margin-right: 6px;
                vertical-align: middle;
            }}
            .q-marks {{
                font-size: 10.5px;
                color: #555;
            }}
            .columns-grid {{
                display: flex;
                gap: 18px;
                margin-bottom: 10px;
            }}
            .col {{
                width: 50%;
            }}
            .question-text {{
                font-weight: bold;
                margin-bottom: 6px;
                color: #0f172a;
            }}
            .options-group {{
                margin-top: 4px;
            }}
            .opt {{
                margin-bottom: 4px;
                color: #222;
            }}
            .solution-card {{
                border-left: 3.5px solid #0284c7;
                background-color: #f0f9ff;
                padding: 8px 12px;
                border-radius: 2px;
                font-size: 11px;
            }}
            .ans-line {{
                color: #0369a1;
                margin-bottom: 4px;
            }}
            .sol-text {{
                color: #334155;
                margin-bottom: 6px;
            }}
            .key-points-box {{
                background-color: #ffffff;
                border: 1px solid #bae6fd;
                padding: 6px 10px;
                border-radius: 3px;
                color: #475569;
                font-size: 10.5px;
            }}
        </style>
    </head>
    <body>
        {body_content}
    </body>
    </html>
    """

    out = io.BytesIO()
    HTML(string=html_template).write_pdf(out)
    out.seek(0)
    return out


# ==========================================
# 4. DOCX जनरेटर (किताब / ई-बुक बनाने के लिए)
# ==========================================
def generate_docx(mcqs, title="MCQ Book"):
    doc = Document()

    # पेज मार्जिन सेट करें
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(0.6)
        section.bottom_margin = Inches(0.6)
        section.left_margin = Inches(0.6)
        section.right_margin = Inches(0.6)

    # किताब का मुख्य शीर्षक
    title_p = doc.add_heading(title, level=1)
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph()

    for i, item in enumerate(mcqs, 1):
        # 1. हेडर टेबल (लाल रंग की पट्टी + Question + Marks)
        h_table = doc.add_table(rows=1, cols=2)
        h_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        h_table.autofit = False
        h_table.columns[0].width = Inches(5.0)
        h_table.columns[1].width = Inches(2.2)

        # बैकग्राउंड कलर सेट करें (#FCF6F4)
        shading = parse_xml(r'<w:shd {} w:fill="FCF6F4"/>'.format(nsdecls('w')))
        h_table.rows[0].cells[0]._tc.get_or_add_tcPr().append(shading)
        shading2 = parse_xml(r'<w:shd {} w:fill="FCF6F4"/>'.format(nsdecls('w')))
        h_table.rows[0].cells[1]._tc.get_or_add_tcPr().append(shading2)

        # Question Title
        cell_l = h_table.rows[0].cells[0].paragraphs[0]
        r_bar = cell_l.add_run("▌ ")
        r_bar.font.color.rgb = RGBColor(168, 36, 20)
        r_title = cell_l.add_run(f"Question {i} / प्रश्न {i}")
        r_title.bold = True
        r_title.font.color.rgb = RGBColor(139, 30, 15)
        r_title.font.size = Pt(10.5)

        # Marks
        cell_r = h_table.rows[0].cells[1].paragraphs[0]
        cell_r.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        r_marks = cell_r.add_run(f"Marks: +{item['pos_marks']}, -{item['neg_marks']}")
        r_marks.font.size = Pt(9.5)
        r_marks.font.color.rgb = RGBColor(85, 85, 85)

        # 2. 2-कॉलम टेबल (प्रश्न और विकल्प)
        c_table = doc.add_table(rows=1, cols=2)
        c_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        c_table.columns[0].width = Inches(3.6)
        c_table.columns[1].width = Inches(3.6)

        # Left Column
        p_left = c_table.rows[0].cells[0].paragraphs[0]
        q_run_l = p_left.add_run(f"{item['question']}\n")
        q_run_l.bold = True
        q_run_l.font.size = Pt(10)
        for lbl, txt in item['options']:
            p_left.add_run(f"({lbl}) {txt.strip()}\n").font.size = Pt(9.5)

        # Right Column (Bilingual)
        p_right = c_table.rows[0].cells[1].paragraphs[0]
        q_run_r = p_right.add_run(f"{item['question']}\n")
        q_run_r.bold = True
        q_run_r.font.size = Pt(10)
        for lbl, txt in item['options']:
            p_right.add_run(f"({lbl}) {txt.strip()}\n").font.size = Pt(9.5)

        # 3. सॉल्यूशन बॉक्स
        sol_table = doc.add_table(rows=1, cols=1)
        sol_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        sol_table.columns[0].width = Inches(7.2)
        
        # नीला शेडिंग (#F0F9FF)
        shd_sol = parse_xml(r'<w:shd {} w:fill="F0F9FF"/>'.format(nsdecls('w')))
        sol_table.rows[0].cells[0]._tc.get_or_add_tcPr().append(shd_sol)

        sp = sol_table.rows[0].cells[0].paragraphs[0]
        ans_r = sp.add_run(f"Answer: ({item['answer']})\n")
        ans_r.bold = True
        ans_r.font.color.rgb = RGBColor(3, 105, 161)
        ans_r.font.size = Pt(10)

        sol_label = sp.add_run("Solution: ")
        sol_label.bold = True
        sp.add_run(f"{item['solution']}\n").font.size = Pt(9.5)

        if item['key_points']:
            kp_label = sp.add_run("Key Points: ")
            kp_label.bold = True
            lines = [ln.strip('• ').strip() for ln in item['key_points'].split('\n') if ln.strip()]
            sp.add_run(" • " + " • ".join(lines)).font.size = Pt(9)

        doc.add_paragraph()  # स्पेसिंग

    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out


# ==========================================
# 5. PPTX जनरेटर (क्लास / प्रोजेक्टर के लिए 16:9 HD)
# ==========================================
def generate_pptx(mcqs, title="Class Presentation"):
    prs = Presentation()
    
    # 16:9 Widescreen स्लाइड साइज (13.33 x 7.5 Inches)
    prs.slide_width = PptInches(13.33)
    prs.slide_height = PptInches(7.5)
    blank_slide_layout = prs.slide_layouts[6]

    for i, item in enumerate(mcqs, 1):
        slide = prs.slides.add_slide(blank_slide_layout)

        # 1. टॉप हेडर बार
        header_box = slide.shapes.add_textbox(PptInches(0.6), PptInches(0.4), PptInches(12.13), PptInches(0.6))
        tf_h = header_box.text_frame
        tf_h.word_wrap = True
        p_h = tf_h.paragraphs[0]
        run_h1 = p_h.add_run()
        run_h1.text = f"▌ Question {i} / प्रश्न {i}"
        run_h1.font.bold = True
        run_h1.font.size = PptPt(18)
        run_h1.font.color.rgb = PptRGBColor(168, 36, 20)

        run_h2 = p_h.add_run()
        run_h2.text = f"                                                                       Marks: +{item['pos_marks']}, -{item['neg_marks']}"
        run_h2.font.size = PptPt(14)
        run_h2.font.color.rgb = PptRGBColor(100, 100, 100)

        # 2. प्रश्न बॉक्स (बड़ा टेक्स्ट ताकि छात्रों को साफ़ दिखे)
        q_box = slide.shapes.add_textbox(PptInches(0.8), PptInches(1.1), PptInches(11.7), PptInches(1.5))
        tf_q = q_box.text_frame
        tf_q.word_wrap = True
        p_q = tf_q.paragraphs[0]
        p_q.text = item['question']
        p_q.font.bold = True
        p_q.font.size = PptPt(20)
        p_q.font.color.rgb = PptRGBColor(15, 23, 42)

        # 3. विकल्प बॉक्स
        opt_box = slide.shapes.add_textbox(PptInches(0.8), PptInches(2.6), PptInches(11.7), PptInches(2.2))
        tf_o = opt_box.text_frame
        tf_o.word_wrap = True
        for idx, (lbl, txt) in enumerate(item['options']):
            p_o = tf_o.paragraphs[0] if idx == 0 else tf_o.add_paragraph()
            p_o.text = f"({lbl})  {txt.strip()}"
            p_o.font.size = PptPt(17)
            p_o.font.color.rgb = PptRGBColor(30, 41, 59)
            p_o.space_after = PptPt(8)

        # 4. बॉटम सॉल्यूशन बार (क्लास में उत्तर और व्याख्या समझाने के लिए)
        sol_box = slide.shapes.add_textbox(PptInches(0.6), PptInches(5.0), PptInches(12.13), PptInches(2.1))
        tf_s = sol_box.text_frame
        tf_s.word_wrap = True

        p_ans = tf_s.paragraphs[0]
        p_ans.text = f"✔ Answer: ({item['answer']})"
        p_ans.font.bold = True
        p_ans.font.size = PptPt(16)
        p_ans.font.color.rgb = PptRGBColor(3, 105, 161)

        if item['solution']:
            p_sol = tf_s.add_paragraph()
            p_sol.text = f"Solution: {item['solution']}"
            p_sol.font.size = PptPt(13)
            p_sol.font.color.rgb = PptRGBColor(71, 85, 105)

        if item['key_points']:
            p_kp = tf_s.add_paragraph()
            lines = [ln.strip('• ').strip() for ln in item['key_points'].split('\n') if ln.strip()]
            p_kp.text = f"Key Points: • " + " • ".join(lines)
            p_kp.font.size = PptPt(12)
            p_kp.font.color.rgb = PptRGBColor(100, 116, 139)

    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out


# ==========================================
# 6. टेलीग्राम मैसेज हैंडलर्स
# ==========================================
@bot.message_handler(commands=['start', 'help'])
def send_welcome(message):
    bot.reply_to(
        message,
        "👋 *MCQ All-in-One Generator Bot*\n\n"
        "मुझे अपनी MCQ `.docx` या `.txt` फ़ाइल भेजें। मैं आपको एक साथ 3 फ़ॉर्मैट दूंगा:\n"
        "1. 🖥️ **PPTX** - क्लास में पढ़ाने / लाइव सेशन के लिए\n"
        "2. 📕 **PDF** - छात्रों के साथ शेयर करने के लिए\n"
        "3. 📘 **DOCX** - किताब या ई-बुक बनाने के लिए",
        parse_mode="Markdown"
    )


@bot.message_handler(content_types=['document'])
def handle_docs(message):
    file_name = message.document.file_name or "MCQs.docx"
    ext = os.path.splitext(file_name)[1].lower()

    if ext not in ['.docx', '.txt']:
        bot.reply_to(message, "⚠️ केवल `.docx` या `.txt` फ़ाइल स्वीकार्य है।")
        return

    wait_msg = bot.reply_to(message, f"⏳ `{file_name}` से PPT, PDF और DOCX तैयार हो रहे हैं...", parse_mode="Markdown")

    try:
        file_info = bot.get_file(message.document.file_id)
        raw_bytes = bot.download_file(file_info.file_path)

        text = extract_text_from_docx(raw_bytes) if ext == '.docx' else raw_bytes.decode('utf-8', errors='ignore')
        mcqs = parse_mcqs(text)

        if not mcqs:
            bot.edit_message_text("❌ फ़ाइल में कोई प्रश्न नहीं मिला।", message.chat.id, wait_msg.message_id)
            return

        base_name = os.path.splitext(file_name)[0]

        # 1. PPTX भेजें (Class Presentation)
        ppt_data = generate_pptx(mcqs, title=base_name)
        ppt_data.name = f"{base_name}_Class.pptx"
        bot.send_document(message.chat.id, ppt_data, caption="🖥️ *PowerPoint (PPTX)* - क्लास में पढ़ाने के लिए", parse_mode="Markdown")

        # 2. PDF भेजें (Sharing)
        pdf_data = generate_pdf(mcqs, title=base_name)
        pdf_data.name = f"{base_name}_Share.pdf"
        bot.send_document(message.chat.id, pdf_data, caption="📕 *PDF Document* - छात्रों के साथ शेयर करने के लिए", parse_mode="Markdown")

        # 3. DOCX भेजें (Book / Printing)
        docx_data = generate_docx(mcqs, title=base_name)
        docx_data.name = f"{base_name}_Book.docx"
        bot.send_document(message.chat.id, docx_data, caption="📘 *Word (DOCX)* - किताब / प्रिंटिंग के लिए", parse_mode="Markdown")

        bot.delete_message(message.chat.id, wait_msg.message_id)

    except Exception as e:
        bot.edit_message_text(f"❌ प्रोसेस करने में त्रुटि: {str(e)}", message.chat.id, wait_msg.message_id)


# ==========================================
# 7. Flask Server & Webhook
# ==========================================
@app.route(f"/{BOT_TOKEN}", methods=['POST'])
def process_webhook():
    if request.headers.get('content-type') == 'application/json':
        bot.process_new_updates([telebot.types.Update.de_json(request.get_data().decode('utf-8'))])
        return 'OK', 200
    return 'Forbidden', 403


@app.route('/', methods=['GET', 'HEAD'])
def index():
    return "Bot is Live with PPTX, PDF and DOCX converters!", 200


def init_webhook():
    full_url = f"{WEBHOOK_URL.rstrip('/')}/{BOT_TOKEN}"
    try:
        if bot.get_webhook_info().url != full_url:
            bot.remove_webhook()
            bot.set_webhook(url=full_url)
    except Exception as err:
        print(f"Webhook error: {err}")

init_webhook()

if __name__ == '__main__':
    app.run(host="0.0.0.0", port=int(os.environ.get('PORT', 10000)))
        
