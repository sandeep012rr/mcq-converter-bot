import os
import io
import re
import sys
import requests
import threading
import telebot
from flask import Flask, request

# Document aur Presentation Libraries
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

from pptx import Presentation
from pptx.util import Inches as PptInches, Pt as PptPt
from pptx.dml.color import RGBColor as PptRGBColor

from weasyprint import HTML

# ==========================================
# 1. Configuration
# ==========================================
BOT_TOKEN = os.getenv("BOT_TOKEN")
WEBHOOK_URL = os.getenv("WEBHOOK_URL", "https://mcq-converter-bot.onrender.com")

if not BOT_TOKEN:
    sys.exit("Error: Render Environment Variables me 'BOT_TOKEN' set karein!")

bot = telebot.TeleBot(BOT_TOKEN, threaded=False)
app = Flask(__name__)


# ==========================================
# 2. Reliable Cloud Translator (Direct API)
# ==========================================
def translate_to_en(text):
    """Render Cloud IP par bina block hue Hindi ko English me convert karta hai"""
    if not text or not text.strip():
        return ""
    try:
        clean_text = text.strip()
        url = "https://translate.googleapis.com/translate_a/single"
        params = {
            "client": "gtx",
            "sl": "hi",
            "tl": "en",
            "dt": "t",
            "q": clean_text
        }
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
        }
        resp = requests.get(url, params=params, headers=headers, timeout=10)
        if resp.status_code == 200:
            data = resp.json()
            parts = [seg[0] for seg in data[0] if seg and seg[0]]
            res = "".join(parts).strip()
            if res:
                return res
        return clean_text
    except Exception as e:
        print(f"Translation Error: {e}")
        return text


# ==========================================
# 3. Extraction aur Parsing
# ==========================================
def extract_text_from_docx(file_bytes):
    doc = Document(io.BytesIO(file_bytes))
    lines = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    for table in doc.tables:
        for row in table.rows:
            row_items = [c.text.strip() for c in row.cells if c.text.strip()]
            if row_items:
                lines.append(" | ".join(row_items))
    return "\n".join(lines)


def parse_mcqs_bilingual(text):
    blocks = re.split(r'\n(?=Question:)', text.strip(), flags=re.IGNORECASE)
    parsed = []

    for block in blocks:
        if not block.strip():
            continue

        q_match = re.search(r'Question:\s*(.*?)(?=\([a-d]\)|Answer:|$)', block, re.DOTALL | re.IGNORECASE)
        ans_match = re.search(r'Answer:\s*([a-d])', block, re.IGNORECASE)
        sol_match = re.search(r'Solution:\s*(.*?)(?=Key Points:|Positive Marks:|$)', block, re.DOTALL | re.IGNORECASE)
        kp_match = re.search(r'Key Points:\s*(.*?)(?=Positive Marks:|Question:|$)', block, re.DOTALL | re.IGNORECASE)
        pos_match = re.search(r'Positive Marks:\s*(\d+)', block, re.IGNORECASE)
        neg_match = re.search(r'Negative Marks:\s*(\d+)', block, re.IGNORECASE)

        raw_opts = re.findall(r'\(([a-d])\)\s*([^(\n]+)', block, re.IGNORECASE)

        if q_match:
            q_hi = q_match.group(1).strip()
            # English Translation of Question
            q_en = translate_to_en(q_hi)

            opts_hi = []
            opts_en = []
            for lbl, opt_val in raw_opts:
                val_clean = opt_val.strip()
                opts_hi.append((lbl, val_clean))
                # English Translation of Option
                opts_en.append((lbl, translate_to_en(val_clean)))

            parsed.append({
                "q_hi": q_hi,
                "q_en": q_en,
                "opts_hi": opts_hi,
                "opts_en": opts_en,
                "answer": ans_match.group(1).strip() if ans_match else "",
                "solution": sol_match.group(1).strip() if sol_match else "",
                "key_points": kp_match.group(1).strip() if kp_match else "",
                "pos_marks": pos_match.group(1) if pos_match else "1",
                "neg_marks": neg_match.group(1) if neg_match else "0"
            })
    return parsed


# ==========================================
# 4. Exact 2-Column PDF Generator
# ==========================================
def generate_pdf(mcqs, title="MCQ Test"):
    cards = []

    for i, item in enumerate(mcqs, 1):
        hi_options = "".join([f'<div class="opt"><b>({l})</b> {t}</div>' for l, t in item['opts_hi']])
        en_options = "".join([f'<div class="opt"><b>({l})</b> {t}</div>' for l, t in item['opts_en']])

        kp_html = ""
        if item['key_points']:
            bullets = [ln.strip('• ').strip() for ln in item['key_points'].split('\n') if ln.strip()]
            kp_html = f'<div class="key-points-box"><b>Key Points:</b> • {" • ".join(bullets)}</div>'

        cards.append(f"""
        <div class="question-container">
            <div class="q-header">
                <div class="q-title"><span class="red-bar"></span> Question {i} / प्रश्न {i}</div>
                <div class="q-marks">Marks: +{item['pos_marks']}, -{item['neg_marks']}</div>
            </div>

            <div class="columns-grid">
                <div class="col">
                    <div class="question-text">{item['q_hi']}</div>
                    <div class="options-group">{hi_options}</div>
                </div>
                <div class="col">
                    <div class="question-text">{item['q_en']}</div>
                    <div class="options-group">{en_options}</div>
                </div>
            </div>

            <div class="solution-card">
                <div class="ans-line"><b>Answer: ({item['answer']})</b></div>
                <div class="sol-text"><b>Solution:</b> {item['solution']}</div>
                {kp_html}
            </div>
        </div>
        """)

    full_html = f"""
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
        {"".join(cards)}
    </body>
    </html>
    """

    out = io.BytesIO()
    HTML(string=full_html).write_pdf(out)
    out.seek(0)
    return out


# ==========================================
# 5. Bilingual DOCX Generator (Book)
# ==========================================
def generate_docx(mcqs, title="MCQ Book"):
    doc = Document()

    for sec in doc.sections:
        sec.top_margin = Inches(0.6)
        sec.bottom_margin = Inches(0.6)
        sec.left_margin = Inches(0.6)
        sec.right_margin = Inches(0.6)

    title_p = doc.add_heading(title, level=1)
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph()

    for i, item in enumerate(mcqs, 1):
        # Header Box
        h_tab = doc.add_table(rows=1, cols=2)
        h_tab.alignment = WD_TABLE_ALIGNMENT.CENTER
        h_tab.autofit = False
        h_tab.columns[0].width = Inches(5.0)
        h_tab.columns[1].width = Inches(2.2)

        shd1 = parse_xml(r'<w:shd {} w:fill="FCF6F4"/>'.format(nsdecls('w')))
        h_tab.rows[0].cells[0]._tc.get_or_add_tcPr().append(shd1)
        shd2 = parse_xml(r'<w:shd {} w:fill="FCF6F4"/>'.format(nsdecls('w')))
        h_tab.rows[0].cells[1]._tc.get_or_add_tcPr().append(shd2)

        c_l = h_tab.rows[0].cells[0].paragraphs[0]
        rb = c_l.add_run("▌ ")
        rb.font.color.rgb = RGBColor(168, 36, 20)
        rt = c_l.add_run(f"Question {i} / प्रश्न {i}")
        rt.bold = True
        rt.font.color.rgb = RGBColor(139, 30, 15)
        rt.font.size = Pt(10.5)

        c_r = h_tab.rows[0].cells[1].paragraphs[0]
        c_r.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        rm = c_r.add_run(f"Marks: +{item['pos_marks']}, -{item['neg_marks']}")
        rm.font.size = Pt(9.5)
        rm.font.color.rgb = RGBColor(85, 85, 85)

        # 2-Column Table
        q_tab = doc.add_table(rows=1, cols=2)
        q_tab.alignment = WD_TABLE_ALIGNMENT.CENTER
        q_tab.columns[0].width = Inches(3.6)
        q_tab.columns[1].width = Inches(3.6)

        # Left: Hindi
        pl = q_tab.rows[0].cells[0].paragraphs[0]
        ql = pl.add_run(f"{item['q_hi']}\n")
        ql.bold = True
        ql.font.size = Pt(10)
        for l, t in item['opts_hi']:
            pl.add_run(f"({l}) {t}\n").font.size = Pt(9.5)

        # Right: English
        pr = q_tab.rows[0].cells[1].paragraphs[0]
        qr = pr.add_run(f"{item['q_en']}\n")
        qr.bold = True
        qr.font.size = Pt(10)
        for l, t in item['opts_en']:
            pr.add_run(f"({l}) {t}\n").font.size = Pt(9.5)

        # Solution Box
        s_tab = doc.add_table(rows=1, cols=1)
        s_tab.alignment = WD_TABLE_ALIGNMENT.CENTER
        s_tab.columns[0].width = Inches(7.2)

        shd_s = parse_xml(r'<w:shd {} w:fill="F0F9FF"/>'.format(nsdecls('w')))
        s_tab.rows[0].cells[0]._tc.get_or_add_tcPr().append(shd_s)

        sp = s_tab.rows[0].cells[0].paragraphs[0]
        ans_run = sp.add_run(f"Answer: ({item['answer']})\n")
        ans_run.bold = True
        ans_run.font.color.rgb = RGBColor(3, 105, 161)
        ans_run.font.size = Pt(10)

        sl = sp.add_run("Solution: ")
        sl.bold = True
        sp.add_run(f"{item['solution']}\n").font.size = Pt(9.5)

        if item['key_points']:
            kl = sp.add_run("Key Points: ")
            kl.bold = True
            lines = [ln.strip('• ').strip() for ln in item['key_points'].split('\n') if ln.strip()]
            sp.add_run(" • " + " • ".join(lines)).font.size = Pt(9)

        doc.add_paragraph()

    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out


# ==========================================
# 6. Bilingual PPTX Generator (Class Slides)
# ==========================================
def generate_pptx(mcqs, title="Class Presentation"):
    prs = Presentation()
    prs.slide_width = PptInches(13.33)
    prs.slide_height = PptInches(7.5)
    layout = prs.slide_layouts[6]

    for i, item in enumerate(mcqs, 1):
        slide = prs.slides.add_slide(layout)

        # Header
        h_box = slide.shapes.add_textbox(PptInches(0.6), PptInches(0.3), PptInches(12.13), PptInches(0.5))
        p_h = h_box.text_frame.paragraphs[0]
        r1 = p_h.add_run()
        r1.text = f"▌ Question {i} / प्रश्न {i}"
        r1.font.bold = True
        r1.font.size = PptPt(16)
        r1.font.color.rgb = PptRGBColor(168, 36, 20)

        # Hindi (Left)
        b_hi = slide.shapes.add_textbox(PptInches(0.6), PptInches(0.9), PptInches(5.9), PptInches(3.8))
        tf_hi = b_hi.text_frame
        tf_hi.word_wrap = True
        pq_hi = tf_hi.paragraphs[0]
        pq_hi.text = item['q_hi']
        pq_hi.font.bold = True
        pq_hi.font.size = PptPt(15)
        pq_hi.font.color.rgb = PptRGBColor(15, 23, 42)
        pq_hi.space_after = PptPt(8)

        for l, t in item['opts_hi']:
            po = tf_hi.add_paragraph()
            po.text = f"({l}) {t}"
            po.font.size = PptPt(13)
            po.font.color.rgb = PptRGBColor(30, 41, 59)
            po.space_after = PptPt(4)

        # English (Right)
        b_en = slide.shapes.add_textbox(PptInches(6.8), PptInches(0.9), PptInches(5.9), PptInches(3.8))
        tf_en = b_en.text_frame
        tf_en.word_wrap = True
        pq_en = tf_en.paragraphs[0]
        pq_en.text = item['q_en']
        pq_en.font.bold = True
        pq_en.font.size = PptPt(15)
        pq_en.font.color.rgb = PptRGBColor(15, 23, 42)
        pq_en.space_after = PptPt(8)

        for l, t in item['opts_en']:
            po = tf_en.add_paragraph()
            po.text = f"({l}) {t}"
            po.font.size = PptPt(13)
            po.font.color.rgb = PptRGBColor(30, 41, 59)
            po.space_after = PptPt(4)

        # Solution Box
        s_box = slide.shapes.add_textbox(PptInches(0.6), PptInches(4.9), PptInches(12.13), PptInches(2.2))
        tf_s = s_box.text_frame
        tf_s.word_wrap = True

        pa = tf_s.paragraphs[0]
        pa.text = f"✔ Answer: ({item['answer']})"
        pa.font.bold = True
        pa.font.size = PptPt(15)
        pa.font.color.rgb = PptRGBColor(3, 105, 161)

        if item['solution']:
            ps = tf_s.add_paragraph()
            ps.text = f"Solution: {item['solution']}"
            ps.font.size = PptPt(12)
            ps.font.color.rgb = PptRGBColor(71, 85, 105)

        if item['key_points']:
            pk = tf_s.add_paragraph()
            lines = [ln.strip('• ').strip() for ln in item['key_points'].split('\n') if ln.strip()]
            pk.text = "Key Points: • " + " • ".join(lines)
            pk.font.size = PptPt(11)
            pk.font.color.rgb = PptRGBColor(100, 116, 139)

    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out


# ==========================================
# 7. Background Async Processor
# ==========================================
def process_and_send(chat_id, file_id, file_name, ext):
    try:
        msg = bot.send_message(chat_id, f"⏳ `{file_name}` ka Hindi se English me accurate translation shuru ho gaya hai...", parse_mode="Markdown")

        f_info = bot.get_file(file_id)
        raw = bot.download_file(f_info.file_path)

        text = extract_text_from_docx(raw) if ext == '.docx' else raw.decode('utf-8', errors='ignore')
        mcqs = parse_mcqs_bilingual(text)

        if not mcqs:
            bot.edit_message_text("❌ File me koi question nahi mila.", chat_id, msg.message_id)
            return

        base = os.path.splitext(file_name)[0]

        # 1. PPTX
        ppt = generate_pptx(mcqs, title=base)
        ppt.name = f"{base}_Class.pptx"
        bot.send_document(chat_id, ppt, caption="🖥️ *PowerPoint (Bilingual)* - Class lene ke liye", parse_mode="Markdown")

        # 2. PDF
        pdf = generate_pdf(mcqs, title=base)
        pdf.name = f"{base}_Share.pdf"
        bot.send_document(chat_id, pdf, caption="📕 *PDF Document (Bilingual)* - Share karne ke liye", parse_mode="Markdown")

        # 3. DOCX
        docx = generate_docx(mcqs, title=base)
        docx.name = f"{base}_Book.docx"
        bot.send_document(chat_id, docx, caption="📘 *Word Document (Bilingual)* - Book / Print ke liye", parse_mode="Markdown")

        bot.delete_message(chat_id, msg.message_id)

    except Exception as e:
        bot.send_message(chat_id, f"❌ Error: {str(e)}")


# ==========================================
# 8. Webhook & Handlers
# ==========================================
@bot.message_handler(commands=['start', 'help'])
def start_cmd(message):
    bot.reply_to(
        message,
        "👋 *MCQ Bilingual Converter Bot Ready Hai!*\n\n"
        "Apni `.docx` ya `.txt` file bhejein. Bot 3 files banakar bhejega:\n"
        "1. 🖥️ **PPTX** (Widescreen Class Presentation)\n"
        "2. 📕 **PDF** (2-Column Bilingual Format)\n"
        "3. 📘 **DOCX** (Printable Book Format)",
        parse_mode="Markdown"
    )


@bot.message_handler(content_types=['document'])
def handle_doc(message):
    fname = message.document.file_name or "MCQs.docx"
    ext = os.path.splitext(fname)[1].lower()

    if ext not in ['.docx', '.txt']:
        bot.reply_to(message, "⚠️ Keval `.docx` ya `.txt` file bhejein.")
        return

    # Background execution taaki duplicate 6 files na banein
    threading.Thread(
        target=process_and_send,
        args=(message.chat.id, message.document.file_id, fname, ext)
    ).start()


@app.route(f"/{BOT_TOKEN}", methods=['POST'])
def webhook():
    if request.headers.get('content-type') == 'application/json':
        data = request.get_data().decode('utf-8')
        bot.process_new_updates([telebot.types.Update.de_json(data)])
        return 'OK', 200
    return 'Forbidden', 403


@app.route('/', methods=['GET', 'HEAD'])
def index():
    return "Bot is Live with Direct Web Translation!", 200


def init_webhook():
    url = f"{WEBHOOK_URL.rstrip('/')}/{BOT_TOKEN}"
    try:
        if bot.get_webhook_info().url != url:
            bot.remove_webhook()
            bot.set_webhook(url=url)
    except Exception as e:
        print(f"Webhook setup warning: {e}")

init_webhook()

if __name__ == '__main__':
    app.run(host="0.0.0.0", port=int(os.environ.get('PORT', 10000)))
