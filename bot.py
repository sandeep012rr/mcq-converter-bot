import os
import io
import re
import sys
import telebot
from flask import Flask, request, jsonify
from docx import Document
from weasyprint import HTML

BOT_TOKEN = os.getenv("BOT_TOKEN")
WEBHOOK_URL = os.getenv("WEBHOOK_URL", "https://mcq-converter-bot.onrender.com")

if not BOT_TOKEN:
    sys.exit("❌ Error: Render Environment Variables में 'BOT_TOKEN' सेट करें!")

bot = telebot.TeleBot(BOT_TOKEN, threaded=False)
app = Flask(__name__)


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


def generate_exact_pdf(mcqs, title="MCQ Test"):
    cards_html = []

    for i, item in enumerate(mcqs, 1):
        # विकल्प HTML
        opts_html = "".join([f'<div class="opt"><b>({lbl})</b> {txt.strip()}</div>' for lbl, txt in item['options']])

        # Key Points बुलेट पॉइंट्स
        kp_bullets = ""
        if item['key_points']:
            lines = [ln.strip('• ').strip() for ln in item['key_points'].split('\n') if ln.strip()]
            bullet_items = " • ".join(lines)
            kp_bullets = f'<div class="key-points-box"><b>Key Points:</b> • {bullet_items}</div>'

        cards_html.append(f"""
        <div class="question-container">
            <!-- हेडर बार -->
            <div class="q-header">
                <div class="q-title"><span class="red-bar"></span> Question {i} / प्रश्न {i}</div>
                <div class="q-marks">Marks: +{item['pos_marks']}, -{item['neg_marks']}</div>
            </div>

            <!-- 2 कॉलम: प्रश्न व विकल्प -->
            <div class="columns-grid">
                <div class="col">
                    <div class="question-text">{item['question']}</div>
                    <div class="options-group">{opts_html}</div>
                </div>
                <div class="col">
                    <!-- यदि आपके पास अनुवादित अंग्रेजी टेक्स्ट उपलब्ध है तो वह यहाँ आएगा -->
                    <div class="question-text">{item['question']}</div>
                    <div class="options-group">{opts_html}</div>
                </div>
            </div>

            <!-- सॉल्यूशन बॉक्स -->
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
# बॉट हैंडलर्स
# ==========================================
@bot.message_handler(commands=['start', 'help'])
def send_welcome(message):
    bot.reply_to(message, "👋 अपनी MCQ `.docx` या `.txt` फ़ाइल भेजें, मैं इसे इमेज वाले लेआउट में फॉर्मेट करके PDF दूंगा।")


@bot.message_handler(content_types=['document'])
def handle_docs(message):
    file_name = message.document.file_name or "MCQs.docx"
    ext = os.path.splitext(file_name)[1].lower()

    if ext not in ['.docx', '.txt']:
        bot.reply_to(message, "⚠️ केवल `.docx` या `.txt` फ़ाइल स्वीकार्य है।")
        return

    wait_msg = bot.reply_to(message, f"⏳ `{file_name}` को फॉर्मेट किया जा रहा है...", parse_mode="Markdown")

    try:
        file_info = bot.get_file(message.document.file_id)
        raw_bytes = bot.download_file(file_info.file_path)

        text = extract_text_from_docx(raw_bytes) if ext == '.docx' else raw_bytes.decode('utf-8', errors='ignore')
        mcqs = parse_mcqs(text)

        if not mcqs:
            bot.edit_message_text("❌ फ़ाइल में कोई प्रश्न नहीं मिला।", message.chat.id, wait_msg.message_id)
            return

        pdf = generate_exact_pdf(mcqs, title=os.path.splitext(file_name)[0])
        pdf.name = f"{os.path.splitext(file_name)[0]}_Formatted.pdf"

        bot.send_document(message.chat.id, pdf, caption=f"✅ आपके फॉर्मेट में तैयार PDF ({len(mcqs)} प्रश्न)")
        bot.delete_message(message.chat.id, wait_msg.message_id)

    except Exception as e:
        bot.edit_message_text(f"❌ एरर: {str(e)}", message.chat.id, wait_msg.message_id)


@app.route(f"/{BOT_TOKEN}", methods=['POST'])
def process_webhook():
    if request.headers.get('content-type') == 'application/json':
        bot.process_new_updates([telebot.types.Update.de_json(request.get_data().decode('utf-8'))])
        return 'OK', 200
    return 'Forbidden', 403


@app.route('/', methods=['GET', 'HEAD'])
def index():
    return "Bot is Live", 200


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
