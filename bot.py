import os
import re
from threading import Thread
from flask import Flask
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters, ContextTypes
from docx import Document
import pdfkit

# ==========================================
# 1. Telegram Bot Token
# ==========================================
BOT_TOKEN = "8903776742:AAGeYC3UemM-JsuHZ2Af3dmTRAaC7THwcP0"

# ==========================================
# 2. Render Port Listener (Flask)
# ==========================================
web_app = Flask(__name__)

@web_app.route('/')
def home():
    return "Bot is online and running!"

def run_web():
    port = int(os.environ.get("PORT", 10000))
    web_app.run(host="0.0.0.0", port=port)

server_thread = Thread(target=run_web, daemon=True)
server_thread.start()

# ==========================================
# 3. DOCX Parser
# ==========================================
def clean_text(txt):
    if not txt:
        return ""
    txt = re.sub(r'[\r\t]', ' ', txt)
    txt = re.sub(r' +', ' ', txt)
    return txt.strip()

def parse_docx(file_path):
    doc = Document(file_path)
    full_text = "\n".join([p.text.strip() for p in doc.paragraphs if p.text.strip()])
    
    raw_blocks = re.split(r'\n(?=Question:|\d+\s*/\s*प्रश्न)', full_text)
    questions = []
    
    for block in raw_blocks:
        if not re.search(r'(Question:|\d+\s*/\s*प्रश्न)', block):
            continue
            
        q_data = {}
        
        # Question text
        q_match = re.search(r'(?:Question:\s*|\d+\s*/\s*प्रश्न\s*\d*\s*)(.*?)(?=\n\([a-d]\)|\nAnswer:)', block, re.DOTALL)
        raw_q = q_match.group(1).strip() if q_match else ""
        
        # English translation extraction (agar slash ya bracket me ho)
        q_data['q_hi'] = clean_text(raw_q)
        q_data['q_en'] = clean_text(raw_q)  # Default fallback
        
        # Options
        opts = re.findall(r'\(([a-d])\)\s*(.*?)(?=\n\([a-d]\)|\nAnswer:|\Z)', block, re.DOTALL)
        q_data['opts'] = {k.lower(): clean_text(v) for k, v in opts}
        
        # Answer
        ans_match = re.search(r'Answer:\s*([a-d])', block, re.IGNORECASE)
        q_data['ans'] = ans_match.group(1).upper() if ans_match else ""
        
        # Solution
        sol_match = re.search(r'Solution:\s*(.*?)(?=\nKey Points:|\nPositive Marks:|\Z)', block, re.DOTALL)
        q_data['sol'] = clean_text(sol_match.group(1)) if sol_match else ""
        
        # Key Points
        kp_match = re.search(r'Key Points:\s*(.*?)(?=\nPositive Marks:|\Z)', block, re.DOTALL)
        q_data['kp'] = clean_text(kp_match.group(1)) if kp_match else ""
        
        questions.append(q_data)
        
    return questions

# ==========================================
# 4. WebKit-Based PDF Generator (Natural Devanagari)
# ==========================================
def generate_pdf(questions, output_pdf):
    html_content = """<!DOCTYPE html>
<html lang="hi">
<head>
<meta charset="utf-8">
<style>
  @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+Devanagari:wght@400;600;700&display=swap');
  
  body {
    font-family: 'Noto Sans Devanagari', 'Times New Roman', serif;
    font-size: 13pt;
    line-height: 1.45;
    color: #111;
    margin: 0;
    padding: 0;
  }
  .q-card {
    page-break-inside: avoid;
    margin-bottom: 22px;
    border-bottom: 1px solid #e0e0e0;
    padding-bottom: 14px;
  }
  .q-header {
    border-left: 5px solid #b71c1c;
    border-bottom: 1.5px solid #b71c1c;
    padding: 4px 8px;
    margin-bottom: 8px;
  }
  .q-header table {
    width: 100%;
    border-collapse: collapse;
  }
  .q-title {
    color: #b71c1c;
    font-weight: 700;
    font-size: 13.5pt;
    text-align: left;
  }
  .q-marks {
    color: #555;
    font-size: 11pt;
    text-align: right;
  }
  .bilingual-table {
    width: 100%;
    border-collapse: collapse;
    margin-bottom: 8px;
  }
  .col-hindi {
    width: 50%;
    vertical-align: top;
    padding-right: 12px;
  }
  .col-eng {
    width: 50%;
    vertical-align: top;
    padding-left: 12px;
    border-left: 1px solid #ccc;
  }
  .q-text {
    font-weight: 700;
    margin-bottom: 6px;
  }
  .opt-line {
    margin: 3px 0;
  }
  .opt-key {
    font-weight: 700;
    color: #0d47a1;
  }
  .ans-box {
    background-color: #f1f8e9;
    border: 1.5px solid #2e7d32;
    padding: 6px 10px;
    color: #2e7d32;
    font-weight: 700;
    margin-bottom: 6px;
    border-radius: 3px;
  }
  .sol-box {
    background-color: #f9fbe7;
    border: 1px solid #cddc39;
    padding: 6px 10px;
    margin-bottom: 6px;
    border-radius: 3px;
  }
  .kp-box {
    background-color: #f5f5f5;
    border-left: 4px solid #1976d2;
    padding: 6px 10px;
    border-radius: 2px;
  }
  .kp-title {
    color: #1976d2;
    font-weight: 700;
    margin-bottom: 4px;
  }
</style>
</head>
<body>
"""

    for idx, q in enumerate(questions, start=1):
        opts_left_html = ""
        opts_right_html = ""
        for key in ['a', 'b', 'c', 'd']:
            val = q['opts'].get(key, '')
            opts_left_html += f'<div class="opt-line"><span class="opt-key">({key})</span> {val}</div>'
            opts_right_html += f'<div class="opt-line"><span class="opt-key">({key})</span> {val}</div>'
            
        kp_formatted = q['kp'].replace('\n', '<br>')
        
        html_content += f"""
<div class="q-card">
  <div class="q-header">
    <table>
      <tr>
        <td class="q-title">Question {idx} / प्रश्न {idx}</td>
        <td class="q-marks">Marks: +1, -0</td>
      </tr>
    </table>
  </div>
  <table class="bilingual-table">
    <tr>
      <td class="col-hindi">
        <div class="q-text">{q['q_hi']}</div>
        {opts_left_html}
      </td>
      <td class="col-eng">
        <div class="q-text">{q['q_en']}</div>
        {opts_right_html}
      </td>
    </tr>
  </table>
  <div class="ans-box">Answer: ({q['ans']})</div>
  <div class="sol-box"><b>Solution:</b> {q['sol']}</div>
  <div class="kp-box">
    <div class="kp-title">Key Points:</div>
    {kp_formatted}
  </div>
</div>
"""

    html_content += "</body></html>"
    
    options = {
        'page-size': 'A4',
        'margin-top': '12mm',
        'margin-bottom': '18mm',
        'margin-left': '12mm',
        'margin-right': '12mm',
        'encoding': "UTF-8",
        'footer-line': '',
        'footer-center': 'Special Education Needs | 9828625119 | Page [page] of [toPage]',
        'footer-font-size': '10',
        'footer-font-name': 'Noto Sans Devanagari',
        'enable-local-file-access': None,
        'quiet': ''
    }
    
    pdfkit.from_string(html_content, output_pdf, options=options)

# ==========================================
# 5. Telegram Handlers
# ==========================================
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Namaste! Apni .docx file upload karein.\n"
        "Main turant clear Hindi fonts aur professional layout ke sath PDF generate karke bhej dunga."
    )

async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc_file = update.message.document
    if not doc_file.file_name.endswith('.docx'):
        await update.message.reply_text("Kripya sirf .docx file upload karein.")
        return

    status_msg = await update.message.reply_text("PDF taiyar ho rahi hai, kripya intezar karein...")
    
    file_id = doc_file.file_id
    new_file = await context.bot.get_file(file_id)
    input_path = f"temp_{doc_file.file_name}"
    output_pdf = input_path.replace(".docx", "_Formatted.pdf")

    await new_file.download_to_drive(input_path)

    try:
        questions = parse_docx(input_path)
        generate_pdf(questions, output_pdf)

        await update.message.reply_document(
            document=open(output_pdf, "rb"),
            filename="Inclusive_Education_Book.pdf",
            caption="Aapki clean PDF taiyar hai!\nSpecial Education Needs | 9828625119"
        )
    except Exception as e:
        await update.message.reply_text(f"Error aaya: {str(e)}")
    finally:
        if os.path.exists(input_path): os.remove(input_path)
        if os.path.exists(output_pdf): os.remove(output_pdf)
        await status_msg.delete()

def main():
    app = ApplicationBuilder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    
    print("Telegram polling started...")
    app.run_polling()

if __name__ == "__main__":
    main()
