import os
import re
from threading import Thread
from flask import Flask
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters, ContextTypes
from docx import Document
from weasyprint import HTML

# ==========================================
# 1. Aapka Telegram Bot Token
# ==========================================
BOT_TOKEN = "8903776742:AAGeYC3UemM-JsuHZ2Af3dmTRAaC7THwcP0"

# ==========================================
# 2. Render Web Service Health-Check Server
# ==========================================
web_app = Flask(__name__)

@web_app.route('/')
def home():
    return "Bot is running perfectly on Render!"

def run_web():
    port = int(os.environ.get("PORT", 10000))
    web_app.run(host="0.0.0.0", port=port)

# ==========================================
# 3. DOCX File Reading Function
# ==========================================
def parse_docx(file_path):
    doc = Document(file_path)
    full_text = "\n".join([p.text.strip() for p in doc.paragraphs if p.text.strip()])
    
    raw_blocks = re.split(r'\n(?=Question:)', full_text)
    questions = []
    
    for block in raw_blocks:
        if not block.strip().startswith("Question:"):
            continue
        q_data = {}
        
        # Hindi question extract
        q_match = re.search(r'Question:\s*(.*?)(?=\n\([a-d]\)|\nAnswer:)', block, re.DOTALL)
        q_data['q_hi'] = q_match.group(1).strip() if q_match else ""
        
        # Options extract
        opts = re.findall(r'\(([a-d])\)\s*(.*?)(?=\n\([a-d]\)|\nAnswer:|\Z)', block, re.DOTALL)
        q_data['opts'] = {k.lower(): v.strip() for k, v in opts}
        
        # Answer extract
        ans_match = re.search(r'Answer:\s*([a-d])', block, re.IGNORECASE)
        q_data['ans'] = ans_match.group(1).lower() if ans_match else ""
        
        # Solution extract
        sol_match = re.search(r'Solution:\s*(.*?)(?=\nKey Points:|\nPositive Marks:|\Z)', block, re.DOTALL)
        q_data['sol'] = sol_match.group(1).strip() if sol_match else ""
        
        # Key Points extract
        kp_match = re.search(r'Key Points:\s*(.*?)(?=\nPositive Marks:|\Z)', block, re.DOTALL)
        q_data['kp'] = kp_match.group(1).strip() if kp_match else ""
        
        questions.append(q_data)
    return questions

# ==========================================
# 4. Premium PDF Generator Function
# ==========================================
def generate_pdf(questions, output_pdf):
    html_content = """
    <!DOCTYPE html>
    <html>
    <head>
    <meta charset="utf-8">
    <style>
      @page {
        size: A4 portrait;
        margin: 12mm 12mm 20mm 12mm;
        @bottom-center {
          content: "Special Education Needs | 9828625119 | Page " counter(page);
          font-family: 'Times New Roman', serif;
          font-size: 14pt;
          font-weight: bold;
          color: #222;
          border-top: 1.5px solid #666;
          width: 100%;
          padding-top: 4px;
        }
      }
      body {
        font-family: 'Times New Roman', serif;
        font-size: 18pt;
        line-height: 1.35;
        color: #000;
      }
      .q-card {
        margin-bottom: 26px;
        page-break-inside: avoid;
      }
      .q-header {
        display: flex;
        justify-content: space-between;
        border-bottom: 3px solid #b71c1c;
        border-left: 7px solid #b71c1c;
        padding-left: 10px;
        padding-bottom: 4px;
        margin-bottom: 10px;
        font-weight: bold;
      }
      .q-title { color: #b71c1c; font-size: 18pt; }
      .q-marks { color: #333; font-size: 16pt; }
      .bilingual-table {
        width: 100%;
        display: table;
        margin-bottom: 10px;
      }
      .col {
        display: table-cell;
        width: 50%;
        vertical-align: top;
      }
      .col-left { padding-right: 14px; }
      .col-right {
        padding-left: 14px;
        border-left: 2px solid #bbb;
      }
      .q-text { font-weight: bold; margin-bottom: 8px; }
      .opt { margin-bottom: 6px; }
      .ans-box {
        background: #f1f8e9;
        border: 2px solid #2e7d32;
        padding: 8px 12px;
        color: #2e7d32;
        font-weight: bold;
        margin-bottom: 8px;
      }
      .sol-box {
        background: #f9fbe7;
        border: 1.5px solid #cddc39;
        padding: 10px 14px;
        margin-bottom: 8px;
      }
      .kp-box {
        background: #f5f5f5;
        border-left: 6px solid #1976d2;
        padding: 10px 14px;
      }
      .kp-title { color: #1976d2; font-weight: bold; margin-bottom: 4px; }
    </style>
    </head>
    <body>
    """
    
    for idx, q in enumerate(questions, start=1):
        html_content += f"""
        <div class="q-card">
          <div class="q-header">
            <span class="q-title">Question {idx} / प्रश्न {idx}</span>
            <span class="q-marks">Marks: +1, -0</span>
          </div>
          <div class="bilingual-table">
            <div class="col col-left">
              <div class="q-text">{q['q_hi']}</div>
              <div class="opt"><b>(a)</b> {q['opts'].get('a', '')}</div>
              <div class="opt"><b>(b)</b> {q['opts'].get('b', '')}</div>
              <div class="opt"><b>(c)</b> {q['opts'].get('c', '')}</div>
              <div class="opt"><b>(d)</b> {q['opts'].get('d', '')}</div>
            </div>
            <div class="col col-right">
              <div class="q-text">{q['q_hi']}</div>
              <div class="opt"><b>(a)</b> {q['opts'].get('a', '')}</div>
              <div class="opt"><b>(b)</b> {q['opts'].get('b', '')}</div>
              <div class="opt"><b>(c)</b> {q['opts'].get('c', '')}</div>
              <div class="opt"><b>(d)</b> {q['opts'].get('d', '')}</div>
            </div>
          </div>
          <div class="ans-box">Answer: ({q['ans']})</div>
          <div class="sol-box"><b>Solution:</b> {q['sol']}</div>
          <div class="kp-box">
            <div class="kp-title">Key Points:</div>
            {q['kp'].replace(chr(10), '<br>')}
          </div>
        </div>
        """
        
    html_content += "</body></html>"
    HTML(string=html_content).write_pdf(output_pdf)

# ==========================================
# 5. Telegram Handlers
# ==========================================
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Namaste! Apni .docx file upload karein.\n"
        "Main turant use 18pt font aur photo jaise 2-column layout me print-ready PDF bana kar bhej dunga."
    )

async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc_file = update.message.document
    if not doc_file.file_name.endswith('.docx'):
        await update.message.reply_text("Kripya sirf .docx file upload karein.")
        return

    status_msg = await update.message.reply_text("File process ho rahi hai, kripya thoda intezar karein...")
    
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
            filename=f"Inclusive_Education_Book.pdf",
            caption="Aapki PDF book taiyar hai!\nSpecial Education Needs | 9828625119"
        )
    except Exception as e:
        await update.message.reply_text(f"Error aaya: {str(e)}")
    finally:
        if os.path.exists(input_path): os.remove(input_path)
        if os.path.exists(output_pdf): os.remove(output_pdf)
        await status_msg.delete()

# ==========================================
# 6. Main Runner
# ==========================================
def main():
    # Render web service ke port scanner ko pass karne ke liye flask start karein
    server_thread = Thread(target=run_web)
    server_thread.daemon = True
    server_thread.start()

    # Telegram Bot Start
    app = ApplicationBuilder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    
    print("Bot Render par successfully start ho gaya hai...")
    app.run_polling()

if __name__ == "__main__":
    main()
    
