import os
import re
from threading import Thread
from flask import Flask
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters, ContextTypes
from docx import Document
from xhtml2pdf import pisa

BOT_TOKEN = "8903776742:AAGeYC3UemM-JsuHZ2Af3dmTRAaC7THwcP0"

# Render Web Service ke port check ke liye Flask app
web_app = Flask(__name__)

@web_app.route('/')
def home():
    return "Bot is running online on Render!"

def run_web():
    port = int(os.environ.get("PORT", 10000))
    web_app.run(host="0.0.0.0", port=port)

def parse_docx(file_path):
    doc = Document(file_path)
    full_text = "\n".join([p.text.strip() for p in doc.paragraphs if p.text.strip()])
    
    raw_blocks = re.split(r'\n(?=Question:)', full_text)
    questions = []
    
    for block in raw_blocks:
        if not block.strip().startswith("Question:"):
            continue
        q_data = {}
        
        q_match = re.search(r'Question:\s*(.*?)(?=\n\([a-d]\)|\nAnswer:)', block, re.DOTALL)
        q_data['q_hi'] = q_match.group(1).strip() if q_match else ""
        
        opts = re.findall(r'\(([a-d])\)\s*(.*?)(?=\n\([a-d]\)|\nAnswer:|\Z)', block, re.DOTALL)
        q_data['opts'] = {k.lower(): v.strip() for k, v in opts}
        
        ans_match = re.search(r'Answer:\s*([a-d])', block, re.IGNORECASE)
        q_data['ans'] = ans_match.group(1).lower() if ans_match else ""
        
        sol_match = re.search(r'Solution:\s*(.*?)(?=\nKey Points:|\nPositive Marks:|\Z)', block, re.DOTALL)
        q_data['sol'] = sol_match.group(1).strip() if sol_match else ""
        
        kp_match = re.search(r'Key Points:\s*(.*?)(?=\nPositive Marks:|\Z)', block, re.DOTALL)
        q_data['kp'] = kp_match.group(1).strip() if kp_match else ""
        
        questions.append(q_data)
    return questions

def generate_pdf(questions, output_pdf):
    html_content = """
    <!DOCTYPE html>
    <html>
    <head>
    <meta charset="utf-8">
    <style>
      @page {
        size: a4 portrait;
        margin: 1.2cm 1.2cm 1.8cm 1.2cm;
        @bottom-center {
          content: "Special Education Needs | 9828625119";
          font-family: 'Times New Roman', serif;
          font-size: 13pt;
          color: #333333;
        }
      }
      body {
        font-family: 'Times New Roman', serif;
        font-size: 15pt;
        line-height: 1.35;
        color: #000000;
      }
      .q-card {
        margin-bottom: 22px;
        page-break-inside: avoid;
      }
      .header-table {
        width: 100%;
        border-bottom: 2px solid #b71c1c;
        margin-bottom: 8px;
        padding-bottom: 4px;
      }
      .q-title {
        color: #b71c1c;
        font-size: 17pt;
        font-weight: bold;
      }
      .q-marks {
        text-align: right;
        font-size: 14pt;
        color: #555555;
      }
      .bilingual-table {
        width: 100%;
        margin-bottom: 10px;
      }
      .col-left {
        width: 50%;
        vertical-align: top;
        padding-right: 10px;
      }
      .col-right {
        width: 50%;
        vertical-align: top;
        padding-left: 10px;
        border-left: 1px solid #cccccc;
      }
      .q-text {
        font-weight: bold;
        margin-bottom: 6px;
      }
      .opt {
        margin-bottom: 4px;
      }
      .ans-box {
        background-color: #f1f8e9;
        border: 1px solid #2e7d32;
        padding: 6px 10px;
        color: #2e7d32;
        font-weight: bold;
        margin-bottom: 6px;
      }
      .sol-box {
        background-color: #f9fbe7;
        border: 1px solid #cddc39;
        padding: 8px 10px;
        margin-bottom: 6px;
      }
      .kp-box {
        background-color: #f5f5f5;
        border-left: 4px solid #1976d2;
        padding: 8px 10px;
      }
      .kp-title {
        color: #1976d2;
        font-weight: bold;
        margin-bottom: 4px;
      }
    </style>
    </head>
    <body>
    """

    for idx, q in enumerate(questions, start=1):
        html_content += f"""
        <div class="q-card">
          <table class="header-table">
            <tr>
              <td class="q-title">Question {idx} / प्रश्न {idx}</td>
              <td class="q-marks">Marks: +1, -0</td>
            </tr>
          </table>
          <table class="bilingual-table">
            <tr>
              <td class="col-left">
                <div class="q-text">{q['q_hi']}</div>
                <div class="opt"><b>(a)</b> {q['opts'].get('a', '')}</div>
                <div class="opt"><b>(b)</b> {q['opts'].get('b', '')}</div>
                <div class="opt"><b>(c)</b> {q['opts'].get('c', '')}</div>
                <div class="opt"><b>(d)</b> {q['opts'].get('d', '')}</div>
              </td>
              <td class="col-right">
                <div class="q-text">{q['q_hi']}</div>
                <div class="opt"><b>(a)</b> {q['opts'].get('a', '')}</div>
                <div class="opt"><b>(b)</b> {q['opts'].get('b', '')}</div>
                <div class="opt"><b>(c)</b> {q['opts'].get('c', '')}</div>
                <div class="opt"><b>(d)</b> {q['opts'].get('d', '')}</div>
              </td>
            </tr>
          </table>
          <div class="ans-box">Answer: ({q['ans']})</div>
          <div class="sol-box"><b>Solution:</b> {q['sol']}</div>
          <div class="kp-box">
            <div class="kp-title">Key Points:</div>
            {q['kp'].replace(chr(10), '<br>')}
          </div>
        </div>
        """

    html_content += "</body></html>"
    
    with open(output_pdf, "wb") as pdf_file:
        pisa_status = pisa.CreatePDF(html_content, dest=pdf_file, encoding='utf-8')
        if pisa_status.err:
            raise Exception("PDF conversion failed")

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Namaste! Apni .docx file upload karein.\n"
        "Main turant 2-column bilingual layout me PDF bana kar bhej dunga."
    )

async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc_file = update.message.document
    if not doc_file.file_name.endswith('.docx'):
        await update.message.reply_text("Kripya sirf .docx file upload karein.")
        return

    status_msg = await update.message.reply_text("File process ho rahi hai, kripya intezar karein...")
    
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
            caption="Aapki PDF book taiyar hai!\nSpecial Education Needs | 9828625119"
        )
    except Exception as e:
        await update.message.reply_text(f"Error aaya: {str(e)}")
    finally:
        if os.path.exists(input_path): os.remove(input_path)
        if os.path.exists(output_pdf): os.remove(output_pdf)
        await status_msg.delete()

def main():
    server_thread = Thread(target=run_web)
    server_thread.daemon = True
    server_thread.start()

    app = ApplicationBuilder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    
    print("Bot live hai...")
    app.run_polling()

if __name__ == "__main__":
    main()
