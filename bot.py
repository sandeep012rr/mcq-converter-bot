import os
import re
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters, ContextTypes
from docx import Document
from weasyprint import HTML

# Aapka Telegram Bot Token
BOT_TOKEN = "8903776742:AAGeYC3UemM-JsuHZ2Af3dmTRAaC7THwcP0"

def parse_docx(file_path):
    doc = Document(file_path)
    full_text = "\n".join([p.text.strip() for p in doc.paragraphs if p.text.strip()])
    
    # Question blocks ko alag karna
    raw_blocks = re.split(r'\n(?=Question:)', full_text)
    questions = []
    
    for block in raw_blocks:
        if not block.strip().startswith("Question:"):
            continue
        q_data = {}
        
        # Question text extract karna
        q_match = re.search(r'Question:\s*(.*?)(?=\n\([a-d]\)|\nAnswer:)', block, re.DOTALL)
        q_data['q_hi'] = q_match.group(1).strip() if q_match else ""
        
        # Options extract karna
        opts = re.findall(r'\(([a-d])\)\s*(.*?)(?=\n\([a-d]\)|\nAnswer:|\Z)', block, re.DOTALL)
        q_data['opts'] = {k.lower(): v.strip() for k, v in opts}
        
        # Answer extract karna
        ans_match = re.search(r'Answer:\s*([a-d])', block, re.IGNORECASE)
        q_data['ans'] = ans_match.group(1).lower() if ans_match else ""
        
        # Solution extract karna
        sol_match = re.search(r'Solution:\s*(.*?)(?=\nKey Points:|\nPositive Marks:|\Z)', block, re.DOTALL)
        q_data['sol'] = sol_match.group(1).strip() if sol_match else ""
        
        # Key Points extract karna
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
        size: A4 portrait;
        margin: 12mm 12mm 18mm 12mm;
        @bottom-center {
          content: "Special Education Needs | 9828625119 | Page " counter(page);
          font-family: 'Times New Roman', serif;
          font-size: 13pt;
          color: #333;
          border-top: 1px solid #777;
          width: 100%;
        }
      }
      body {
        font-family: 'Times New Roman', serif;
        font-size: 15pt;
        line-height: 1.35;
        color: #000;
      }
      .q-card {
        margin-bottom: 24px;
        page-break-inside: avoid;
      }
      .q-header {
        display: flex;
        justify-content: space-between;
        border-bottom: 2.5px solid #b71c1c;
        border-left: 6px solid #b71c1c;
        padding-left: 8px;
        padding-bottom: 3px;
        margin-bottom: 8px;
        font-weight: bold;
      }
      .q-title { color: #b71c1c; font-size: 16pt; }
      .q-marks { color: #444; font-size: 13pt; }
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
      .col-left { padding-right: 12px; }
      .col-right {
        padding-left: 12px;
        border-left: 1.5px solid #bbb;
      }
      .q-text { font-weight: bold; margin-bottom: 6px; }
      .opt { margin-bottom: 4px; }
      .ans-box {
        background: #f1f8e9;
        border: 2px solid #2e7d32;
        padding: 6px 12px;
        color: #2e7d32;
        font-weight: bold;
        margin-bottom: 6px;
      }
      .sol-box {
        background: #f9fbe7;
        border: 1.5px solid #cddc39;
        padding: 8px 12px;
        margin-bottom: 6px;
      }
      .kp-box {
        background: #f5f5f5;
        border-left: 5px solid #1976d2;
        padding: 8px 12px;
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

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Namaste! Apni .docx file bhejein, main use turant print-ready PDF book format me convert kar dunga.")

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
            caption="Aapki PDF book taiyar hai!"
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
    print("Bot live hai aur kaam kar raha hai...")
    app.run_polling()

if __name__ == "__main__":
    main()
