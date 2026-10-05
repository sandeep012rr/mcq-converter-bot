import os
import io
import re
import sys
import telebot
from flask import Flask, request, jsonify
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from weasyprint import HTML

# ==========================================
# 1. कॉन्फ़िगरेशन और एनवायरनमेंट वेरिएबल्स
# ==========================================
BOT_TOKEN = os.getenv("BOT_TOKEN")
WEBHOOK_URL = os.getenv("WEBHOOK_URL", "https://mcq-converter-bot.onrender.com")

if not BOT_TOKEN:
    sys.exit("❌ Error: कृपया Render Environment Variables में 'BOT_TOKEN' सेट करें!")

# Flask और TeleBot इनिशियलाइज़ेशन
# threaded=False वेबहुक मोड में Gunicorn के साथ सबसे सुरक्षित रहता है
bot = telebot.TeleBot(BOT_TOKEN, threaded=False)
app = Flask(__name__)


# ==========================================
# 2. सहायक फ़ंक्शंस (MCQ DOCX और PDF जनरेटर)
# ==========================================
def create_mcq_docx(raw_text):
    """दिए गए टेक्स्ट से फॉर्मेटेड Word (.docx) फाइल बनाता है (RAM में)"""
    doc = Document()
    
    # डॉक्यूमेंट हेडिंग
    heading = doc.add_heading("MCQ Question Bank", level=1)
    heading.alignment = 1  # Center alignment
    
    doc.add_paragraph("Generated automatically via MCQ Converter Bot\n" + "-" * 50)
    
    # लाइन-बाय-लाइन जोड़ना
    for line in raw_text.split('\n'):
        line_clean = line.strip()
        if not line_clean:
            continue
        
        # प्रश्न या विकल्प को हाईलाइट करना
        p = doc.add_paragraph()
        if re.match(r'^(Q\d+|प्रश्न|\d+[\.\)])', line_clean, re.IGNORECASE):
            run = p.add_run(line_clean)
            run.bold = True
            run.font.size = Pt(11.5)
        elif re.match(r'^(Ans|Answer|उत्तर|Correct):?', line_clean, re.IGNORECASE):
            run = p.add_run(line_clean)
            run.bold = True
            run.font.color.rgb = RGBColor(0, 128, 0)  # हरा रंग
        else:
            p.add_run(line_clean)
            
    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io


def create_mcq_pdf(raw_text):
    """WeasyPrint का उपयोग करके स्वच्छ PDF बनाता है (RAM में)"""
    html_lines = []
    for line in raw_text.split('\n'):
        line_clean = line.strip()
        if not line_clean:
            continue
        if re.match(r'^(Q\d+|प्रश्न|\d+[\.\)])', line_clean, re.IGNORECASE):
            html_lines.append(f'<div class="question">{line_clean}</div>')
        elif re.match(r'^(Ans|Answer|उत्तर|Correct):?', line_clean, re.IGNORECASE):
            html_lines.append(f'<div class="answer">{line_clean}</div>')
        else:
            html_lines.append(f'<div class="option">{line_clean}</div>')
            
    content_html = "\n".join(html_lines)
    
    html_template = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <style>
            @page {{ size: A4; margin: 20mm; }}
            body {{ font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; font-size: 13px; line-height: 1.6; color: #333; }}
            h1 {{ text-align: center; color: #1a365d; border-bottom: 2px solid #e2e8f0; padding-bottom: 8px; }}
            .question {{ font-weight: bold; margin-top: 15px; margin-bottom: 5px; color: #0f172a; font-size: 14px; }}
            .option {{ margin-left: 15px; margin-bottom: 3px; color: #334155; }}
            .answer {{ font-weight: bold; color: #15803d; margin-left: 15px; margin-top: 4px; margin-bottom: 8px; }}
        </style>
    </head>
    <body>
        <h1>MCQ Question Sheet</h1>
        {content_html}
    </body>
    </html>
    """
    
    pdf_io = io.BytesIO()
    HTML(string=html_template).write_pdf(pdf_io)
    pdf_io.seek(0)
    return pdf_io


# ==========================================
# 3. टेलीग्राम बॉट हैंडलर्स
# ==========================================
@bot.message_handler(commands=['start'])
def send_welcome(message):
    msg = (
        "👋 *MCQ कनवर्टर बॉट में आपका स्वागत है!*\n\n"
        "मुझे अपने प्रश्न (MCQs) भेजें। आप:\n"
        "1. सीधे यहाँ टेक्स्ट पेस्ट कर सकते हैं।\n"
        "2. या कोई टेक्स्ट फ़ाइल (.txt) भेज सकते हैं।\n\n"
        "मैं आपको तुरंत उसकी फॉर्मेटेड **Word (.docx)** और **PDF** बनाकर भेजूंगा।"
    )
    bot.reply_to(message, msg, parse_mode="Markdown")


@bot.message_handler(commands=['ping'])
def send_ping(message):
    bot.reply_to(message, "🏓 Pong! बॉट Webhook पर सक्रिय है।")


# टेक्स्ट प्राप्त होने पर DOCX और PDF बनाकर भेजना
@bot.message_handler(content_types=['text'])
def handle_text_mcqs(message):
    user_text = message.text.strip()
    
    if len(user_text) < 10:
        bot.reply_to(message, "⚠️ कृपया कम से कम एक पूरा प्रश्न या विकल्प भेजें।")
        return

    status_msg = bot.reply_to(message, "⏳ आपकी फाइल्स तैयार हो रही हैं...")
    
    try:
        # DOCX बनाएं और भेजें
        docx_file = create_mcq_docx(user_text)
        docx_file.name = "MCQ_Document.docx"
        bot.send_document(message.chat.id, docx_file, caption="📄 आपकी Word (.docx) फ़ाइल")
        
        # PDF बनाएं और भेजें
        pdf_file = create_mcq_pdf(user_text)
        pdf_file.name = "MCQ_Document.pdf"
        bot.send_document(message.chat.id, pdf_file, caption="📕 आपकी PDF फ़ाइल")
        
        bot.delete_message(message.chat.id, status_msg.message_id)
    except Exception as e:
        bot.edit_message_text(f"❌ प्रोसेस करने में त्रुटि आई: {str(e)}", message.chat.id, status_msg.message_id)


# TXT फ़ाइल प्राप्त होने पर
@bot.message_handler(content_types=['document'])
def handle_file_upload(message):
    try:
        file_name = message.document.file_name or ""
        
        if not file_name.endswith('.txt'):
            bot.reply_to(message, "⚠️ कृपया केवल `.txt` फाइल या सीधे टेक्स्ट मैसेज भेजें।")
            return
            
        status_msg = bot.reply_to(message, f"⏳ `{file_name}` डाउनलोड और कन्वर्ट हो रही है...", parse_mode="Markdown")
        
        file_info = bot.get_file(message.document.file_id)
        downloaded = bot.download_file(file_info.file_path)
        text_content = downloaded.decode('utf-8', errors='ignore')
        
        # DOCX और PDF बनाएं
        docx_file = create_mcq_docx(text_content)
        docx_file.name = f"{os.path.splitext(file_name)[0]}.docx"
        bot.send_document(message.chat.id, docx_file, caption="📄 आपकी Word फ़ाइल")
        
        pdf_file = create_mcq_pdf(text_content)
        pdf_file.name = f"{os.path.splitext(file_name)[0]}.pdf"
        bot.send_document(message.chat.id, pdf_file, caption="📕 आपकी PDF फ़ाइल")
        
        bot.delete_message(message.chat.id, status_msg.message_id)
    except Exception as e:
        bot.reply_to(message, f"❌ एरर: {str(e)}")


# ==========================================
# 4. Flask Webhook और एंडपॉइंट्स
# ==========================================
@app.route(f"/{BOT_TOKEN}", methods=['POST'])
def process_webhook():
    """Telegram इस रूट पर नए अपडेट्स पोस्ट करता है"""
    if request.headers.get('content-type') == 'application/json':
        json_data = request.get_data().decode('utf-8')
        update = telebot.types.Update.de_json(json_data)
        bot.process_new_updates([update])
        return 'OK', 200
    return 'Forbidden', 403


@app.route('/', methods=['GET', 'HEAD'])
@app.route('/ping', methods=['GET'])
def health_check():
    """Render हेल्थ चेक और UptimeRobot पिंग के लिए"""
    return jsonify({
        "status": "online",
        "service": "mcq-converter-bot",
        "mode": "webhook"
    }), 200


# ==========================================
# 5. ऑटोमैटिक Webhook रजिस्ट्रेशन (सर्वर स्टार्ट)
# ==========================================
def init_webhook():
    full_webhook_url = f"{WEBHOOK_URL.rstrip('/')}/{BOT_TOKEN}"
    try:
        current_info = bot.get_webhook_info()
        if current_info.url != full_webhook_url:
            bot.remove_webhook()
            bot.set_webhook(url=full_webhook_url)
            print(f"[+] Telegram Webhook successfully configured to: {full_webhook_url}")
        else:
            print("[+] Telegram Webhook already up-to-date.")
    except Exception as ex:
        print(f"[-] Webhook setup failed: {ex}")

# जब Gunicorn या डायरेक्ट रनर इस फाइल को लोड करेगा, वेबहुक ऑटो-सेट हो जाएगा
init_webhook()

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 10000))
    app.run(host="0.0.0.0", port=port)
