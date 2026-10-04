import os
import re
import urllib.parse
from threading import Thread
import requests
from flask import Flask
import telebot
from docx import Document
from weasyprint import HTML

# ===================================================
# 1. Telegram Bot Token & Web Server
# ===================================================
BOT_TOKEN = "8903776742:AAGeYC3UemM-JsuHZ2Af3dmTRAaC7THwcP0"
bot = telebot.TeleBot(BOT_TOKEN)

web_app = Flask(__name__)

@web_app.route('/')
def home():
    return "Special Education Bot is live and running!"

def run_web():
    port = int(os.environ.get("PORT", 10000))
    web_app.run(host="0.0.0.0", port=port)

server_thread = Thread(target=run_web, daemon=True)
server_thread.start()

# ===================================================
# 2. Translation Engine
# ===================================================
trans_cache = {}

def translate_to_english(text):
    if not text or not text.strip():
        return ""
    
    clean_text = text.strip()
    if clean_text in trans_cache:
        return trans_cache[clean_text]

    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/120.0.0.0 Safari/537.36"
    }

    try:
        url = "https://translate.googleapis.com/translate_a/single"
        params = {
            "client": "gtx",
            "sl": "hi",
            "tl": "en",
            "dt": "t",
            "q": clean_text
        }
        res = requests.get(url, params=params, headers=headers, timeout=7)
        if res.status_code == 200:
            translated = "".join([chunk[0] for chunk in res.json()[0] if chunk and chunk[0]])
            if translated.strip():
                trans_cache[clean_text] = translated.strip()
                return translated.strip()
    except Exception:
        pass

    try:
        url2 = f"https://api.mymemory.translated.net/get?q={urllib.parse.quote(clean_text)}&langpair=hi|en"
        res2 = requests.get(url2, headers=headers, timeout=7)
        if res2.status_code == 200:
            tr_mem = res2.json().get("responseData", {}).get("translatedText", "")
            if tr_mem.strip():
                trans_cache[clean_text] = tr_mem.strip()
                return tr_mem.strip()
    except Exception:
        pass

    return clean_text

# ===================================================
# 3. DOCX Parser
# ===================================================
def parse_docx(file_path):
    doc = Document(file_path)
    full_text = "\n".join([p.text.strip() for p in doc.paragraphs if p.text.strip()])
    
    raw_blocks = re.split(r'\n(?=Question:|\d+\s*/\s*प्रश्न)', full_text)
    questions = []
    
    for block in raw_blocks:
        if not re.search(r'(Question:|\d+\s*/\s*प्रश्न)', block):
            continue
            
        q_data = {}
        
        q_match = re.search(r'(?:Question:\s*|\d+\s*/\s*प्रश्न\s*\d*\s*)(.*?)(?=\n\([a-d]\)|\nAnswer:)', block, re.DOTALL)
        q_data['q_hi'] = q_match.group(1).strip() if q_match else ""
        
        opts = re.findall(r'\(([a-d])\)\s*(.*?)(?=\n\([a-d]\)|\nAnswer:|\Z)', block, re.DOTALL)
        q_data['opts'] = {k.lower(): v.strip() for k, v in opts}
        
        ans_match = re.search(r'Answer:\s*([a-d])', block, re.IGNORECASE)
        q_data['ans'] = ans_match.group(1).upper() if ans_match else ""
        
        sol_match = re.search(r'Solution:\s*(.*?)(?=\nKey Points:|\nPositive Marks:|\Z)', block, re.DOTALL)
        q_data['sol'] = sol_match.group(1).strip() if sol_match else ""
        
        kp_match = re.search(r'Key Points:\s*(.*?)(?=\nPositive Marks:|\Z)', block, re.DOTALL)
        q_data['kp'] = kp_match.group(1).strip() if kp_match else ""
        
        questions.append(q_data)
        
    return questions

# ===================================================
# 4. WeasyPrint Engine (Perfect Hindi Text Shaping)
# ===================================================
def generate_pdf(questions, output_pdf):
    html_content = """<!DOCTYPE html>
<html lang="hi">
<head>
<meta charset="utf-8">
<style>
  @page {
    size: A4 portrait;
    margin: 10mm 10mm 16mm 10mm;
    @bottom-center {
      content: "Special Education Needs | Contact: 9828625119 | Page " counter(page) " of " counter(pages);
      font-family: 'Noto Sans Devanagari', 'DejaVu Sans', serif;
      font-size: 10pt;
      font-weight: bold;
      color: #333333;
      border-top: 1px solid #777777;
      width: 100%;
      padding-top: 4px;
    }
  }
  body {
    font-family: 'Noto Sans Devanagari', 'DejaVu Sans', serif;
    font-size: 11pt;
    line-height: 1.4;
    color: #000;
  }
  .q-card {
    page-break-inside: avoid;
    margin-bottom: 18px;
    border-bottom: 1px solid #e0e0e0;
    padding-bottom: 10px;
  }
  .header-table {
    width: 100%;
    border-bottom: 2px solid #b71c1c;
    border-left: 5px solid #b71c1c;
    padding-left: 6px;
    margin-bottom: 6px;
  }
  .q-title {
    color: #b71c1c;
    font-size: 11.5pt;
    font-weight: bold;
    text-align: left;
  }
  .q-marks {
    text-align: right;
    font-size: 10pt;
    color: #555555;
    font-weight: bold;
  }
  .bi-table {
    width: 100%;
    border-collapse: collapse;
    margin-bottom: 6px;
  }
  .col-hi {
    width: 50%;
    vertical-align: top;
    padding-right: 8px;
  }
  .col-en {
    width: 50%;
    vertical-align: top;
    padding-left: 8px;
    border-left: 1px solid #cccccc;
  }
  .q-text {
    font-weight: bold;
    margin-bottom: 5px;
  }
  .opt-item {
    margin-bottom: 3px;
  }
  .ans-box {
    background-color: #f1f8e9;
    border: 1px solid #2e7d32;
    padding: 4px 8px;
    color: #2e7d32;
    font-weight: bold;
    margin-bottom: 4px;
    font-size: 10.5pt;
  }
  .sol-box {
    background-color: #f9fbe7;
    border: 1px solid #cddc39;
    padding: 5px 8px;
    margin-bottom: 4px;
    font-size: 10pt;
  }
  .kp-box {
    background-color: #f5f5f5;
    border-left: 4px solid #1976d2;
    padding: 5px 8px;
    font-size: 9.5pt;
  }
  .kp-title {
    color: #1976d2;
    font-weight: bold;
  }
</style>
</head>
<body>
"""

    for idx, q in enumerate(questions, start=1):
        en_question = translate_to_english(q['q_hi'])
        
        opts_hi_html = ""
        opts_en_html = ""
        for k in ['a', 'b', 'c', 'd']:
            val_hi = q['opts'].get(k, '')
            val_en = translate_to_english(val_hi)
            opts_hi_html += f'<div class="opt-item"><b>({k})</b> {val_hi}</div>'
            opts_en_html += f'<div class="opt-item"><b>({k})</b> {val_en}</div>'

        kp_html = ""
        if q['kp'].strip():
            formatted_kp = q['kp'].replace('\n', '<br/>')
            kp_html = f"""
            <div class="kp-box">
              <span class="kp-title">Key Points:</span><br/>{formatted_kp}
            </div>
            """

        html_content += f"""
        <div class="q-card">
          <table class="header-table">
            <tr>
              <td class="q-title">Question {idx} / प्रश्न {idx}</td>
              <td class="q-marks">Marks: +1, -0</td>
            </tr>
          </table>
          <table class="bi-table">
            <tr>
              <td class="col-hi">
                <div class="q-text">{q['q_hi']}</div>
                {opts_hi_html}
              </td>
              <td class="col-en">
                <div class="q-text">{en_question}</div>
                {opts_en_html}
              </td>
            </tr>
          </table>
          <div class="ans-box">Answer: ({q['ans']})</div>
          <div class="sol-box"><b>Solution:</b> {q['sol']}</div>
          {kp_html}
        </div>
        """

    html_content += "</body></html>"
    HTML(string=html_content).write_pdf(output_pdf)

# ===================================================
# 5. Telegram Handlers
# ===================================================
@bot.message_handler(commands=['start'])
def send_welcome(message):
    bot.reply_to(
        message,
        "नमस्ते! अपनी .docx फ़ाइल अपलोड करें।\n"
        "बॉट हिंदी का सटीक English Translation करके दो-कॉलम (Bilingual) लेआउट में बिना किसी फॉन्ट त्रुटि के PDF बना देगा।"
    )

@bot.message_handler(content_types=['document'])
def handle_docs(message):
    file_name = message.document.file_name
    if not file_name.endswith('.docx'):
        bot.reply_to(message, "कृपया सिर्फ़ .docx फ़ाइल भेजें।")
        return

    status_msg = bot.reply_to(message, "PDF तैयार की जा रही है, कृपया थोड़ा इंतज़ार करें...")
    
    file_info = bot.get_file(message.document.file_id)
    downloaded_file = bot.download_file(file_info.file_path)
    
    input_path = f"temp_{file_name}"
    output_pdf = input_path.replace(".docx", "_Bilingual_Book.pdf")
    
    with open(input_path, 'wb') as f:
        f.write(downloaded_file)
        
    try:
        questions = parse_docx(input_path)
        generate_pdf(questions, output_pdf)
        
        with open(output_pdf, 'rb') as pdf_file:
            bot.send_document(
                message.chat.id,
                pdf_file,
                caption="आपकी शुद्ध Bilingual PDF तैयार है!\nSpecial Education Needs"
            )
    except Exception as e:
        bot.reply_to(message, f"त्रुटि: {str(e)}")
    finally:
        for f in [input_path, output_pdf]:
            if os.path.exists(f):
                os.remove(f)
        try:
            bot.delete_message(message.chat.id, status_msg.message_id)
        except Exception:
            pass

if __name__ == '__main__':
    print("Bot live ho gaya hai...")
    bot.infinity_polling(skip_pending=True)
        
