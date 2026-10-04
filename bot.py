import os
import re
import urllib.parse
import urllib.request
from threading import Thread
import requests
from flask import Flask
import telebot
from docx import Document

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, KeepTogether
)
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# ===================================================
# 1. Bot Token & Flask Web Service Setup
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
# 2. Hindi Fonts Auto-Setup (No Broken Matras)
# ===================================================
FONT_REGULAR = "NotoSansDevanagari-Regular.ttf"
FONT_BOLD = "NotoSansDevanagari-Bold.ttf"

def load_hindi_fonts():
    if not os.path.exists(FONT_REGULAR):
        url = "https://raw.githubusercontent.com/googlefonts/noto-fonts/main/hinted/ttf/NotoSansDevanagari/NotoSansDevanagari-Regular.ttf"
        urllib.request.urlretrieve(url, FONT_REGULAR)
    pdfmetrics.registerFont(TTFont('DevaFont', FONT_REGULAR))

    if not os.path.exists(FONT_BOLD):
        url_bold = "https://raw.githubusercontent.com/googlefonts/noto-fonts/main/hinted/ttf/NotoSansDevanagari/NotoSansDevanagari-Bold.ttf"
        urllib.request.urlretrieve(url_bold, FONT_BOLD)
    pdfmetrics.registerFont(TTFont('DevaFont-Bold', FONT_BOLD))

try:
    load_hindi_fonts()
    print("Devanagari fonts loaded successfully!")
except Exception as e:
    print(f"Font loading error: {e}")

# ===================================================
# 3. Robust Translation Engine (Hindi -> English)
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

    # Primary: Google Translate Direct Web Endpoint
    try:
        url = "https://translate.googleapis.com/translate_a/single"
        params = {
            "client": "gtx",
            "sl": "hi",
            "tl": "en",
            "dt": "t",
            "q": clean_text
        }
        res = requests.get(url, params=params, headers=headers, timeout=6)
        if res.status_code == 200:
            translated = "".join([chunk[0] for chunk in res.json()[0] if chunk and chunk[0]])
            if translated.strip():
                trans_cache[clean_text] = translated.strip()
                return translated.strip()
    except Exception:
        pass

    # Secondary: MyMemory Free Translation API
    try:
        url2 = f"https://api.mymemory.translated.net/get?q={urllib.parse.quote(clean_text)}&langpair=hi|en"
        res2 = requests.get(url2, headers=headers, timeout=6)
        if res2.status_code == 200:
            tr_mem = res2.json().get("responseData", {}).get("translatedText", "")
            if tr_mem.strip():
                trans_cache[clean_text] = tr_mem.strip()
                return tr_mem.strip()
    except Exception:
        pass

    return clean_text

# ===================================================
# 4. DOCX Parser
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
# 5. Canvas for Page Counter & Footer Branding
# ===================================================
class NumberedCanvas(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_footer(num_pages)
            super().showPage()
        super().save()

    def draw_footer(self, page_count):
        self.saveState()
        self.setFont("DevaFont", 10.5)
        self.setStrokeColor(colors.HexColor("#777777"))
        self.setLineWidth(0.8)
        self.line(36, 42, 595 - 36, 42)
        
        footer_text = f"Special Education Needs | 9828625119 | Page {self._pageNumber} of {page_count}"
        self.setFillColor(colors.HexColor("#222222"))
        self.drawCentredString(595 / 2.0, 26, footer_text)
        self.restoreState()

# ===================================================
# 6. Bilingual PDF Generator (Left Hindi, Right English)
# ===================================================
def generate_pdf(questions, output_pdf):
    doc = SimpleDocTemplate(
        output_pdf,
        pagesize=A4,
        leftMargin=32,
        rightMargin=32,
        topMargin=32,
        bottomMargin=54
    )
    
    styles = getSampleStyleSheet()
    
    title_left = ParagraphStyle(
        'HeaderTitle',
        parent=styles['Normal'],
        fontName='DevaFont-Bold',
        fontSize=12,
        leading=15,
        textColor=colors.HexColor('#B71C1C')
    )
    title_right = ParagraphStyle(
        'HeaderMarks',
        parent=styles['Normal'],
        fontName='DevaFont',
        fontSize=10.5,
        leading=15,
        alignment=2,
        textColor=colors.HexColor('#444444')
    )
    q_hi_style = ParagraphStyle(
        'QHi',
        parent=styles['Normal'],
        fontName='DevaFont-Bold',
        fontSize=10,
        leading=14,
        textColor=colors.black
    )
    q_en_style = ParagraphStyle(
        'QEn',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=10,
        leading=14,
        textColor=colors.black
    )
    opt_hi_style = ParagraphStyle(
        'OptHi',
        parent=styles['Normal'],
        fontName='DevaFont',
        fontSize=9.5,
        leading=13.5,
        textColor=colors.black
    )
    opt_en_style = ParagraphStyle(
        'OptEn',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=9.5,
        leading=13.5,
        textColor=colors.black
    )
    ans_style = ParagraphStyle(
        'Ans',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=11,
        leading=14,
        textColor=colors.HexColor('#2E7D32')
    )
    sol_style = ParagraphStyle(
        'Sol',
        parent=styles['Normal'],
        fontName='DevaFont',
        fontSize=9.5,
        leading=13.5,
        textColor=colors.black
    )
    kp_style = ParagraphStyle(
        'KP',
        parent=styles['Normal'],
        fontName='DevaFont',
        fontSize=9,
        leading=13,
        textColor=colors.HexColor('#222222')
    )
    
    story = []
    content_width = 595 - 64  # 531 pt
    col_width = (content_width - 8) / 2.0  # ~261 pt
    
    for idx, q in enumerate(questions, start=1):
        q_elements = []
        
        # 1. Header (Red line + Marks)
        hdr_table = Table([[
            Paragraph(f"Question {idx} / प्रश्न {idx}", title_left),
            Paragraph("Marks: +1, -0", title_right)
        ]], colWidths=[col_width, col_width])
        hdr_table.setStyle(TableStyle([
            ('LINEBEFORE', (0, 0), (0, -1), 4, colors.HexColor('#B71C1C')),
            ('LINEBELOW', (0, 0), (-1, -1), 1.5, colors.HexColor('#B71C1C')),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
            ('TOPPADDING', (0, 0), (-1, -1), 2),
            ('LEFTPADDING', (0, 0), (0, -1), 6),
            ('RIGHTPADDING', (-1, 0), (-1, -1), 2),
        ]))
        q_elements.append(hdr_table)
        q_elements.append(Spacer(1, 5))
        
        # 2. English Translation Process
        en_question = translate_to_english(q['q_hi'])
        
        # Left Flowables (Hindi)
        left_items = [Paragraph(q['q_hi'], q_hi_style), Spacer(1, 4)]
        for k in ['a', 'b', 'c', 'd']:
            val_hi = q['opts'].get(k, '')
            left_items.append(Paragraph(f"<b>({k})</b> {val_hi}", opt_hi_style))
            left_items.append(Spacer(1, 2))
            
        # Right Flowables (English)
        right_items = [Paragraph(en_question, q_en_style), Spacer(1, 4)]
        for k in ['a', 'b', 'c', 'd']:
            val_hi = q['opts'].get(k, '')
            val_en = translate_to_english(val_hi)
            right_items.append(Paragraph(f"<b>({k})</b> {val_en}", opt_en_style))
            right_items.append(Spacer(1, 2))
            
        # 3. Two-Column Bilingual Table
        bi_table = Table([[left_items, right_items]], colWidths=[col_width, col_width])
        bi_table.setStyle(TableStyle([
            ('LINEBEFORE', (1, 0), (1, -1), 1, colors.HexColor('#D0D0D0')),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('RIGHTPADDING', (0, 0), (0, -1), 8),
            ('LEFTPADDING', (1, 0), (1, -1), 8),
            ('TOPPADDING', (0, 0), (-1, -1), 2),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ]))
        q_elements.append(bi_table)
        q_elements.append(Spacer(1, 5))
        
        # 4. Answer Box (Green)
        ans_table = Table([[Paragraph(f"Answer: ({q['ans']})", ans_style)]], colWidths=[content_width])
        ans_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#F1F8E9')),
            ('BOX', (0, 0), (-1, -1), 1, colors.HexColor('#2E7D32')),
            ('TOPPADDING', (0, 0), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
            ('LEFTPADDING', (0, 0), (-1, -1), 8),
        ]))
        q_elements.append(ans_table)
        q_elements.append(Spacer(1, 4))
        
        # 5. Solution Box
        sol_table = Table([[Paragraph(f"<b>Solution:</b> {q['sol']}", sol_style)]], colWidths=[content_width])
        sol_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#F9FBE7')),
            ('BOX', (0, 0), (-1, -1), 1, colors.HexColor('#CDDC39')),
            ('TOPPADDING', (0, 0), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
            ('LEFTPADDING', (0, 0), (-1, -1), 8),
        ]))
        q_elements.append(sol_table)
        q_elements.append(Spacer(1, 4))
        
        # 6. Key Points Box
        if q['kp'].strip():
            kp_text = q['kp'].replace('\n', '<br/>')
            kp_table = Table([[Paragraph(f"<font color='#1976D2'><b>Key Points:</b></font><br/>{kp_text}", kp_style)]], colWidths=[content_width])
            kp_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#F5F5F5')),
                ('LINEBEFORE', (0, 0), (0, -1), 3.5, colors.HexColor('#1976D2')),
                ('TOPPADDING', (0, 0), (-1, -1), 4),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
                ('LEFTPADDING', (0, 0), (-1, -1), 8),
            ]))
            q_elements.append(kp_table)
            q_elements.append(Spacer(1, 12))
            
        story.append(KeepTogether(q_elements))
        
    doc.build(story, canvasmaker=NumberedCanvas)

# ===================================================
# 7. Telegram Handlers (telebot)
# ===================================================
@bot.message_handler(commands=['start'])
def send_welcome(message):
    bot.reply_to(
        message,
        "नमस्ते! अपनी .docx फ़ाइल अपलोड करें।\n"
        "बॉट हिंदी का सटीक English Translation करके दो-कॉलम (Bilingual) लेआउट में PDF बुक बना देगा।"
    )

@bot.message_handler(content_types=['document'])
def handle_docs(message):
    file_name = message.document.file_name
    if not file_name.endswith('.docx'):
        bot.reply_to(message, "कृपया सिर्फ़ .docx फ़ाइल भेजें।")
        return

    status_msg = bot.reply_to(message, "Bilingual Translation और PDF निर्माण जारी है, कृपया 1-2 मिनट प्रतीक्षा करें...")
    
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
                caption="आपकी Bilingual PDF तैयार है!\nSpecial Education Needs | 9828625119"
            )
    except Exception as e:
        bot.reply_to(message, f"त्रुटि आई: {str(e)}")
    finally:
        for f in [input_path, output_pdf]:
            if os.path.exists(f):
                os.remove(f)
        try:
            bot.delete_message(message.chat.id, status_msg.message_id)
        except Exception:
            pass

if __name__ == '__main__':
    print("Telegram Bot polling started...")
    bot.infinity_polling(skip_pending=True)
    
