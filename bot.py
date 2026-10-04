import os
import re
from threading import Thread
from flask import Flask
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters, ContextTypes
from docx import Document

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, KeepTogether
)
from reportlab.pdfgen import canvas

# ==========================================
# 1. Telegram Bot Token
# ==========================================
BOT_TOKEN = "8903776742:AAGeYC3UemM-JsuHZ2Af3dmTRAaC7THwcP0"

# ==========================================
# 2. Render Web Service Port Listener (Flask)
# ==========================================
web_app = Flask(__name__)

@web_app.route('/')
def home():
    return "Bot is live and running perfectly on Render!"

def run_web():
    port = int(os.environ.get("PORT", 10000))
    print(f"Starting web server on port {port}...")
    web_app.run(host="0.0.0.0", port=port)

# Background me Web server turant chalu karein
server_thread = Thread(target=run_web, daemon=True)
server_thread.start()

# ==========================================
# 3. Canvas for Auto Footer with Page Count
# ==========================================
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
        self.setFont("Times-Roman", 11)
        self.setStrokeColor(colors.HexColor("#777777"))
        self.setLineWidth(0.8)
        self.line(36, 42, 595 - 36, 42)
        
        footer_text = f"Special Education Needs | 9828625119 | Page {self._pageNumber} of {page_count}"
        self.setFillColor(colors.HexColor("#222222"))
        self.drawCentredString(595 / 2.0, 26, footer_text)
        self.restoreState()

# ==========================================
# 4. DOCX Parser
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

# ==========================================
# 5. Clean PDF Generator (ReportLab)
# ==========================================
def generate_pdf(questions, output_pdf):
    doc = SimpleDocTemplate(
        output_pdf,
        pagesize=A4,
        leftMargin=36,
        rightMargin=36,
        topMargin=36,
        bottomMargin=54
    )
    
    styles = getSampleStyleSheet()
    
    title_left = ParagraphStyle(
        'HeaderTitle',
        parent=styles['Normal'],
        fontName='Times-Bold',
        fontSize=13,
        leading=16,
        textColor=colors.HexColor('#B71C1C')
    )
    title_right = ParagraphStyle(
        'HeaderMarks',
        parent=styles['Normal'],
        fontName='Times-Roman',
        fontSize=11,
        leading=16,
        alignment=2,
        textColor=colors.HexColor('#444444')
    )
    q_style = ParagraphStyle(
        'QuestionText',
        parent=styles['Normal'],
        fontName='Times-Bold',
        fontSize=11,
        leading=15,
        textColor=colors.black
    )
    opt_style = ParagraphStyle(
        'OptionText',
        parent=styles['Normal'],
        fontName='Times-Roman',
        fontSize=10.5,
        leading=14,
        textColor=colors.black
    )
    ans_style = ParagraphStyle(
        'AnswerText',
        parent=styles['Normal'],
        fontName='Times-Bold',
        fontSize=11.5,
        leading=15,
        textColor=colors.HexColor('#2E7D32')
    )
    sol_style = ParagraphStyle(
        'SolutionText',
        parent=styles['Normal'],
        fontName='Times-Roman',
        fontSize=10.5,
        leading=14.5,
        textColor=colors.black
    )
    kp_style = ParagraphStyle(
        'KeyPointsText',
        parent=styles['Normal'],
        fontName='Times-Roman',
        fontSize=10,
        leading=14,
        textColor=colors.HexColor('#222222')
    )
    
    story = []
    content_width = 595 - 72
    col_width = content_width / 2.0
    
    for idx, q in enumerate(questions, start=1):
        q_elements = []
        
        # 1. Header Bar
        hdr_data = [[
            Paragraph(f"Question {idx} / प्रश्न {idx}", title_left),
            Paragraph("Marks: +1, -0", title_right)
        ]]
        hdr_table = Table(hdr_data, colWidths=[col_width, col_width])
        hdr_table.setStyle(TableStyle([
            ('LINEBEFORE', (0, 0), (0, -1), 4, colors.HexColor('#B71C1C')),
            ('LINEBELOW', (0, 0), (-1, -1), 1.5, colors.HexColor('#B71C1C')),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
            ('TOPPADDING', (0, 0), (-1, -1), 2),
            ('LEFTPADDING', (0, 0), (0, -1), 6),
            ('RIGHTPADDING', (-1, 0), (-1, -1), 2),
        ]))
        q_elements.append(hdr_table)
        q_elements.append(Spacer(1, 6))
        
        # 2. Bilingual Two Columns
        left_flowables = [Paragraph(q['q_hi'], q_style), Spacer(1, 4)]
        for opt_key in ['a', 'b', 'c', 'd']:
            text = q['opts'].get(opt_key, '')
            left_flowables.append(Paragraph(f"<b>({opt_key})</b> {text}", opt_style))
            left_flowables.append(Spacer(1, 2))
            
        right_flowables = [Paragraph(q['q_hi'], q_style), Spacer(1, 4)]
        for opt_key in ['a', 'b', 'c', 'd']:
            text = q['opts'].get(opt_key, '')
            right_flowables.append(Paragraph(f"<b>({opt_key})</b> {text}", opt_style))
            right_flowables.append(Spacer(1, 2))
            
        bi_table = Table([[left_flowables, right_flowables]], colWidths=[col_width - 6, col_width - 6])
        bi_table.setStyle(TableStyle([
            ('LINEBEFORE', (1, 0), (1, -1), 1, colors.HexColor('#D0D0D0')),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('RIGHTPADDING', (0, 0), (0, -1), 10),
            ('LEFTPADDING', (1, 0), (1, -1), 10),
            ('TOPPADDING', (0, 0), (-1, -1), 2),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ]))
        q_elements.append(bi_table)
        q_elements.append(Spacer(1, 6))
        
        # 3. Answer Box
        ans_data = [[Paragraph(f"Answer: ({q['ans'].upper()})", ans_style)]]
        ans_table = Table(ans_data, colWidths=[content_width])
        ans_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#F1F8E9')),
            ('BOX', (0, 0), (-1, -1), 1, colors.HexColor('#2E7D32')),
            ('TOPPADDING', (0, 0), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
            ('LEFTPADDING', (0, 0), (-1, -1), 8),
        ]))
        q_elements.append(ans_table)
        q_elements.append(Spacer(1, 4))
        
        # 4. Solution Box
        sol_data = [[Paragraph(f"<b>Solution:</b> {q['sol']}", sol_style)]]
        sol_table = Table(sol_data, colWidths=[content_width])
        sol_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#F9FBE7')),
            ('BOX', (0, 0), (-1, -1), 1, colors.HexColor('#CDDC39')),
            ('TOPPADDING', (0, 0), (-1, -1), 5),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
            ('LEFTPADDING', (0, 0), (-1, -1), 8),
        ]))
        q_elements.append(sol_table)
        q_elements.append(Spacer(1, 4))
        
        # 5. Key Points Box
        kp_html = q['kp'].replace('\n', '<br/>')
        kp_data = [[Paragraph(f"<font color='#1976D2'><b>Key Points:</b></font><br/>{kp_html}", kp_style)]]
        kp_table = Table(kp_data, colWidths=[content_width])
        kp_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#F5F5F5')),
            ('LINEBEFORE', (0, 0), (0, -1), 3.5, colors.HexColor('#1976D2')),
            ('TOPPADDING', (0, 0), (-1, -1), 5),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
            ('LEFTPADDING', (0, 0), (-1, -1), 8),
        ]))
        q_elements.append(kp_table)
        q_elements.append(Spacer(1, 14))
        
        story.append(KeepTogether(q_elements))
        
    doc.build(story, canvasmaker=NumberedCanvas)

# ==========================================
# 6. Telegram Handlers
# ==========================================
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

# ==========================================
# 7. Main Polling Function
# ==========================================
def main():
    app = ApplicationBuilder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    
    print("Telegram polling started...")
    app.run_polling()

if __name__ == "__main__":
    main()
