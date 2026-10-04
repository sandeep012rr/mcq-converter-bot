import os
import re
import subprocess
from threading import Thread
from flask import Flask
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters, ContextTypes
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from deep_translator import GoogleTranslator

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

translator = GoogleTranslator(source='hi', target='en')

def translate_safe(text):
    if not text or not text.strip():
        return ""
    try:
        return translator.translate(text[:4500])
    except Exception:
        return text

# ==========================================
# 3. DOCX Parser
# ==========================================
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
        q_data['ans'] = ans_match.group(1).lower() if ans_match else ""
        
        sol_match = re.search(r'Solution:\s*(.*?)(?=\nKey Points:|\nPositive Marks:|\Z)', block, re.DOTALL)
        q_data['sol'] = sol_match.group(1).strip() if sol_match else ""
        
        kp_match = re.search(r'Key Points:\s*(.*?)(?=\nPositive Marks:|\Z)', block, re.DOTALL)
        q_data['kp'] = kp_match.group(1).strip() if kp_match else ""
        
        questions.append(q_data)
        
    return questions

# ==========================================
# 4. Word Document Builder
# ==========================================
def set_cell_background(cell, fill_hex):
    tcPr = cell._element.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._element.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def create_formatted_docx(questions, output_docx):
    doc = Document()
    
    for section in doc.sections:
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.6)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)
        
        footer = section.footer
        f_p = footer.paragraphs[0]
        f_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        f_run = f_p.add_run("Special Education Needs  |  Contact: 9828625119")
        f_run.font.name = "Times New Roman"
        f_run.font.size = Pt(11)
        f_run.font.bold = True
        f_run.font.color.rgb = RGBColor(80, 80, 80)
        
    for idx, q in enumerate(questions, start=1):
        # 1. Header (Red Title + Marks)
        h_table = doc.add_table(rows=1, cols=2)
        h_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        h_table.autofit = False
        h_table.columns[0].width = Inches(5.5)
        h_table.columns[1].width = Inches(2.0)
        
        r0 = h_table.rows[0].cells[0].paragraphs[0].add_run(f"Question {idx} / प्रश्न {idx}")
        r0.font.name = "Times New Roman"
        r0.font.size = Pt(12)
        r0.font.bold = True
        r0.font.color.rgb = RGBColor(183, 28, 28)
        
        r1 = h_table.rows[0].cells[1].paragraphs[0].add_run("Marks: +1, -0")
        r1.font.name = "Times New Roman"
        r1.font.size = Pt(11)
        r1.font.color.rgb = RGBColor(100, 100, 100)
        h_table.rows[0].cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
        
        p_line = doc.add_paragraph()
        p_line.paragraph_format.space_before = Pt(0)
        p_line.paragraph_format.space_after = Pt(4)
        run_line = p_line.add_run("―" * 58)
        run_line.font.color.rgb = RGBColor(183, 28, 28)
        run_line.font.bold = True
        
        # 2. Bilingual 2-Column Table (Left Hindi, Right English)
        bi_table = doc.add_table(rows=1, cols=2)
        bi_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        bi_table.autofit = False
        bi_table.columns[0].width = Inches(3.7)
        bi_table.columns[1].width = Inches(3.7)
        
        # Left (Hindi)
        c_left = bi_table.rows[0].cells[0]
        p_q_hi = c_left.paragraphs[0]
        r_qh = p_q_hi.add_run(q['q_hi'])
        r_qh.font.name = "Noto Sans Devanagari"
        r_qh.font.size = Pt(11)
        r_qh.font.bold = True
        
        for k in ['a', 'b', 'c', 'd']:
            val = q['opts'].get(k, '')
            p_opt = c_left.add_paragraph()
            p_opt.paragraph_format.space_after = Pt(2)
            r_k = p_opt.add_run(f"({k}) ")
            r_k.font.bold = True
            r_k.font.color.rgb = RGBColor(13, 71, 161)
            r_v = p_opt.add_run(val)
            r_v.font.name = "Noto Sans Devanagari"
            r_v.font.size = Pt(10.5)
            
        # Right (English Translation)
        c_right = bi_table.rows[0].cells[1]
        p_q_en = c_right.paragraphs[0]
        en_q_text = translate_safe(q['q_hi'])
        r_qe = p_q_en.add_run(en_q_text)
        r_qe.font.name = "Times New Roman"
        r_qe.font.size = Pt(11)
        r_qe.font.bold = True
        
        for k in ['a', 'b', 'c', 'd']:
            val_hi = q['opts'].get(k, '')
            val_en = translate_safe(val_hi)
            p_opte = c_right.add_paragraph()
            p_opte.paragraph_format.space_after = Pt(2)
            r_ke = p_opte.add_run(f"({k}) ")
            r_ke.font.bold = True
            r_ke.font.color.rgb = RGBColor(13, 71, 161)
            r_ve = p_opte.add_run(val_en)
            r_ve.font.name = "Times New Roman"
            r_ve.font.size = Pt(10.5)
            
        # 3. Answer Box
        ans_table = doc.add_table(rows=1, cols=1)
        ans_cell = ans_table.rows[0].cells[0]
        set_cell_background(ans_cell, "F1F8E9")
        set_cell_margins(ans_cell, top=80, bottom=80, left=120, right=120)
        p_ans = ans_cell.paragraphs[0]
        r_ans = p_ans.add_run(f"Answer: ({q['ans'].upper()})")
        r_ans.font.name = "Times New Roman"
        r_ans.font.size = Pt(11)
        r_ans.font.bold = True
        r_ans.font.color.rgb = RGBColor(46, 125, 50)
        
        # 4. Solution Box
        sol_table = doc.add_table(rows=1, cols=1)
        sol_cell = sol_table.rows[0].cells[0]
        set_cell_background(sol_cell, "F9FBE7")
        set_cell_margins(sol_cell, top=100, bottom=100, left=120, right=120)
        p_sol = sol_cell.paragraphs[0]
        r_sol_lbl = p_sol.add_run("Solution: ")
        r_sol_lbl.font.bold = True
        r_sol_txt = p_sol.add_run(q['sol'])
        r_sol_txt.font.name = "Noto Sans Devanagari"
        r_sol_txt.font.size = Pt(10)
        
        # 5. Key Points Box
        kp_table = doc.add_table(rows=1, cols=1)
        kp_cell = kp_table.rows[0].cells[0]
        set_cell_background(kp_cell, "F5F5F5")
        set_cell_margins(kp_cell, top=100, bottom=100, left=120, right=120)
        p_kp = kp_cell.paragraphs[0]
        r_kp_lbl = p_kp.add_run("Key Points:\n")
        r_kp_lbl.font.bold = True
        r_kp_lbl.font.color.rgb = RGBColor(25, 118, 210)
        r_kp_txt = p_kp.add_run(q['kp'])
        r_kp_txt.font.name = "Noto Sans Devanagari"
        r_kp_txt.font.size = Pt(10)
        
        p_space = doc.add_paragraph()
        p_space.paragraph_format.space_after = Pt(12)
        
    doc.save(output_docx)

def convert_docx_to_pdf(input_docx, output_pdf):
    out_dir = os.path.dirname(os.path.abspath(output_pdf)) or "."
    cmd = [
        "libreoffice",
        "--headless",
        "--convert-to",
        "pdf",
        os.path.abspath(input_docx),
        "--outdir",
        out_dir
    ]
    subprocess.run(cmd, check=True)
    
    generated = os.path.join(out_dir, os.path.splitext(os.path.basename(input_docx))[0] + ".pdf")
    if os.path.exists(generated) and generated != os.path.abspath(output_pdf):
        os.rename(generated, os.path.abspath(output_pdf))

# ==========================================
# 5. Telegram Handlers
# ==========================================
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Namaste! Apni .docx file upload karein.\n"
        "Main Hindi ka accurate English translation karke 2-column bilingual layout me clean PDF bana kar bhej dunga."
    )

async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc_file = update.message.document
    if not doc_file.file_name.endswith('.docx'):
        await update.message.reply_text("Kripya sirf .docx file upload karein.")
        return

    status_msg = await update.message.reply_text("Bilingual translation aur formatting chal rahi hai, kripya thoda intezar karein...")
    
    file_id = doc_file.file_id
    new_file = await context.bot.get_file(file_id)
    input_path = f"temp_{doc_file.file_name}"
    temp_docx = input_path.replace(".docx", "_formatted.docx")
    output_pdf = input_path.replace(".docx", "_Formatted.pdf")

    await new_file.download_to_drive(input_path)

    try:
        questions = parse_docx(input_path)
        create_formatted_docx(questions, temp_docx)
        convert_docx_to_pdf(temp_docx, output_pdf)

        await update.message.reply_document(
            document=open(output_pdf, "rb"),
            filename="Inclusive_Education_Bilingual_Book.pdf",
            caption="Aapki Bilingual PDF book taiyar hai!\nSpecial Education Needs | 9828625119"
        )
    except Exception as e:
        await update.message.reply_text(f"Error aaya: {str(e)}")
    finally:
        for f in [input_path, temp_docx, output_pdf]:
            if os.path.exists(f):
                os.remove(f)
        await status_msg.delete()

def main():
    app = ApplicationBuilder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    
    print("Telegram polling started...")
    app.run_polling()

if __name__ == "__main__":
    main()
        
