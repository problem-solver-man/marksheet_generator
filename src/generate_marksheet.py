from utils import temp_dir
from data_manager import get_rank
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.pagesizes import A4, landscape
# from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from pypdf import PdfReader, PdfWriter
from io import BytesIO


# Font style(s)
fontpath = temp_dir / "fonts/cour.ttf"
pdfmetrics.registerFont(TTFont('cour', fontpath))

# Function to generate the final marksheet of a student
def generate_marksheets(clas, student, rank, class_folder, template, marksheet_order):

    # Canvas
    overlay_buffer = BytesIO()
    c = canvas.Canvas(overlay_buffer, pagesize=landscape(A4))
    c.setFont("cour", 16)

    # Load default stylesheet and modify
    styles = getSampleStyleSheet()
    style = styles["Normal"]
    style.fontName = "cour"
    style.fontSize = 16
    
    # Overlay values
    c.drawString(744, 475, student[0][0])       # roll
    c.drawString(135, 475, student[0][1])       # learner's name
    c.drawString(640, 475, clas)                # class

    c.drawString(346, 389.5, student[1][0])     # eng th
    c.drawString(346, 371.5, student[1][1])     # eng pr
    c.drawString(475, 380.5, student[1][2])     # eng dis
    c.drawString(518, 380.5, student[1][3])     # eng ht
    c.drawString(563, 380.5, student[1][4])     # eng att
    c.drawString(616, 380.5, student[1][5])     # eng hand
    c.drawString(675, 380.5, student[1][6])     # eng pres

    c.drawString(346, 353.5, student[2][0])     # hin th
    c.drawString(346, 336.5, student[2][1])     # hin pr
    c.drawString(475, 345.5, student[2][2])     # hin dis
    c.drawString(518, 345.5, student[2][3])     # hin ht
    c.drawString(563, 345.5, student[2][4])     # hin att
    c.drawString(616, 345.5, student[2][5])     # hin hand
    c.drawString(675, 345.5, student[2][6])     # hin pres

    c.drawString(346, 319.5, student[3][0])     # math th
    c.drawString(346, 301.5, student[3][1])     # math pr
    c.drawString(475, 310.5, student[3][2])     # math dis
    c.drawString(518, 310.5, student[3][3])     # math ht
    c.drawString(563, 310.5, student[3][4])     # math att
    c.drawString(616, 310.5, student[3][5])     # math hand
    c.drawString(675, 310.5, student[3][6])     # math pres

    c.drawString(346, 284.5, student[4][0])     # sci th
    c.drawString(346, 266.5, student[4][1])     # sci pr
    c.drawString(475, 275.5, student[4][2])     # sci dis
    c.drawString(518, 275.5, student[4][3])     # sci ht
    c.drawString(563, 275.5, student[4][4])     # sci att
    c.drawString(616, 275.5, student[4][5])     # sci hand
    c.drawString(675, 275.5, student[4][6])     # sci pres

    c.drawString(346, 249.5, student[5][0])     # socl th
    c.drawString(346, 231.5, student[5][1])     # socl pr
    c.drawString(475, 240.5, student[5][2])     # socl dis
    c.drawString(518, 240.5, student[5][3])     # socl ht
    c.drawString(563, 240.5, student[5][4])     # socl att
    c.drawString(616, 240.5, student[5][5])     # socl hand
    c.drawString(675, 240.5, student[5][6])     # socl pres

    c.drawString(346, 214.5, student[6][0])     # comp th
    c.drawString(346, 196.5, student[6][1])     # comp pr
    c.drawString(475, 205.5, student[6][2])     # comp dis
    c.drawString(518, 205.5, student[6][3])     # comp ht
    c.drawString(563, 205.5, student[6][4])     # comp att
    c.drawString(616, 205.5, student[6][5])     # comp hand
    c.drawString(675, 205.5, student[6][6])     # comp pres

    c.drawString(346, 178.5, student[7][0])     # gk th
    c.drawString(475, 178.5, student[7][1])     # gk dis
    c.drawString(518, 178.5, student[7][2])     # gk ht
    c.drawString(563, 178.5, student[7][3])     # gk att
    c.drawString(616, 178.5, student[7][4])     # gk hand
    c.drawString(675, 178.5, student[7][5])     # gk pres

    c.drawString(415, 380.5, student[8][0])     # eng tssm
    c.drawString(415, 345.5, student[8][1])     # hin tssm
    c.drawString(415, 310.5, student[8][2])     # math tssm
    c.drawString(415, 275.5, student[8][3])     # sci tssm
    c.drawString(415, 240.5, student[8][4])     # socl tssm
    c.drawString(415, 205.5, student[8][5])     # comp tssm
    c.drawString(415, 178.5, student[8][6])     # gk tssm

    c.drawString(732, 380.5, student[9][0])     # eng tsm
    c.drawString(732, 345.5, student[9][1])     # hin tsm
    c.drawString(732, 310.5, student[9][2])     # math tsm
    c.drawString(732, 275.5, student[9][3])     # sci tsm
    c.drawString(732, 240.5, student[9][4])     # socl tsm
    c.drawString(732, 205.5, student[9][5])     # comp tsm
    c.drawString(732, 178.5, student[9][6])     # gk tsm

    c.drawString(375, 145, student[10][1])      # sum of tssm
    c.drawString(700, 145, student[10][2])      # sum of tsm
    c.drawString(70, 96, student[10][3])        # result
    c.drawString(557, 96, student[10][4])       # percentage

    rank = get_rank(rank)                       # rank    
    p = Paragraph(rank, style)
    width, height = p.wrap(100, 50)
    p.drawOn(c, 712, 100)   

    c.save()    # Save the pdf overlay

    
    # Set buffer pointer to point at start
    overlay_buffer.seek(0)

    # Read overlay pdf from buffer
    overlay_pdf = PdfReader(overlay_buffer)

    # Use pypdf to merge overlay pdf with template
    writer = PdfWriter()
    page = template.pages[0]
    page.merge_page(overlay_pdf.pages[0])
    writer.add_page(page)

    # Decide the order of generated marksheets
    if marksheet_order == "Order by Rank":
        filename = f"{class_folder}/Rank_{rank[0]}_{student[0][1].replace(' ', '_')}.pdf"
    else:
        filename = f"{class_folder}/Roll_{student[0][0]}_{student[0][1].replace(' ', '_')}.pdf"

    # Save the final pdf marksheets of a class
    with open(filename, "wb") as f:
        writer.write(f)
