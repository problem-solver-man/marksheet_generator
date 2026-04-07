from utils import base_dir, temp_dir
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4, landscape
# from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from pypdf import PdfReader, PdfWriter
from io import BytesIO
from typing import Callable
import data_manager as dm
import logging
import re


# Initialize logger
logger = logging.getLogger(__name__)

# Font style(s)
fontpath1 = temp_dir / "fonts/cour.ttf"
fontpath2 = temp_dir / "fonts/timesbd.ttf"
pdfmetrics.registerFont(TTFont('cour', fontpath1))
pdfmetrics.registerFont(TTFont('timesbd', fontpath2))


# ------------------------------
# Draw overlay PDF of a Student
# ------------------------------

# Convert evaluation term in html format
def _eval(term) -> str:
    if term == "1st":
        return "1<sup>st</sup>"
    elif term == "2nd":
        return "2<sup>nd</sup>"
    elif term == "3rd":
        return "3<sup>rd</sup>"
    return term


def generate_overlay(student, eval_term, eval_year):
    # Canvas
    overlay_buffer = BytesIO()
    c = canvas.Canvas(overlay_buffer, pagesize=landscape(A4))

    # Load default stylesheet
    styles = getSampleStyleSheet()
    style = styles["Normal"]

    # --------------------- Overlay values ---------------------

    # evaluation term and year
    style.fontName = "timesbd"
    style.fontSize = 22
    style.alignment = TA_CENTER

    eval_term_year = f"{_eval(eval_term)} Evaluation Sheet - {eval_year[:4]}"
    p = Paragraph(eval_term_year, style)
    width, height = p.wrap(315, 50)
    p.drawOn(c, 246.5, 510)

    c.setFont("cour", 16)
    c.drawString(744, 475, student[0][0])       # roll
    c.drawString(135, 475, student[0][1])       # learner's name
    c.drawString(640, 475, student[0][2])       # class

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

    style.fontName = "cour"                     # rank
    style.fontSize = 16
    style.alignment = TA_LEFT

    p = Paragraph(student[10][5], style)
    width, height = p.wrap(100, 50)
    p.drawOn(c, 712, 100)

    c.save()                # Finish the overlay drawing

    overlay_buffer.seek(0)  # Set buffer pointer to point at start
    overlay_pdf = PdfReader(overlay_buffer)  # Read overlay pdf from buffer

    return overlay_pdf



# -------------------------
# Final marksheet(s)
# -------------------------

# Define a no-op function
def _noop(*args, **kwargs):
    pass
    

# Function to generate and save the marksheet
def generate_marksheets(
        excel_path, output_folder, 
        eval_term, eval_year,
        selected_class, marksheet_order, output_type, 
        progress_callback: Callable[[str, int, int], None] = _noop
) -> None:

    # Load the data from excel file sheet by sheet
    wb = dm.load_data(excel_path, selected_class)

    # Process a sheet (a class of students)
    for ws in wb:

        # Log the current class name being processed
        logger.info(
            f"Processing {ws.title} "
            f"by {marksheet_order[9:]} "
            f"for {output_type} "
            f"for {eval_term} term {eval_year}"
        )

        # Read the template once per class and return if not exists
        if ws.title.replace("Class-", "") in "1A1B2A2B34A4B5":
            template_path = temp_dir / "templates/marksheet-format-12345.pdf"
        else:
            template_path = temp_dir / "templates/marksheet-format-67.pdf"
            
        if template_path.is_file():
            with open(template_path, "rb") as f:
                template_bytes = f.read()
        else:
            logger.warning(
                f"Marksheets generation aborted - "
                f"missing template '{template_path.name}'."
            )
            return

        # Format all students data in a list from worksheet
        student_list = dm.prepare_data(ws, marksheet_order)
        total_students = len(student_list)

        # Check if student data exist in current class sheet
        if total_students == 0:
            logger.warning(
                f"We couldn’t find any student data, "
                f"so marksheets were not created for {ws.title}"
            )
            continue

        # Create single writer for single PDF type
        is_single_pdf = output_type == "Single PDF"
        combined_writer = PdfWriter() if is_single_pdf else None

        # Path where marksheets of current class will be saved
        class_folder = output_folder / ws.title
        class_folder.mkdir(parents=True, exist_ok=True)

        for rank, student in enumerate(student_list):
            # Convert all values to string in student
            student = dm.polish_data(student)

            # Generate overlay pdf
            overlay_pdf = generate_overlay(student, eval_term, eval_year)

            # Get fresh template
            template = PdfReader(BytesIO(template_bytes))

            # Merge overlay and template using pypdf
            page = template.pages[0]
            page.merge_page(overlay_pdf.pages[0])

            # Current student's roll, name, and rank for unique file names
            _roll = student[0][0]
            _name = re.sub(r'[^a-zA-Z0-9_-]', '', student[0][1].replace(' ', '_'))
            _rank = rank + 1

            # ================ Single PDF type =================
            if combined_writer is not None:
                combined_writer.add_page(page)

            # =============== Separate PDFs type ===============
            else:
                # Create writer for each student for separate PDFs
                writer = PdfWriter()
                writer.add_page(page)

                # Decide the order of generated marksheets
                if marksheet_order == "Order by Rank":
                    filename = f"{class_folder}/Rank_{_rank}_{_name}.pdf"
                else:
                    filename = f"{class_folder}/Roll_{_roll}_{_name}.pdf"

                # Save the final pdf marksheets of a student
                with open(filename, "wb") as f:
                    writer.write(f)

            # Update the progress bar
            progress_callback(ws.title, _rank, total_students)
        
        if is_single_pdf:
            # Just to use assert keyword; And because I'm 100% sure
            # Otherwise, [if combined_writer is not None:] is also fine
            assert combined_writer is not None
            if marksheet_order == "Order by Rank":
                filenames = f"{class_folder}/Rank_wise_all_{ws.title}_marksheets.pdf"
            else:
                filenames = f"{class_folder}/Roll_wise_all_{ws.title}_marksheets.pdf"

            # Save the final pdf marksheets of a class
            with open(filenames, "wb") as f:
                combined_writer.write(f)
        
        # Log the successful generation of marksheets
        logger.info(f"Generated marksheets folder = {class_folder}")
