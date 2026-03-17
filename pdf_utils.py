from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas
from reportlab.platypus import Image
import io
from datetime import datetime
import textwrap
import os
import sys
import logging

# Set up logging
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

# Determine base path for bundled files (PyInstaller) or local files
if getattr(sys, 'frozen', False):
    base_path = sys._MEIPASS
else:
    base_path = os.path.dirname(__file__)

# Path to logo image in static folder
LOGO_PATH = os.path.join(base_path, 'static', 'logo.png')

# Log the logo path for debugging
logger.debug(f"Attempting to access logo at: {LOGO_PATH}")
if not os.path.exists(LOGO_PATH):
    logger.error(f"Logo file not found at: {LOGO_PATH}")
else:
    logger.info(f"Logo file found at: {LOGO_PATH}")

def generate_typing_test_pdf(name, typing_results, handwritten_results=None, excel_quiz_results=None, 
                            excel_score=0, excel_total=0, excel_practical_file=None, 
                            excel_practical_tasks=None, excel_practical_score=None, 
                            excel_sheet_scores=None, location="", distance=0.0, 
                            attempt_number="", signup_date="", dob=""):
    # Sanitize inputs for filename
    sanitized_name = "".join(c for c in name if c.isalnum() or c in (' ', '_')).strip().replace(' ', '_')
    try:
        signup_date_obj = datetime.strptime(signup_date, '%Y-%m-%d %H:%M:%S')
        sanitized_date = signup_date_obj.strftime('%Y%m%d_%H%M%S')
        formatted_signup_date = signup_date_obj.strftime('%d %B %Y')
    except ValueError:
        sanitized_date = 'unknown_date'
        formatted_signup_date = 'Unknown'
    try:
        dob_obj = datetime.strptime(dob, '%Y-%m-%d')
        formatted_dob = dob_obj.strftime('%d %B %Y')
        sanitized_dob = dob_obj.strftime('%Y%m%d')
    except ValueError:
        formatted_dob = 'Unknown'
        sanitized_dob = 'unknown_dob'
    sanitized_attempt = str(attempt_number).replace(' ', '_').lower()
    filename = f"Test_Result_{sanitized_name}_{sanitized_dob}_{sanitized_date}_{sanitized_attempt}.pdf"
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=0.5*inch, leftMargin=0.5*inch, 
                           topMargin=0.75*inch, bottomMargin=1.4*inch)
    styles = getSampleStyleSheet()
    # Define custom styles
    custom_styles = {
        'Title': ParagraphStyle(
            name='CustomTitle',
            parent=styles['Title'],
            fontSize=16,
            spaceAfter=8,
            alignment=1,  # Center
            textColor=colors.black
        ),
        'Heading2': ParagraphStyle(
            name='CustomHeading2',
            parent=styles['Heading2'],
            fontSize=12,
            spaceAfter=6,
            spaceBefore=12,
            textColor=colors.navy
        ),
        'Normal': ParagraphStyle(
            name='CustomNormal',
            parent=styles['Normal'],
            fontSize=8,
            spaceAfter=4
        ),
        'ResultBox': ParagraphStyle(
            name='ResultBox',
            parent=styles['Normal'],
            fontSize=8,
            alignment=2,  # Right
            textColor=colors.white
        ),
        'OverallResult': ParagraphStyle(
            name='OverallResult',
            parent=styles['Normal'],
            fontSize=9,
            alignment=1,  # Center
            textColor=colors.white
        ),
        'Footer': ParagraphStyle(
            name='Footer',
            parent=styles['Normal'],
            fontSize=8,
            alignment=1,  # Center
            textColor=colors.black
        )
    }
    story = []
    # Header function with logo
    def add_header(canvas, doc):
        canvas.saveState()
        logger.debug("Executing add_header in generate_typing_test_pdf")
        try:
            # Add text (Generated on date)
            canvas.setFont('Helvetica', 8)
            canvas.drawString(0.5*inch, letter[1]-0.4*inch, f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            # Add logo to top-right corner
            if os.path.exists(LOGO_PATH):
                logger.debug(f"Drawing logo at: {LOGO_PATH}")
                logo_width = 1 * inch
                logo = Image(LOGO_PATH)
                aspect = logo.drawHeight / logo.drawWidth
                logo.drawHeight = logo_width * aspect
                logo.drawWidth = logo_width
                x_pos = letter[0] - logo_width - 0.5*inch
                y_pos = letter[1] - logo_width * aspect - 0.25*inch
                canvas.drawImage(LOGO_PATH, x_pos, y_pos, 
                                width=logo_width, height=logo_width * aspect, mask='auto')
                logger.debug(f"Logo drawn at x={x_pos}, y={y_pos}, width={logo_width}, height={logo_width * aspect}")
            else:
                logger.error(f"Logo file not found during rendering: {LOGO_PATH}")
                canvas.setFont('Helvetica', 8)
                canvas.drawString(letter[0] - 2*inch, letter[1] - 0.4*inch, "Logo not found")
        except Exception as e:
            logger.error(f"Error rendering logo in generate_typing_test_pdf: {str(e)}")
            canvas.setFont('Helvetica', 8)
            canvas.drawString(letter[0] - 2*inch, letter[1] - 0.4*inch, f"Logo error: {str(e)}")
        canvas.restoreState()
    # Footer function with reordered signatures and copyright
    def add_footer(canvas, doc):
        canvas.saveState()
        logger.debug("Executing add_footer in generate_typing_test_pdf")

        canvas.setFont('Helvetica', 8)

        # Applicant Signature - fully left
        canvas.drawString(0.5*inch, 0.85*inch, "Applicant Signature")

        # Evaluator Signature - centered with respect to page
        canvas.drawCentredString(letter[0]/2, 0.85*inch, "Evaluator Signature")

        # Hiring Manager Signature - fully right
        canvas.drawRightString(letter[0]-0.5*inch, 0.85*inch, "Hiring Manager Signature")

        # Draw horizontal line above copyright
        line_y = 0.65*inch  # slightly above copyright
        canvas.setLineWidth(0.5)
        canvas.line(0.5*inch, line_y, letter[0]-0.5*inch, line_y)

        # Copyright - center bottom
        canvas.drawCentredString(letter[0]/2, 0.5*inch, "© SBL | InterviewAutomation2025")

        canvas.restoreState()

    # Heading
    story.append(Paragraph("Interview Automation Test Results", custom_styles['Title']))
    story.append(Spacer(1, 8))
    # Candidate Information (2 columns, 3 rows)
    story.append(Paragraph("Candidate Information", custom_styles['Heading2']))
    candidate_data = [
        ['Name:', name, 'Location:', location],
        ['Sign-up Date:', formatted_signup_date, 'Date of Birth:', formatted_dob],
        ['Attempt Number:', attempt_number, 'Distance:', f"{distance} km"]
    ]
    candidate_table = Table(candidate_data, colWidths=[1.5*inch, 2*inch, 1.5*inch, 2*inch], hAlign='LEFT')
    candidate_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.white),
        ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
        ('FONTNAME', (2, 0), (2, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (-1, -1), 5),
        ('RIGHTPADDING', (0, 0), (-1, -1), 5),
        ('LINEBEFORE', (2, 0), (2, -1), 2, colors.black),
    ]))
    story.append(candidate_table)
    story.append(Spacer(1, 12))
    # Typing Test Results (table with Attempts, WPM, Accuracy, Time, Result)
    story.append(Paragraph("Typing Test Results", custom_styles['Heading2']))
    # Modified logic: Pass if at least 2 out of 3 attempts meet WPM >= 25 and Accuracy >= 90
    pass_count = sum(1 for result in typing_results if result['wpm'] >= 25 and result['accuracy'] >= 90)
    overall_typing_result = 'Pass' if typing_results and len(typing_results) <= 3 and pass_count >= 2 else 'Fail'
    if typing_results:
        typing_data = [['Attempts', 'WPM', 'Accuracy', 'Time', 'Result']]
        for result in typing_results[:3]:
            pass_status = 'Pass' if result['wpm'] >= 25 and result['accuracy'] >= 90 else 'Fail'
            typing_data.append([
                f"Attempt {result['attempt']}",
                f"{result['wpm']:.1f}",
                f"{result['accuracy']:.1f}%",
                f"{result['time_limit'] // 60}:{result['time_limit'] % 60:02d}",
                pass_status
            ])
        typing_table = Table(typing_data, colWidths=[1.4*inch, 1*inch, 1.2*inch, 1*inch, 1*inch], hAlign='LEFT')
        typing_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('BACKGROUND', (0, 1), (-1, -1), colors.white),
            ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
            ('TOPPADDING', (0, 0), (-1, -1), 4),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('LEFTPADDING', (0, 0), (-1, -1), 5),
            ('RIGHTPADDING', (0, 0), (-1, -1), 5),
        ]))
        story.append(typing_table)
    else:
        typing_table = Table([["No typing test results available."]], colWidths=[6.5*inch], hAlign='LEFT')
        typing_table.setStyle(TableStyle([
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ]))
        story.append(typing_table)
    typing_result_table = Table([[f"Result: {overall_typing_result}"]], colWidths=[2*inch], hAlign='RIGHT')
    typing_result_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.green if overall_typing_result == 'Pass' else colors.red),
        ('TEXTCOLOR', (0, 0), (-1, -1), colors.white),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('BOX', (0, 0), (-1, -1), 1, colors.black),
        ('PADDING', (0, 0), (-1, -1), 4),
    ]))
    story.append(typing_result_table)
    story.append(Spacer(1, 12))
    # Handwritten Verification Results (2 columns, 1 row)
    story.append(Paragraph("Handwritten Verification Results", custom_styles['Heading2']))
    handwritten_correct = sum(1 for result in handwritten_results if result['status'] == 'Correct') if handwritten_results else 0
    handwritten_total = len(handwritten_results) if handwritten_results else 0
    handwritten_result = 'Pass' if handwritten_correct >= 8 else 'Fail'
    if handwritten_results:
        # Handwritten Verification Results table (only 2 columns, reduced width)
        handwritten_data = [['Correct Answers:', f"{handwritten_correct}/{handwritten_total}"]]

        handwritten_table = Table(handwritten_data, colWidths=[1.5*inch, 4*inch], hAlign='LEFT')
        handwritten_table.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,-1), colors.white),
            ('TEXTCOLOR', (0,0), (-1,-1), colors.black),
            ('ALIGN', (0,0), (-1,-1), 'LEFT'),
            ('FONTNAME', (0,0), (0,-1), 'Helvetica-Bold'),
            ('FONTSIZE', (0,0), (-1,-1), 8),
            ('BOTTOMPADDING', (0,0), (-1,-1), 4),
            ('TOPPADDING', (0,0), (-1,-1), 4),
            ('GRID', (0,0), (-1,-1), 0.5, colors.black),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('LEFTPADDING', (0,0), (-1,-1), 5),
            ('RIGHTPADDING', (0,0), (-1,-1), 5),
        ]))
        story.append(handwritten_table)
    else:
        handwritten_table = Table([["No handwritten test results available."]], colWidths=[6.5*inch], hAlign='LEFT')
        handwritten_table.setStyle(TableStyle([
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ]))
        story.append(handwritten_table)
    handwritten_result_table = Table([[f"Result: {handwritten_result}"]], colWidths=[2*inch], hAlign='RIGHT')
    handwritten_result_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.green if handwritten_result == 'Pass' else colors.red),
        ('TEXTCOLOR', (0, 0), (-1, -1), colors.white),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('BOX', (0, 0), (-1, -1), 1, colors.black),
        ('PADDING', (0, 0), (-1, -1), 4),
    ]))
    story.append(handwritten_result_table)
    story.append(Spacer(1, 12))
    # Excel Results (2 columns, 2 rows)
    # Excel Results (merged cells)
    story.append(Paragraph("Excel Results", custom_styles['Heading2']))

    excel_total_marks = excel_score + (sum(excel_sheet_scores.values()) if excel_sheet_scores else 0)
    excel_result = 'Pass' if excel_total_marks >= 15 else 'Fail'

    # Table data
    excel_data = [
        ['Quiz:', f"Correct Answers: {excel_score}/{excel_total}", f"Excel Total: {excel_total_marks}/20"],
        ['Practical:', f"Total Score: {sum(excel_sheet_scores.values()) if excel_sheet_scores else 0}/10", '']
    ]

    col_widths = [1.5*inch, 2*inch, 2*inch]

    excel_table = Table(excel_data, colWidths=col_widths, hAlign='LEFT')
    excel_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.white),
        ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
        ('FONTNAME', (2, 0), (2, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (-1, -1), 5),
        ('RIGHTPADDING', (0, 0), (-1, -1), 5),
        # Merge 'Excel Total' cell across 2 rows vertically
        ('SPAN', (2, 0), (2, 1)),
    ]))

    story.append(excel_table)

    # Result box
    excel_result_table = Table([[f"Result: {excel_result}"]], colWidths=[2*inch], hAlign='RIGHT')
    excel_result_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.green if excel_result == 'Pass' else colors.red),
        ('TEXTCOLOR', (0, 0), (-1, -1), colors.white),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('BOX', (0, 0), (-1, -1), 1, colors.black),
        ('PADDING', (0, 0), (-1, -1), 4),
    ]))
    story.append(excel_result_table)

    story.append(Spacer(1, 12))
    # Overall Result
    story.append(Paragraph("Overall Result", custom_styles['Heading2']))
    pass_count = sum(1 for result in [overall_typing_result, handwritten_result, excel_result] if result == 'Pass')
    overall_result = 'Pass' if pass_count >= 3 else 'Fail'
    overall_table = Table([[f"{overall_result}"]], colWidths=[7*inch], hAlign='LEFT')
    overall_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.green if overall_result == 'Pass' else colors.red),
        ('TEXTCOLOR', (0, 0), (-1, -1), colors.white),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 9),
        ('BOX', (0, 0), (-1, -1), 1, colors.black),
        ('PADDING', (0, 0), (-1, -1), 6),
    ]))
    story.append(overall_table)
    story.append(Spacer(1, 12))
    # Build the document with header and footer
    logger.debug("Building PDF with header and footer in generate_typing_test_pdf")
    doc.build(story, onFirstPage=lambda c, d: (add_header(c, d), add_footer(c, d)), 
             onLaterPages=lambda c, d: (add_header(c, d), add_footer(c, d)))
    buffer.seek(0)
    return buffer, filename

def generate_error_report_pdf(name, handwritten_results=None, excel_quiz_results=None, signup_date="", dob=""):
    # Sanitize inputs for filename
    sanitized_name = "".join(c for c in name if c.isalnum() or c in (' ', '_')).strip().replace(' ', '_')
    try:
        signup_date_obj = datetime.strptime(signup_date, '%Y-%m-%d %H:%M:%S')
        sanitized_date = signup_date_obj.strftime('%Y%m%d_%H%M%S')
    except ValueError:
        sanitized_date = 'unknown_date'
    try:
        dob_obj = datetime.strptime(dob, '%Y-%m-%d')
        sanitized_dob = dob_obj.strftime('%Y%m%d')
    except ValueError:
        sanitized_dob = 'unknown_dob'
    filename = f"Error_Report_{sanitized_name}_{sanitized_dob}_{sanitized_date}.pdf"
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=0.5*inch, leftMargin=0.5*inch, 
                           topMargin=0.75*inch, bottomMargin=0.75*inch)
    styles = getSampleStyleSheet()
    custom_styles = {
        'Title': ParagraphStyle(
            name='CustomTitle',
            parent=styles['Title'],
            fontSize=18,
            spaceAfter=20,
            alignment=1
        ),
        'Heading2': ParagraphStyle(
            name='CustomHeading2',
            parent=styles['Heading2'],
            fontSize=14,
            spaceAfter=12,
            spaceBefore=12,
            textColor=colors.navy
        ),
        'Normal': ParagraphStyle(
            name='CustomNormal',
            parent=styles['Normal'],
            fontSize=10,
            spaceAfter=8
        )
    }
    story = []
    def add_header(canvas, doc):
        canvas.saveState()
        logger.debug("Executing add_header in generate_error_report_pdf")
        try:
            # Add text (Generated on date)
            canvas.setFont('Helvetica', 9)
            canvas.drawString(0.5*inch, letter[1]-0.65*inch, f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            # Add logo to top-right corner
            if os.path.exists(LOGO_PATH):
                logger.debug(f"Drawing logo at: {LOGO_PATH}")
                logo_width = 1 * inch
                logo = Image(LOGO_PATH)
                aspect = logo.drawHeight / logo.drawWidth
                logo.drawHeight = logo_width * aspect
                logo.drawWidth = logo_width
                x_pos = letter[0] - logo_width - 0.5*inch
                y_pos = letter[1] - logo_width * aspect - 0.25*inch
                canvas.drawImage(LOGO_PATH, x_pos, y_pos, 
                                width=logo_width, height=logo_width * aspect, mask='auto')
                logger.debug(f"Logo drawn at x={x_pos}, y={y_pos}, width={logo_width}, height={logo_width * aspect}")
            else:
                logger.error(f"Logo file not found during rendering: {LOGO_PATH}")
                canvas.setFont('Helvetica', 9)
                canvas.drawString(letter[0] - 2*inch, letter[1] - 0.65*inch, "Logo not found")
        except Exception as e:
            logger.error(f"Error rendering logo in generate_error_report_pdf: {str(e)}")
            canvas.setFont('Helvetica', 9)
            canvas.drawString(letter[0] - 2*inch, letter[1] - 0.65*inch, f"Logo error: {str(e)}")
        canvas.restoreState()
    def add_footer(canvas, doc):
        canvas.saveState()
        logger.debug("Executing add_footer in generate_error_report_pdf")
        canvas.setFont('Helvetica', 9)
        page_number = canvas.getPageNumber()
        canvas.drawCentredString(letter[0]/2, 0.5*inch, f"Page {page_number}")
        canvas.restoreState()
    story.append(Paragraph("Interview Automation Error Report", custom_styles['Title']))
    story.append(Spacer(1, 12))
    if handwritten_results:
        correct_count = sum(1 for result in handwritten_results if result['status'] == 'Correct')
        total_count = len(handwritten_results)
        story.append(Paragraph("Handwritten Verification Results", custom_styles['Heading2']))
        story.append(Paragraph(f"Correct Answers: {correct_count}/{total_count}", custom_styles['Normal']))
        incorrect_handwritten = [r for r in handwritten_results if r['status'] == 'Incorrect']
        if incorrect_handwritten:
            story.append(Spacer(1, 12))
            story.append(Paragraph("Incorrect Handwritten Answers:", custom_styles['Heading2']))
            handwritten_data = [['Image', 'User Input', 'Correct Text']]
            for result in incorrect_handwritten:
                user_input = result.get('user_input', 'None')
                correct_text = result.get('correct_text', 'None')
                wrapped_input = textwrap.wrap(user_input, width=50) if user_input else ['None']
                wrapped_correct = textwrap.wrap(correct_text, width=50) if correct_text else ['None']
                user_input_text = '<br/>'.join(wrapped_input)
                correct_text_text = '<br/>'.join(wrapped_correct)
                handwritten_data.append([
                    Paragraph(result.get('image', 'Unknown'), custom_styles['Normal']),
                    Paragraph(user_input_text, custom_styles['Normal']),
                    Paragraph(correct_text_text, custom_styles['Normal'])
                ])
            handwritten_table = Table(handwritten_data, colWidths=[2*inch, 2.5*inch, 2.5*inch])
            handwritten_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, -1), 10),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('LEFTPADDING', (0, 0), (-1, -1), 5),
                ('RIGHTPADDING', (0, 0), (-1, -1), 5),
                ('TOPPADDING', (0, 0), (-1, -1), 5),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 5)
            ]))
            story.append(handwritten_table)
        story.append(Spacer(1, 24))
    if excel_quiz_results:
        excel_score = sum(1 for result in excel_quiz_results if result['status'] == 'Correct')
        excel_total = len(excel_quiz_results)
        story.append(Paragraph("Excel Quiz Results", custom_styles['Heading2']))
        story.append(Paragraph(f"Correct Answers: {excel_score}/{excel_total}", custom_styles['Normal']))
        incorrect_excel_quiz = [r for r in excel_quiz_results if r['status'] == 'Incorrect']
        if incorrect_excel_quiz:
            story.append(Spacer(1, 12))
            story.append(Paragraph("Incorrect Excel Quiz Answers:", custom_styles['Heading2']))
            excel_data = [['Question', 'User Answer', 'Correct Answer']]
            for result in incorrect_excel_quiz:
                question = result.get('question', 'Unknown')
                user_answer = result.get('user_answer', 'None')
                correct_answer = result.get('correct_answer', 'None')
                wrapped_question = textwrap.wrap(question, width=50) if question else ['Unknown']
                wrapped_answer = textwrap.wrap(user_answer, width=50) if user_answer else ['None']
                wrapped_correct = textwrap.wrap(correct_answer, width=50) if correct_answer else ['None']
                question_text = '<br/>'.join(wrapped_question)
                answer_text = '<br/>'.join(wrapped_answer)
                correct_text = '<br/>'.join(wrapped_correct)
                excel_data.append([
                    Paragraph(question_text, custom_styles['Normal']),
                    Paragraph(answer_text, custom_styles['Normal']),
                    Paragraph(correct_text, custom_styles['Normal'])
                ])
            excel_table = Table(excel_data, colWidths=[2.5*inch, 2*inch, 2*inch])
            excel_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, -1), 10),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('LEFTPADDING', (0, 0), (-1, -1), 5),
                ('RIGHTPADDING', (0, 0), (-1, -1), 5),
                ('TOPPADDING', (0, 0), (-1, -1), 5),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 5)
            ]))
            story.append(excel_table)
        story.append(Spacer(1, 24))
    logger.debug("Building PDF with header and footer in generate_error_report_pdf")
    doc.build(story, onFirstPage=lambda c, d: (add_header(c, d), add_footer(c, d)), 
             onLaterPages=lambda c, d: (add_header(c, d), add_footer(c, d)))
    buffer.seek(0)
    return buffer, filename