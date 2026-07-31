"""
create_word_doc.py -- Generates a styled Word document (docx) of all P1 and P2 email templates.
"""
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def create_element(name):
    return OxmlElement(name)

def set_cell_background(cell, fill_hex):
    shading_xml = f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>'
    cell._tc.get_or_add_tcPr().append(parse_xml(shading_xml))

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def add_callout(doc, text_paragraphs):
    # Create a 1x1 table to act as a callout block
    table = doc.add_table(rows=1, cols=1)
    table.alignment = docx.enum.table.WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    
    cell = table.cell(0, 0)
    cell.width = Inches(6.0)
    set_cell_background(cell, "F2F4F7")
    set_cell_margins(cell, top=140, bottom=140, left=200, right=200)
    
    # Set left border to thick blue, others to none
    tcPr = cell._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>\n'
        f'  <w:top w:val="none"/>\n'
        f'  <w:left w:val="single" w:sz="36" w:space="0" w:color="1F4E78"/>\n'
        f'  <w:bottom w:val="none"/>\n'
        f'  <w:right w:val="none"/>\n'
        f'</w:tcBorders>'
    )
    tcPr.append(borders)
    
    # Write paragraphs in the callout
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.15
    
    for i, line in enumerate(text_paragraphs):
        if i > 0:
            p = cell.add_paragraph()
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(4)
            p.paragraph_format.line_spacing = 1.15
        
        # Format Subject vs Body
        if line.startswith("Subject:"):
            run = p.add_run("Subject: ")
            run.bold = True
            run.font.name = "Arial"
            run.font.size = Pt(10.5)
            run.font.color.rgb = RGBColor(0x1F, 0x4E, 0x78)
            
            run_val = p.add_run(line[len("Subject:"):].strip())
            run_val.italic = True
            run_val.font.name = "Arial"
            run_val.font.size = Pt(10.5)
            run_val.font.color.rgb = RGBColor(0x22, 0x22, 0x22)
        elif line.startswith("Body:"):
            run = p.add_run("Body:\n")
            run.bold = True
            run.font.name = "Arial"
            run.font.size = Pt(10)
            run.font.color.rgb = RGBColor(0x1F, 0x4E, 0x78)
            
            run_val = p.add_run(line[len("Body:"):].strip())
            run_val.font.name = "Courier New"
            run_val.font.size = Pt(9.5)
            run_val.font.color.rgb = RGBColor(0x33, 0x33, 0x33)
        else:
            run = p.add_run(line)
            run.font.name = "Arial"
            run.font.size = Pt(10)
            run.font.color.rgb = RGBColor(0x44, 0x44, 0x44)

def build_docx():
    doc = docx.Document()
    
    # Margins
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)
        
    # Document Title
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.LEFT
    title.paragraph_format.space_after = Pt(18)
    run_title = title.add_run("DriverAI Hiring Agent — Email Templates Catalog")
    run_title.font.name = "Arial"
    run_title.font.size = Pt(22)
    run_title.bold = True
    run_title.font.color.rgb = RGBColor(0x1F, 0x4E, 0x78)
    
    # Intro
    intro = doc.add_paragraph()
    intro.paragraph_format.space_after = Pt(12)
    run_intro = intro.add_run("This catalog contains all automated emails sent to candidates and administrators during Phase 1 (Intake & Routing) and Phase 2 (AI Extraction & Scoring).")
    run_intro.font.name = "Arial"
    run_intro.font.size = Pt(11)
    run_intro.font.color.rgb = RGBColor(0x44, 0x44, 0x44)

    # PAGE BREAK / DIVISION
    doc.add_paragraph().add_run("―" * 45).font.color.rgb = RGBColor(0xCC, 0xCC, 0xCC)

    # =========================================================================
    # SECTION 1: PHASE 1
    # =========================================================================
    h1 = doc.add_paragraph()
    h1.paragraph_format.space_before = Pt(18)
    h1.paragraph_format.space_after = Pt(8)
    h1_run = h1.add_run("1. Phase 1 (Intake & Routing) Email Templates")
    h1_run.bold = True
    h1_run.font.name = "Arial"
    h1_run.font.size = Pt(16)
    h1_run.font.color.rgb = RGBColor(0x1F, 0x4E, 0x78)
    
    p1_desc = doc.add_paragraph()
    p1_desc.paragraph_format.space_after = Pt(12)
    p1_desc_run = p1_desc.add_run("All Phase 1 emails are hardcoded in the Power Automate flow package. They send from apply@driverai.io.")
    p1_desc_run.italic = True
    p1_desc_run.font.name = "Arial"
    p1_desc_run.font.size = Pt(10.5)
    p1_desc_run.font.color.rgb = RGBColor(0x66, 0x66, 0x66)

    # EMAIL 1
    doc.add_heading("Email 1: Application Acknowledgment", level=2)
    p = doc.add_paragraph()
    p.add_run("Trigger Condition: ").bold = True
    p.add_run("A brand new candidate applies with a valid resume (.pdf or .docx).\n")
    p.add_run("Database Action: ").bold = True
    p.add_run("Creates a new row in the master table (Status = New Email Received).\n")
    p.add_run("Mailbox Action: ").bold = True
    p.add_run("Moved to Archive and marked as read.")
    for run in p.runs:
        run.font.name = "Arial"
        run.font.size = Pt(10)
    
    add_callout(doc, [
        "Subject: Application Received - DriverAI (Ref: {AppRef})",
        "Body: \nHello,\nThanks for applying to DriverAI. Your application has been received and is under review.\nReference number: {AppRef}\nIf you are selected, a member of our team will contact you to discuss next steps. You do not need to reply unless you are updating your resume. To do that, send a new or reply email to apply@driverai.io with the subject \"Update - {AppRef}\" and attach the new file.\nGood luck on your next journey.\nWarm regards, The DriverAI Recruiting Team\n\n-------------------------------------------------\nThis is an auto-generated email and this mailbox is not monitored."
    ])
    doc.add_paragraph() # space

    # EMAIL 2
    doc.add_heading("Email 2: Request CV", level=2)
    p = doc.add_paragraph()
    p.add_run("Trigger Condition: ").bold = True
    p.add_run("An email is received without a resume attached, but contains application keywords (resume, cv, apply, hiring).\n")
    p.add_run("Database Action: ").bold = True
    p.add_run("No database row is created (the database is resume-first only).\n")
    p.add_run("Mailbox Action: ").bold = True
    p.add_run("Moved to Archive and marked as read.")
    for run in p.runs:
        run.font.name = "Arial"
        run.font.size = Pt(10)
        
    add_callout(doc, [
        "Subject: Please Attach Your Resume - DriverAI",
        "Body: \nHello,\nThanks for your interest in DriverAI. We didn't see a resume attached, so your application isn't complete yet.\nPlease reply with your resume attached in PDF or Word (.docx) format. Once we have it, your application goes under review, and if you are selected, a member of our team will contact you to discuss next steps.\nGood luck on your next journey.\nWarm regards, The DriverAI Recruiting Team\n\n-------------------------------------------------\nThis is an auto-generated email and this mailbox is not monitored."
    ])
    doc.add_paragraph()

    # EMAIL 3
    doc.add_heading("Email 3: Duplicate Application Notice", level=2)
    p = doc.add_paragraph()
    p.add_run("Trigger Condition: ").bold = True
    p.add_run("An applicant resubmits within the 90-day window without quoting a reference ID (up to 3 total emails).\n")
    p.add_run("Database Action: ").bold = True
    p.add_run("Overwrites the resume file on SharePoint and re-queues the row (Status = New Email Received).\n")
    p.add_run("Mailbox Action: ").bold = True
    p.add_run("Moved to Archive and marked as read.")
    for run in p.runs:
        run.font.name = "Arial"
        run.font.size = Pt(10)
        
    add_callout(doc, [
        "Subject: Your DriverAI Application Is Already On File (Ref: {app_id})",
        "Body: \nHello,\nYour application is already on file with DriverAI and under review, so there's no need to resubmit.\n\n[For Contacts 1 and 2]\nIf you are selected, a member of our team will contact you to discuss next steps. You do not need to reply unless you are updating your resume. To do that, send a new or reply email to apply@driverai.io with the subject \"Update - {app_id}\" and attach the new file.\n\n[Swapped on Contact 3 (Final Notice)]\nThis is our final automated reply regarding this application. If there is a match, a member of our team will contact you directly. You do not need to reply, and you are welcome to apply again after 90 days.\n\nGood luck on your next journey.\nWarm regards, The DriverAI Recruiting Team\n\n-------------------------------------------------\nThis is an auto-generated email and this mailbox is not monitored."
    ])
    doc.add_paragraph()

    # EMAIL 4
    doc.add_heading("Email 4: Wrong File Format Notice", level=2)
    p = doc.add_paragraph()
    p.add_run("Trigger Condition: ").bold = True
    p.add_run("Candidate emails with an attachment that is not .pdf or .docx (e.g. .txt, .jpg, .zip).\n")
    p.add_run("Database Action: ").bold = True
    p.add_run("No database row is created.\n")
    p.add_run("Mailbox Action: ").bold = True
    p.add_run("Moved to Archive and marked as read.")
    for run in p.runs:
        run.font.name = "Arial"
        run.font.size = Pt(10)
        
    add_callout(doc, [
        "Subject: Please Resend Your Resume as PDF or Word - DriverAI",
        "Body: \nHello,\nThanks for your interest in DriverAI. We received your submission, but we couldn't open the attached file.\nPlease reply with your resume as a PDF (.pdf) or Word (.docx) file and we'll process it right away. If you are selected, a member of our team will contact you to discuss next steps.\nGood luck on your next journey.\nWarm regards, The DriverAI Recruiting Team\n\n-------------------------------------------------\nThis is an auto-generated email and this mailbox is not monitored."
    ])
    doc.add_paragraph()

    # EMAIL 5
    doc.add_heading("Email 5: Update Acknowledgment", level=2)
    p = doc.add_paragraph()
    p.add_run("Trigger Condition: ").bold = True
    p.add_run("Candidate replies to their original thread quoting their reference ID and attaches a new resume (up to 3 total emails).\n")
    p.add_run("Database Action: ").bold = True
    p.add_run("Saves the updated resume and re-queues the row (Status = New Email Received) to be re-scored.\n")
    p.add_run("Mailbox Action: ").bold = True
    p.add_run("Moved to Archive and marked as read.")
    for run in p.runs:
        run.font.name = "Arial"
        run.font.size = Pt(10)
        
    add_callout(doc, [
        "Subject: Updated Resume Received - DriverAI (Ref: {app_id})",
        "Body: \nHello,\nThanks for sending your updated resume. Your application now reflects the latest version and our team will review it.\n\n[For Contacts 1 and 2]\nIf you are selected, a member of our team will contact you to discuss next steps. You do not need to reply unless you are updating your resume again. To do that, send a new or reply email to apply@driverai.io with the subject \"Update - {app_id}\" and attach the new file.\n\n[Swapped on Contact 3 (Final Notice)]\nThis is our final automated reply regarding this application. If there is a match, a member of our team will contact you directly. You do not need to reply, and you are welcome to apply again after 90 days.\n\nGood luck on your next journey.\nWarm regards, The DriverAI Recruiting Team\n\n-------------------------------------------------\nThis is an auto-generated email and this mailbox is not monitored."
    ])
    doc.add_paragraph()

    # EMAIL 6
    doc.add_heading("Email 6: Noted Follow-up", level=2)
    p = doc.add_paragraph()
    p.add_run("Trigger Condition: ").bold = True
    p.add_run("Candidate replies quoting their reference ID but does not attach a file (up to 3 total emails).\n")
    p.add_run("Database Action: ").bold = True
    p.add_run("Increments Application Updates but keeps the row status unchanged (not re-scored).\n")
    p.add_run("Mailbox Action: ").bold = True
    p.add_run("Moved to Archive and marked as read.")
    for run in p.runs:
        run.font.name = "Arial"
        run.font.size = Pt(10)
        
    add_callout(doc, [
        "Subject: Message Received - DriverAI (Ref: {app_id})",
        "Body: \nHello,\nThanks for following up. Your message has been noted alongside your application.\n\n[For Contacts 1 and 2]\nIf you are selected, a member of our team will contact you to discuss next steps. You do not need to reply unless you are updating your resume. To send one, reply with the subject \"Update - {app_id}\" and attach the file (PDF or Word).\n\n[Swapped on Contact 3 (Final Notice)]\nThis is our final automated reply regarding this application. If there is a match, a member of our team will contact you directly. You do not need to reply, and you are welcome to apply again after 90 days.\n\nGood luck on your next journey.\nWarm regards, The DriverAI Recruiting Team\n\n-------------------------------------------------\nThis is an auto-generated email and this mailbox is not monitored."
    ])
    doc.add_paragraph()

    # PAGE BREAK / DIVISION
    doc.add_page_break()

    # =========================================================================
    # SECTION 2: PHASE 2
    # =========================================================================
    h2 = doc.add_paragraph()
    h2.paragraph_format.space_before = Pt(18)
    h2.paragraph_format.space_after = Pt(8)
    h2_run = h2.add_run("2. Phase 2 (AI Extraction & Scoring) Email Templates")
    h2_run.bold = True
    h2_run.font.name = "Arial"
    h2_run.font.size = Pt(16)
    h2_run.font.color.rgb = RGBColor(0x1F, 0x4E, 0x78)
    
    p2_desc = doc.add_paragraph()
    p2_desc.paragraph_format.space_after = Pt(12)
    p2_desc_run = p2_desc.add_run("Phase 2 emails are sent dynamically by the Python application (hiring-bot). They are configured in config.yaml.")
    p2_desc_run.italic = True
    p2_desc_run.font.name = "Arial"
    p2_desc_run.font.size = Pt(10.5)
    p2_desc_run.font.color.rgb = RGBColor(0x66, 0x66, 0x66)

    # LACK OF DATA
    doc.add_heading("Email 1: Lack of Data / Missing Info Nudge", level=2)
    p = doc.add_paragraph()
    p.add_run("Trigger Condition: ").bold = True
    p.add_run("A scored candidate is missing critical fields (Phone, Location, or Portfolios) on their resume.\n")
    p.add_run("Database Action: ").bold = True
    p.add_run("Tracks nudge state in the Info Request Sent column.")
    for run in p.runs:
        run.font.name = "Arial"
        run.font.size = Pt(10)
        
    add_callout(doc, [
        "Subject: Quick follow-up on your {company} application (Ref: {ref})",
        "Body: \n<div style=\"font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#222;line-height:1.6;\">\n<p>Hi {greeting},</p>\n<p>Thanks again for applying to {company} - your application (Ref: <strong>{ref}</strong>) is with our recruiting team.</p>\n<p>While reviewing it, we noticed we don't have {missing_list} on file. Replying to this email with an updated resume that includes that information would help us move your application along faster.</p>\n<p>No action is needed if you'd rather not share it - we'll continue reviewing your application either way.</p>\n<p>Warm regards,<br>The {company} Recruiting Team</p>\n</div>"
    ])
    doc.add_paragraph()

    # REJECTION
    doc.add_heading("Email 2: Rejection — Non-USA Location", level=2)
    p = doc.add_paragraph()
    p.add_run("Trigger Condition: ").bold = True
    p.add_run("The candidate's resume/profile indicates they are based outside the United States.\n")
    p.add_run("Database Action: ").bold = True
    p.add_run("Moves the candidate row to the Rejected sheet and marks Status = Rejected - Non-USA Location and Decline Sent = Sent.")
    for run in p.runs:
        run.font.name = "Arial"
        run.font.size = Pt(10)
        
    add_callout(doc, [
        "Subject: Update on Your {company} Application",
        "Body: \n<div style=\"font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#222;line-height:1.7;\">\n<p>Hi {greeting},</p>\n<p>Thank you for your interest in <strong>{company}</strong> and for taking the time to apply.</p>\n<p>After reviewing your application, we are currently able to consider candidates based in the United States only, so we are unable to move forward at this time.</p>\n<p>We truly appreciate your interest. We are growing and hope to expand into more regions in the future, and we would be glad to have you apply again when we do.</p>\n<p>You do not need to reply. We wish you all the best in your search.</p>\n<p>Warm regards,<br>The {company} Recruiting Team</p>\n</div>"
    ])
    doc.add_paragraph()

    # ADMIN ALERT
    doc.add_heading("Email 3: Administrator Alert (Error Notification)", level=2)
    p = doc.add_paragraph()
    p.add_run("Trigger Condition: ").bold = True
    p.add_run("A scored candidate's row encounters a processing exception 3 times in a row (e.g. unreadable resume corruptions).\n")
    p.add_run("Database Action: ").bold = True
    p.add_run("Moves the candidate to the Rejected sheet under Status = Rejected - Processing Error.")
    for run in p.runs:
        run.font.name = "Arial"
        run.font.size = Pt(10)
        
    add_callout(doc, [
        "Subject: [Hiring Agent] Row given up after {SCORE_RETRY_MAX} failures ({app_id})",
        "Body: \n<p>A candidate row failed processing {SCORE_RETRY_MAX} times and has been moved to the Rejected sheet under 'Rejected - Processing Error' instead of being retried forever.</p>\n<ul>\n  <li>Application ID: {app_id}</li>\n  <li>Sender: {sender_email}</li>\n  <li>Reason: {reason}</li>\n</ul>\n<p>Check P2_Logs for the full error history on this row.</p>"
    ])

    # Header and Footer styling for Headings
    for paragraph in doc.paragraphs:
        if paragraph.style.name.startswith('Heading 2'):
            paragraph.paragraph_format.space_before = Pt(12)
            paragraph.paragraph_format.space_after = Pt(6)
            for run in paragraph.runs:
                run.font.name = "Arial"
                run.font.size = Pt(12.5)
                run.bold = True
                run.font.color.rgb = RGBColor(0x1F, 0x4E, 0x78)

    # Save to disk
    output_filename = "HiringAgent_Email_Templates.docx"
    doc.save(output_filename)
    print(f"[OK] Word document saved to {output_filename}")

if __name__ == "__main__":
    build_docx()
