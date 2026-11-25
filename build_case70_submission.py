from __future__ import annotations

import csv
import io
import json
import math
import unicodedata
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.dml.color import RGBColor as SlideColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Inches as PInches, Pt as PPt


ROOT = Path(__file__).resolve().parent
PROJECT = ROOT / "Case70_Vehicle_Service_Booking"
LAB = ROOT / "Case70_Lab_Outputs"
EVIDENCE = ROOT / "Case70_Evidence_Screenshots"
REPORT_DIR = ROOT / "Case70_Report"
REPORT_DIR.mkdir(exist_ok=True)

DOCX_PATH = REPORT_DIR / "Case70_57_Suhas_Final.docx"
PDF_PATH = REPORT_DIR / "Case70_57_Suhas_Final.pdf"
PPTX_PATH = REPORT_DIR / "Case70_57_Suhas_Final_Presentation.pptx"

NAVY = "142B4A"
BLUE = "2F75DF"
TEAL = "0D8B8B"
PALE = "EAF2FA"
PALE_TEAL = "E7F5F4"
INK = "24364B"
GREY = "607086"
AMBER = "A75B00"
RED = "A63432"


def load_font(size: int, bold: bool = False):
    candidates = [
        r"C:\Windows\Fonts\segoeuib.ttf" if bold else r"C:\Windows\Fonts\segoeui.ttf",
        r"C:\Windows\Fonts\arialbd.ttf" if bold else r"C:\Windows\Fonts\arial.ttf",
    ]
    for candidate in candidates:
        if Path(candidate).exists():
            return ImageFont.truetype(candidate, size)
    return ImageFont.load_default()


def make_figure(kind: str) -> io.BytesIO:
    image = Image.new("RGB", (1400, 620), "white")
    draw = ImageDraw.Draw(image)
    title_font = load_font(30, True)
    label_font = load_font(21, True)
    body_font = load_font(18)
    if kind == "architecture":
        draw.text((50, 25), "Vehicle Service Booking - Prototype Architecture", fill=f"#{NAVY}", font=title_font)
        boxes = [
            (60, 230, 330, 400, "Customer / Staff\nBrowser UI"),
            (420, 230, 690, 400, "Flask routes\nBooking + consent"),
            (780, 230, 1050, 400, "SQLite database\n6 entities"),
            (1140, 230, 1340, 400, "Reminder\nscheduler"),
        ]
        for x1, y1, x2, y2, text in boxes:
            draw.rounded_rectangle((x1, y1, x2, y2), 20, fill=f"#{PALE}", outline=f"#{BLUE}", width=4)
            draw.multiline_text((x1 + 20, y1 + 50), text, fill=f"#{NAVY}", font=label_font, spacing=12, align="center")
        for x1, x2 in [(330, 420), (690, 780), (1050, 1140)]:
            draw.line((x1, 315, x2, 315), fill=f"#{TEAL}", width=6)
            draw.polygon([(x2 - 18, 302), (x2, 315), (x2 - 18, 328)], fill=f"#{TEAL}")
        draw.text((80, 470), "Health endpoint: /health  |  Utilisation API: /api/utilisation  |  Reminders: /api/run-reminders", fill=f"#{GREY}", font=body_font)
    elif kind == "incremental":
        draw.text((50, 25), "Incremental Release Path", fill=f"#{NAVY}", font=title_font)
        steps = [
            ("Increment 1", "Workshop + vehicle\nselection", "Review"),
            ("Increment 2", "Future-dated\nbooking flow", "Review"),
            ("Increment 3", "Reminder scheduling\n+ duplicate safety", "Review"),
            ("Increment 4", "Consent, weekday\nutilisation + hardening", "Release"),
        ]
        start = 50
        for i, (name, body, checkpoint) in enumerate(steps):
            x1 = start + i * 330
            draw.rounded_rectangle((x1, 190, x1 + 270, 420), 18, fill=f"#{PALE if i % 2 == 0 else PALE_TEAL}", outline=f"#{BLUE if i % 2 == 0 else TEAL}", width=4)
            draw.text((x1 + 20, 220), name, fill=f"#{NAVY}", font=label_font)
            draw.multiline_text((x1 + 20, 275), body, fill=f"#{INK}", font=body_font, spacing=8)
            draw.text((x1 + 20, 375), checkpoint, fill=f"#{TEAL}", font=body_font)
            if i < 3:
                draw.line((x1 + 270, 305, x1 + 310, 305), fill=f"#{TEAL}", width=5)
                draw.polygon([(x1 + 300, 294), (x1 + 315, 305), (x1 + 300, 316)], fill=f"#{TEAL}")
        draw.text((60, 485), "Each increment is reviewed against FR-01...FR-10 and accepted before expanding scope.", fill=f"#{GREY}", font=body_font)
    elif kind == "velocity":
        draw.text((50, 25), "Forecast Burn-down (Assumed Velocity)", fill=f"#{NAVY}", font=title_font)
        x0, y0, x1, y1 = 120, 110, 1280, 500
        draw.line((x0, y1, x1, y1), fill=f"#{GREY}", width=3)
        draw.line((x0, y0, x0, y1), fill=f"#{GREY}", width=3)
        data = [82, 56, 31, 5, 0]
        labels = ["Start", "Sprint 1", "Sprint 2", "Sprint 3", "Sprint 4"]
        points = []
        for i, value in enumerate(data):
            x = x0 + i * ((x1 - x0) // 4)
            y = y1 - int(value / 82 * (y1 - y0))
            points.append((x, y))
            draw.ellipse((x - 9, y - 9, x + 9, y + 9), fill=f"#{BLUE}")
            draw.text((x - 35, y - 45), str(value), fill=f"#{NAVY}", font=label_font)
            draw.text((x - 45, y1 + 18), labels[i], fill=f"#{GREY}", font=body_font)
        draw.line(points, fill=f"#{BLUE}", width=6)
        draw.text((150, 530), "82-point planning backlog / mean assumed velocity 25.67 points = 3.2 sprint-equivalents; plan 4 whole sprints.", fill=f"#{INK}", font=body_font)
    elif kind == "monitor":
        rows = list(csv.DictReader((PROJECT / "monitor_log.csv").open(encoding="utf-8-sig")))
        rows = rows[-10:]
        values = [float(r["response_ms"]) for r in rows]
        draw.text((50, 25), "Actual /health Monitor - 10 Samples", fill=f"#{NAVY}", font=title_font)
        x0, y0, x1, y1 = 100, 100, 1320, 500
        draw.line((x0, y1, x1, y1), fill=f"#{GREY}", width=3)
        draw.line((x0, y0, x0, y1), fill=f"#{GREY}", width=3)
        threshold_y = y1 - int(500 / max(600, max(values)) * (y1 - y0))
        draw.line((x0, threshold_y, x1, threshold_y), fill=f"#{RED}", width=3)
        draw.text((x0 + 8, threshold_y - 30), "500 ms alert threshold", fill=f"#{RED}", font=body_font)
        bar_width = 70
        gap = (x1 - x0 - 10 * bar_width) / 9
        for i, val in enumerate(values):
            x = int(x0 + i * (bar_width + gap))
            h = int(val / max(600, max(values)) * (y1 - y0))
            color = RED if val > 500 else BLUE
            draw.rectangle((x, y1 - h, x + bar_width, y1), fill=f"#{color}")
            draw.text((x + 5, y1 + 12), str(i + 1), fill=f"#{GREY}", font=body_font)
            draw.text((x, max(y0, y1 - h - 28)), f"{val:.0f}", fill=f"#{NAVY}", font=body_font)
        draw.text((100, 550), f"HTTP 200: {sum(r['status_code']=='200' for r in rows)}/10. Alerted samples >500 ms: {sum(v>500 for v in values)}.", fill=f"#{INK}", font=body_font)
    buffer = io.BytesIO()
    image.save(buffer, format="PNG")
    buffer.seek(0)
    return buffer


def set_cell_shading(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    tc_pr.append(shd)


def set_cell_margins(cell, top=80, start=90, bottom=80, end=90):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for m, v in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tc_mar.find(qn(f"w:{m}"))
        if node is None:
            node = OxmlElement(f"w:{m}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(v))
        node.set(qn("w:type"), "dxa")


def add_page_field(paragraph):
    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    text = OxmlElement("w:t")
    text.text = "1"
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.extend([begin, instr, separate, text, end])


def add_doc_table(doc, headers, rows, font_size=8.5, widths=None):
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True
    for i, header in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = str(header)
        set_cell_shading(cell, NAVY)
        set_cell_margins(cell)
        for p in cell.paragraphs:
            for run in p.runs:
                run.bold = True
                run.font.color.rgb = RGBColor(255, 255, 255)
                run.font.size = Pt(font_size)
    for row_index, row in enumerate(rows):
        cells = table.add_row().cells
        for i, value in enumerate(row):
            cells[i].text = str(value)
            cells[i].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            set_cell_margins(cells[i])
            if row_index % 2 == 1:
                set_cell_shading(cells[i], "F2F6FB")
            for p in cells[i].paragraphs:
                p.paragraph_format.space_after = Pt(1)
                p.paragraph_format.space_before = Pt(0)
                for run in p.runs:
                    run.font.size = Pt(font_size)
                    run.font.color.rgb = RGBColor.from_string(INK)
    if widths:
        for row in table.rows:
            for i, width in enumerate(widths):
                row.cells[i].width = Inches(width)
    doc.add_paragraph().paragraph_format.space_after = Pt(1)
    return table


def add_body(doc, text, style="Normal"):
    p = doc.add_paragraph(style=style)
    p.paragraph_format.space_after = Pt(5)
    p.paragraph_format.line_spacing = 1.05
    p.add_run(text)
    return p


def add_bullets(doc, items):
    for item in items:
        p = doc.add_paragraph(style="List Bullet")
        p.paragraph_format.space_after = Pt(2)
        p.add_run(item)


def add_code(doc, code):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.18)
    p.paragraph_format.right_indent = Inches(0.18)
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(7)
    pPr = p._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), "F1F4F8")
    pPr.append(shd)
    r = p.add_run(code)
    r.font.name = "Consolas"
    r.font.size = Pt(8)
    return p


def heading(doc, title, level=1):
    p = doc.add_heading(title, level=level)
    p.paragraph_format.keep_with_next = True
    return p


def page_title(doc, title, kicker=None):
    if doc.paragraphs or len(doc.tables):
        doc.add_page_break()
    if kicker:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(3)
        r = p.add_run(kicker.upper())
        r.bold = True
        r.font.size = Pt(9)
        r.font.color.rgb = RGBColor.from_string(TEAL)
    heading(doc, title, 1)


def figure(doc, image_stream, caption, width=6.2):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run().add_picture(image_stream, width=Inches(width))
    cap = doc.add_paragraph(caption)
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap.paragraph_format.space_after = Pt(6)
    for run in cap.runs:
        run.italic = True
        run.font.size = Pt(8)
        run.font.color.rgb = RGBColor.from_string(GREY)


def prepare_doc():
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Inches(0.62)
    section.bottom_margin = Inches(0.62)
    section.left_margin = Inches(0.7)
    section.right_margin = Inches(0.7)
    normal = doc.styles["Normal"]
    normal.font.name = "Aptos"
    normal.font.size = Pt(9.5)
    normal.font.color.rgb = RGBColor.from_string(INK)
    for name, size, color in [("Heading 1", 17, NAVY), ("Heading 2", 12, BLUE), ("Heading 3", 10, TEAL)]:
        style = doc.styles[name]
        style.font.name = "Aptos Display"
        style.font.size = Pt(size)
        style.font.bold = True
        style.font.color.rgb = RGBColor.from_string(color)
        style.paragraph_format.space_before = Pt(7)
        style.paragraph_format.space_after = Pt(5)
    header = section.header.paragraphs[0]
    header.text = "CASE 70  |  VEHICLE SERVICE BOOKING  |  FINAL SUBMISSION"
    header.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    for run in header.runs:
        run.font.size = Pt(8)
        run.font.color.rgb = RGBColor.from_string(GREY)
    footer = section.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer.add_run("Suhas  •  ENG24CS1021  •  2026–2027     |     Page ")
    add_page_field(footer)
    return doc


def add_lab_artifact_details(doc, lab_title):
    if lab_title.startswith("L1.1"):
        heading(doc, "SDLC plan deliverable", 3)
        add_doc_table(doc, ["Phase", "Activity", "Output", "Owner (proposed)"], [
            ["Planning", "Confirm 40-workshop scope", "Approved charter", "Suhas"],
            ["Requirements", "Define booking/reminder needs", "FR-01 to FR-10; NFR-01 to NFR-05", "Product Owner"],
            ["Design", "Design Flask and SQLite flow", "Architecture and data model", "Developer/architect"],
            ["Implementation", "Build reminder scheduler", "Working increment", "Developer"],
            ["Testing", "Unit, browser and reliability checks", "Test report and defect log", "Tester"],
            ["Deployment", "Package CI/container configuration", "Release candidate", "DevOps role"],
            ["Maintenance", "Monitor health and reminders", "Incident/change backlog", "Service owner"],
        ], font_size=7.6)
    elif lab_title.startswith("L1.2"):
        heading(doc, "Weighted decision matrix", 3)
        add_body(doc, "Criteria weights: feedback speed 25; risk control 20; change handling 25; documentation 15; prototype fit 15. Each rating is 1–5. Weighted points are weight × rating/5; maximum total is 100.")
        add_doc_table(doc, ["Model", "Weighted points in criterion order", "Total / 100"], [
            ["Incremental", "25, 16, 25, 12, 15", "93"],
            ["Agile", "25, 12, 25, 9, 15", "86"],
            ["Spiral", "20, 20, 20, 15, 9", "84"],
            ["V-Model", "10, 20, 10, 15, 6", "61"],
            ["Waterfall", "5, 12, 5, 15, 3", "40"],
        ], font_size=7.8)
        add_body(doc, "Conclusion (five lines): Incremental ranks first at 93/100. It supports early booking/reminder delivery. Reviews let the team incorporate workshop feedback. Smaller releases help limit change risk. This supports the Incremental model fixed by Case 70.")
    elif lab_title.startswith("L1.3"):
        heading(doc, "Handwritten draft and reviewed improvement", 3)
        add_body(doc, "The submitted files are function fragments; they assume a date import and a reminders collection supplied elsewhere.")
        add_code(doc, "def schedule_reminder(booking):\n    if booking[\"date\"] >= date.today():\n        reminders.append({\"booking\": booking[\"id\"], \"status\": \"Scheduled\"})")
        add_code(doc, "def schedule_reminder(booking):\n    if not isinstance(booking.get(\"date\"), date):\n        raise ValueError(\"Invalid booking date\")\n    if booking[\"date\"] < date.today():\n        raise ValueError(\"Past booking date\")\n    if any(r[\"booking\"] == booking[\"id\"] and r[\"status\"] == \"Sent\" for r in reminders):\n        return False\n    reminders.append({\"booking\": booking[\"id\"], \"status\": \"Scheduled\"})\n    return True")
        add_doc_table(doc, ["Change", "Assessment"], [
            ["Date type guard", "Valid; prevents type-related comparison failures."],
            ["Past-date rejection", "Valid; aligns with future booking rule."],
            ["Existing Sent check", "Valid sequential safeguard; not safe against concurrent workers."],
            ["Boolean return", "Valid observability improvement; caller must use it."],
        ], font_size=7.8)
    elif lab_title.startswith("L1.4"):
        heading(doc, "Before → testable requirement", 3)
        rows = [
            ["Need to make a booking", "FR-01: For a valid booking, create exactly one linked reminder record."],
            ["Select a vehicle and workshop", "FR-02: /book accepts an existing vehicle, workshop, date and service type."],
            ["Do not allow past dates", "FR-03: Reject a date earlier than server-local today without creating a booking."],
            ["Do not send twice", "FR-04: A repeat scheduler call does not process a booking with a Sent reminder again."],
            ["Show low weekday use", "FR-05: Dashboard returns scheduled-booking counts for Monday–Sunday."],
            ["Respect marketing opt-out", "FR-06: Store opt-out and ensure promotional sending checks it before dispatch."],
            ["Support all locations", "FR-07: Seeded demo exposes 40 distinct workshops."],
            ["Show service health", "FR-08: /health returns 200 while database query succeeds."],
            ["Update reminder outcome", "FR-09: Successful processing changes reminder status and records sent_at."],
            ["Find reminder by state", "FR-10: Reminder page can filter rows by selected status."],
            ["Be responsive", "NFR-01: 95% of 50 health checks under 500 ms in a defined local test."],
            ["Work in a browser", "NFR-02: Core flow works in a supported desktop browser."],
            ["Keep records", "NFR-03: Bookings survive process restart using persistent database storage."],
            ["Protect consent information", "NFR-04: Do not write customer consent values to application logs."],
            ["Automate browser tests", "NFR-05: Selenium flow runs without manual database edits."],
        ]
        add_doc_table(doc, ["Initial scenario need", "Rewritten measurable requirement"], rows, font_size=7.2)
    elif lab_title.startswith("L1.5"):
        heading(doc, "Test candidate review and accuracy", 3)
        add_doc_table(doc, ["Test idea", "Origin", "Usable as drafted?", "Follow-up"], [
            ["Future booking creates reminder", "AI", "Yes", "Assert one linked Scheduled row."],
            ["Second scheduler run is idempotent", "AI", "Yes", "Assert processed count is zero on second run."],
            ["Past-date booking rejected", "AI", "Yes", "Assert error and no booking persisted."],
            ["Blank vehicle rejected", "AI", "No", "Expected response/validation rule was missing."],
            ["Unknown workshop handled", "AI", "Yes", "Specify safe validation response before release."],
            ["Duplicate Sent reminder", "Human", "Yes", "Boundary/idempotency case."],
            ["Consent false", "Human", "Yes", "Negative promotional-message case."],
            ["Invalid date format", "Human", "Yes", "Negative parsing case."],
            ["Empty scheduler queue", "Human", "Yes", "Boundary case; expect zero processed."],
            ["Concurrent scheduler invocation", "Human", "Yes", "Negative concurrency case; needs stronger persistence guard."],
        ], font_size=7.2)
        add_body(doc, "AI suggestion usefulness = 4 directly usable candidates / 5 AI candidates = 80%. This is a human review metric, not model accuracy on a benchmark or application pass rate.")
    elif lab_title.startswith("L2.2"):
        add_doc_table(doc, ["Scrum element", "Setup for this assignment"], [
            ["Product Owner", "Suhas (student prototype owner)"],
            ["Scrum Master", "Suhas (facilitator role; individual project)"],
            ["Developer", "Suhas (implementation and tests)"],
            ["Stakeholder reviewer", "Dealership service manager role (simulated)"],
            ["Events", "Planning, daily check-in, review, retrospective"],
            ["Artifacts", "Product backlog, sprint backlog, increment, Definition of Done"],
        ], font_size=7.5)
    elif lab_title.startswith("L2.3"):
        add_doc_table(doc, ["Story", "Acceptance criterion"], [
            ["US-01 Search workshops", "Customer can view the seeded 40 workshop list."],
            ["US-02 Create booking", "A valid booking is persisted and assigned an ID."],
            ["US-03 Validate date", "Past-date booking is rejected without persistence."],
            ["US-04 Schedule reminder", "New booking links one reminder; second run does not duplicate Sent state."],
            ["US-05 Opt out", "Consent choice persists; promotional sender must honour false."],
        ], font_size=7.5)
    elif lab_title.startswith("L2.5"):
        add_doc_table(doc, ["Ticket group", "Tickets", "AI duplicate interpretation"], [
            ["Reminder duplication", "T1, T6", "Duplicates"],
            ["Date validation", "T2, T8", "Duplicates"],
            ["Consent opt-out", "T3, T10", "Duplicates"],
            ["Workshop capacity", "T4", "Unique"],
            ["Scheduler state", "T5", "Unique"],
            ["Registration validation", "T7", "Unique"],
            ["Utilisation chart", "T9", "Unique"],
        ], font_size=7.5)
        add_body(doc, "Human-reviewed classification: 10/10 ticket decisions correct for this constructed set. The test set is small and not an independent AI-accuracy benchmark.")
    elif lab_title.startswith("L3.1"):
        add_doc_table(doc, ["Style", "Prompt used", "Expected advantage"], [
            ["Role", "You are a senior Flask developer. Implement duplicate-safe reminder scheduling.", "Sets perspective and code context."],
            ["Few-shot", "Example 1: a future booking creates Scheduled reminder. Example 2: a Sent reminder is not sent twice.", "Makes input/output examples explicit."],
            ["Step-by-step", "1 parse booking; 2 validate date; 3 check consent; 4 check Sent record; 5 update status.", "Makes logic sequence reviewable."],
        ], font_size=7.3)
    elif lab_title.startswith("L3.2"):
        add_doc_table(doc, ["Suggestion", "Decision", "Reason"], [
            ["Use SQLite Sent record check", "Accept with caveat", "Prevents simple sequential duplicate; concurrency guard still needed."],
            ["Automatically send email", "Reject", "No approved external provider or delivery confirmation."],
            ["Catch every exception silently", "Reject", "Hides failures and corrupts evidence."],
            ["Use synthetic customer data", "Accept", "Reduces exposure risk."],
        ], font_size=7.5)
    elif lab_title.startswith("L4.1"):
        add_doc_table(doc, ["Area", "Exploration / finding", "Qualification"], [
            ["Date validation", "Past-date path is covered by a test.", "Manual 20-minute timing is not independently demonstrated."],
            ["Duplicate run", "Sequential duplicate reminder test exists.", "Concurrency not tested."],
            ["Consent", "Consent endpoint can update preference.", "Scheduler enforcement missing."],
            ["Health", "/health returns healthy status locally.", "Local smoke test only."],
        ], font_size=7.4)
    elif lab_title.startswith("L4.3"):
        add_doc_table(doc, ["Repeat", "Manual seconds*", "Selenium seconds*"], [
            ["1", "18", "2.1"], ["2", "19", "1.9"], ["3", "17", "2.0"],
            ["4", "20", "2.2"], ["5", "18", "1.8"], ["Mean", "18.4", "2.0"],
        ], font_size=7.5)
        add_body(doc, "*Values in the supplied timing artifact are illustrative assumptions, not timer-captured results. Do not claim a measured 9.2× improvement or break-even.")
    elif lab_title.startswith("L4.4"):
        add_doc_table(doc, ["Rank", "Test", "Likelihood", "Impact", "Score"], [
            ["1", "Past-date booking", "5", "5", "25"],
            ["2", "Duplicate reminder", "5", "5", "25"],
            ["3", "Consent opt-out", "5", "4", "20"],
            ["4", "Database unavailable", "4", "5", "20"],
            ["5", "Invalid registration", "4", "3", "12"],
            ["6", "Workshop capacity", "3", "4", "12"],
            ["7", "Utilisation chart", "3", "2", "6"],
            ["8", "Marketing template", "2", "3", "6"],
            ["9", "Navigation", "2", "2", "4"],
            ["10", "Footer text", "1", "1", "1"],
        ], font_size=7.3)
    elif lab_title.startswith("L5.3"):
        add_doc_table(doc, ["Review check", "Applied finding"], [
            ["Requirement traceability", "Duplicate prevention maps to FR-04."],
            ["Input validation", "Past dates are rejected; malformed IDs/dates need guards."],
            ["Database safety", "ORM queries used; concurrency uniqueness still required."],
            ["Consent and privacy", "Consent API exists; scheduler does not enforce it."],
            ["Exceptions / status", "No silent catch; Sent state is not proof of delivery."],
            ["Tests", "Sequential repeat-run test exists; add concurrent and invalid-input cases."],
        ], font_size=7.3)
    elif lab_title.startswith("L5.5"):
        add_doc_table(doc, ["Synthetic input name", "Region", "Gender", "Observed status"], [
            ["Aarav", "Karnataka", "Male", "Scheduled"],
            ["Mary", "Tamil Nadu", "Female", "Scheduled"],
            ["Lin", "Maharashtra", "Non-binary", "Scheduled"],
        ], font_size=7.5)
        add_body(doc, "The three examples show identical status, but are too few to establish fairness. The mitigation is to avoid using protected attributes for reminder eligibility, compare matched synthetic test pairs, and review exceptions with a human.")


def create_report():
    doc = prepare_doc()

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(78)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("CASE STUDY ASSIGNMENT")
    r.bold = True
    r.font.size = Pt(14)
    r.font.color.rgb = RGBColor.from_string(TEAL)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(12)
    r = p.add_run("VEHICLE SERVICE\nBOOKING")
    r.bold = True
    r.font.size = Pt(32)
    r.font.color.rgb = RGBColor.from_string(NAVY)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("Incremental Development • Scrum • Reminder Scheduler")
    r.font.size = Pt(14)
    r.font.color.rgb = RGBColor.from_string(BLUE)
    doc.add_paragraph()
    add_doc_table(doc, ["STUDENT DETAILS", "CASE DETAILS"], [
        ["Name: Suhas\nUSN: ENG24CS1021\nRoll No.: 57\nSection: 5B\nDepartment: Computer Science and Engineering",
         "Case 70: Vehicle Service Booking\nAcademic Year: 2026–2027\nProcess model: Incremental\nAgile method: Scrum\nAI tool named in assignment: ChatGPT"],
    ], font_size=10)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(15)
    p.add_run("Repository: https://github.com/productivity1947-ux/case70-vehicle-service-booking").font.size = Pt(9)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run("Submission date: 26 November 2026").font.color.rgb = RGBColor.from_string(GREY)

    page_title(doc, "Executive Summary", "Project overview")
    add_body(doc, "This report follows the Case 70 assignment sheet for Vehicle Service Booking: a dealership chain with 40 workshops wants online service bookings and reminders, while weekday slots are underused. The assigned process model is Incremental, the Agile method is Scrum, and the prototype focus is a reminder scheduler.")
    add_body(doc, "The delivered local prototype is a Flask application backed by SQLite. Its seeded demonstration data contains 40 workshops, 30 synthetic customers, 30 vehicles, 30 bookings and 30 reminder records. Implemented flows include the dashboard, booking page, reminder list, reminder processing endpoint, marketing-consent endpoint and a health endpoint. These are prototype capabilities, not a production deployment.")
    add_doc_table(doc, ["Evidence item", "Observed result", "Interpretation"], [
        ["Automated test suite", "8 tests passed", "PyTest run against Flask application"],
        ["Statement coverage", "72% (142 statements; 40 missed)", "Measured coverage.py result; below proposed 90% gate"],
        ["Health monitoring", "10/10 HTTP 200; responses 9.21–37.90 ms", "One actual 10-minute run; no 500 ms alert"],
        ["Selenium", "Dashboard, booking page, reminder page and scheduler interaction", "The assignment's 10-case × 5-repeat benchmark was not completed"],
        ["Jenkins / Docker / Sonar", "Configuration or environment only", "No live Jenkins build, Docker image or Sonar scan is claimed"],
    ], font_size=8.5)
    add_body(doc, "The report distinguishes implemented behaviour, measured results, planning assumptions and incomplete external-tool exercises. In particular, marketing preference is stored through an API but is not yet enforced by reminder processing; reminders are marked Sent in the database without a connected email provider.")

    page_title(doc, "Case, Scope and Identifiers", "Project baseline")
    add_doc_table(doc, ["Fixed case item", "Value used in this submission"], [
        ["Business scenario", "40-workshop dealership chain; online booking and reminders; weekday slot under-utilisation"],
        ["Prototype module", "Reminder scheduler, with a small booking workflow for end-to-end context"],
        ["Process model", "Incremental (assignment-prescribed)"],
        ["Agile method", "Scrum (assignment-prescribed)"],
        ["AI tool", "ChatGPT named by the assignment; synthetic examples only"],
        ["Requirement IDs", "FR-01 through FR-10 and NFR-01 through NFR-05"],
        ["Story / test IDs", "US-01 through US-05; TC-01 and related test identifiers"],
    ], font_size=9)
    figure(doc, make_figure("architecture"), "Figure 1. Prototype architecture based on the implemented Flask routes and SQLite models.")
    add_body(doc, "System boundaries: no real email/SMS delivery provider, authentication/role system, production monitoring stack, real customer data or multi-workshop deployment is included. The workshop count is seeded locally for demonstration and does not prove capacity planning or weekday utilisation improvement in a live dealership.")

    page_title(doc, "Part 1 — Task A: Incremental Lifecycle and AI Prioritisation", "Part 1 • Case tasks")
    add_body(doc, "The assignment asks AI to prioritise features. The following is a reasoned draft prioritisation; a product owner must validate the assumptions with workshop staff and customers before implementation. AI can rank the backlog, but it cannot establish that a feature is commercially or legally required.")
    add_doc_table(doc, ["Priority", "Feature / ID", "Why it matters", "Human checkpoint"], [
        ["Must", "FR-02 online service booking", "Creates the core customer path and captures workshop/date/service selection.", "Confirm allowed dates, slot capacity and service categories."],
        ["Must", "FR-01 reminder scheduling", "The case specifically names reminders and weekday utilisation as the problem.", "Confirm timing, channel, consent and cancellation behaviour."],
        ["Must", "FR-03 future-date validation", "Prevents invalid bookings and underpins reminders.", "Agree timezone and whether same-day appointments are permitted."],
        ["Should", "FR-04 consent management", "Marketing messages require preference control; operational reminders must be distinguished.", "Legal/business owner approves purpose and message classification."],
        ["Should", "FR-05 weekday utilisation view", "Provides a signal for later slot-promotion decisions.", "Define denominator, location filters and business-day rules."],
    ], font_size=8)
    figure(doc, make_figure("incremental"), "Figure 2. A staged release path; each increment is reviewed before adding scope.")
    add_body(doc, "The prototype implements the booking and reminder flow, a weekday count view and a consent endpoint. The prioritisation itself is a proposal, not measured evidence that the solution has improved weekday fill. Pilot evaluation should compare booked/available weekday slots by workshop over a baseline period and use a controlled rollout.")

    page_title(doc, "Part 1 — Task B: Scrum and AI Estimation", "Part 1 • Case tasks")
    add_body(doc, "For a student-sized prototype, roles are proposed for the assignment exercise rather than claimed as appointments at a real dealership. Suhas acts as Product Owner for the case backlog and also implements the prototype; a facilitator can take the Scrum Master responsibilities; the dealership service manager is the proposed stakeholder reviewer.")
    add_doc_table(doc, ["Role / event", "Proposed assignment setup"], [
        ["Product Owner", "Suhas — prioritises FR-01…FR-10 and accepts each increment."],
        ["Scrum Master", "Suhas as facilitator for this individual prototype; removes blockers and keeps events time-boxed."],
        ["Developers", "Suhas — Flask/SQLite implementation, testing, evidence and documentation."],
        ["Stakeholder reviewer", "Dealership service manager role (simulated; no real person is represented)."],
        ["Sprint cadence", "Two-week planning assumption; daily check-in, review and retrospective."],
        ["Artifacts", "Product backlog, sprint backlog, increment, Definition of Done and a simple To Do / Doing / Review / Done board."],
    ], font_size=8.5)
    add_doc_table(doc, ["Story", "Planning-poker points", "AI estimate", "Decision / note"], [
        ["US-01 Workshop and vehicle selection", "5", "5", "Accept as initial estimate; confirm data quality and search needs."],
        ["US-02 Create future-dated booking", "5", "8", "Use 5 provisionally; validate form, persistence and date tests."],
        ["US-03 Schedule one reminder", "8", "8", "Highest technical risk; idempotency and scheduling rules."],
        ["US-04 Run reminder and show status", "5", "5", "Estimate excludes real email provider integration."],
        ["US-05 Store marketing opt-out", "3", "5", "Consent persistence is small; enforcement across channels is not."],
    ], font_size=8)
    add_body(doc, "Estimates are illustrative planning-poker and AI-assisted estimates, not historical measurements. AI should receive synthetic story text, prior anonymised velocity and dependency information only. Estimate discussion must include test effort, external integrations, unknowns, staffing and stakeholder availability. Velocity is a planning aid, not a performance target.")

    page_title(doc, "Part 1 — Task C: Prompting, Git, Jenkins and Docker", "Part 1 • Case tasks")
    heading(doc, "Prompt refinement for the reminder scheduler", 2)
    add_doc_table(doc, ["Version", "Prompt", "Review"], [
        ["V1 — weak", "Write a reminder scheduler.", "Missing stack, data shape, date rules, duplicate handling, consent and test expectations."],
        ["V2 — contextual", "For Vehicle Service Booking, write a Flask/SQLAlchemy function that processes scheduled reminders for a booking and returns counts.", "Adds context and output; still omits idempotency, privacy, invalid inputs and tests."],
        ["V3 — constrained", "Given synthetic Flask/SQLAlchemy Booking and Reminder models, process only due Scheduled reminders, avoid a second Sent reminder for the same booking, set sent_at once, respect the approved consent rule, use a transaction, and provide PyTest cases for repeat-run, future-dated and opted-out examples. Do not send real messages or suppress exceptions.", "Specific, testable, bounded; still requires code review and a consent requirement decision."],
    ], font_size=8)
    add_body(doc, "GitHub Flow is appropriate for a small prototype: short-lived feature branch, focused commit, pull request, test and review before merge. The GitHub repository is private on the master branch. Branch protection was attempted but GitHub returned 403 because required protection for a private repository needs an eligible plan. Therefore no blocked-PR screenshot is claimed.")
    add_body(doc, "A Jenkinsfile exists with checkout, install and test stages, but no Jenkins server/job/webhook or two build results were available. A push-trigger is not configured merely by having a Jenkinsfile. Dockerfile and Compose files exist; `docker version` showed the CLI but could not connect to the Docker Desktop Linux engine. No image build/container run is claimed.")
    add_code(doc, "FROM python:3.14-slim   # pinned family of Python runtime image\nWORKDIR /app            # set container working directory\nCOPY requirements.txt . # copy dependency manifest first\nRUN pip install --no-cache-dir -r requirements.txt\nCOPY . .                # copy application and templates\nEXPOSE 5000             # document application port\nCMD [\"python\", \"app.py\"] # start Flask application")

    page_title(doc, "Part 1 — Task D: Selenium Testing and Monitoring", "Part 1 • Case tasks")
    add_body(doc, "The Selenium script opens the dashboard, opens the booking form, navigates to the reminder page and clicks the Run scheduler button. It also saves browser screenshots. It does not exercise a complete booking/reschedule/cancel sequence, and the application currently has no reschedule/cancel workflow.")
    add_doc_table(doc, ["Monitoring sample", "Time (ms)", "HTTP", "Alert (>500 ms)"], [
        [str(i + 1), row["response_ms"], row["status_code"], row["alert"]]
        for i, row in enumerate(list(csv.DictReader((PROJECT / "monitor_log.csv").open(encoding="utf-8-sig")))[-10:])
    ], font_size=8)
    figure(doc, make_figure("monitor"), "Figure 3. Actual last 10 rows of monitor_log.csv. Alert threshold is 500 ms.")
    add_body(doc, "Measured run: 10 scheduled samples, all HTTP 200; minimum 9.21 ms, maximum 37.90 ms; no threshold alerts. This is a local health endpoint result, not a load test or availability guarantee. The assignment's separate manual-versus-Selenium requirement (same 10 cases, five repetitions each) is not fulfilled by the available timing table: those numbers are assumed/illustrative and must not be interpreted as measured benchmark data.")
    add_body(doc, "Priority test cases include valid future booking, date boundary (today), past date rejection, malformed date, invalid vehicle/workshop IDs, duplicate scheduler run, opt-out behaviour, database failure and weekday aggregation. TC-01 is the repeat-run/idempotency test. The actual PyTest suite contains eight tests; the Selenium script is a limited browser smoke test.")

    page_title(doc, "Part 1 — Task E: Quality, Usability and Ethics", "Part 1 • Case tasks")
    add_doc_table(doc, ["Quality / ethical dimension", "Proposed measure or control", "Observed status"], [
        ["Correctness", "Critical requirements mapped to executable tests; no unresolved critical defects.", "8 tests passed; not a full acceptance suite."],
        ["Reliability", "50 repeated scheduler runs; record failure rate and define target.", "50-run scheduler experiment not performed; no measured claim."],
        ["Maintainability", "Reviewability, clear functions, lint and complexity scan.", "SonarQube unavailable; no scan results."],
        ["Usability", "At least 90% of a small representative first-time-user sample completes booking without help.", "Target only; no usability study performed."],
        ["Performance", "95th percentile health response under 500 ms over 50 checks.", "One ten-sample local run below 500 ms; insufficient for the 50-check target."],
        ["Consent / privacy", "Separate operational service reminders from promotional marketing; enforce opt-out before sending.", "Consent flag API exists; scheduler does not yet enforce it; no provider sends mail."],
    ], font_size=8)
    add_body(doc, "The prototype marks reminders Sent in SQLite but does not deliver messages. That status should be renamed to Processed or the delivery provider and delivery confirmation should be added before production. Missing consent enforcement, absent authentication/authorisation, malformed input handling and lack of an audit trail are release blockers.")
    add_body(doc, "SonarQube was not run because no scanner/server was available. A quality gate is proposed, not reported as active: tests pass; coverage at least the current 72% baseline with a ratcheting improvement target; no new critical/high security issue; secrets scan clean; approved pull request. Never paste names, phone numbers, vehicle registrations, tokens, database files or meeting recordings into unapproved AI services.")

    theory = [
        ("T1.1 — Software engineering and SDLC",
         "Software engineering applies requirements, design, implementation, verification, release and maintenance disciplines to a system whose users and data are distributed across 40 workshops. Ad-hoc code would make booking rules, reminders, consent and records inconsistent. The SDLC outputs are: plan/scope; approved requirements; architecture/data model; reviewed code; test report; release/configuration; and monitored change backlog. The prototype is an increment, not a complete dealership production system."),
        ("T1.2 — Process-model fit",
         "For this case, a reasoned fit ranking is Incremental, Agile, Spiral, V-Model, Waterfall. Incremental is prescribed and fits staged value: booking first, then reminder idempotency, consent enforcement and analytics. Agile supports feedback but is a method family rather than a strict lifecycle model in all comparisons. Spiral is useful if integration/security risks dominate but has overhead. V-Model adds strong verification traceability, while pure Waterfall delays feedback despite operational variation among 40 workshops. Risk and regulator constraints could change the ranking."),
        ("T1.3 — AI-driven versus traditional lifecycle",
         "AI can assist each conventional phase, but does not remove its owner. Planning: summarise goals (human approves scope); requirements: draft stories (stakeholder accepts measurable criteria); design: suggest schema/flows (architect checks privacy and consistency); implementation: code completion (developer tests/reviews); testing: propose cases (tester checks expected outcomes); deployment: draft configuration (operator validates secrets and rollback); maintenance: cluster incidents (service owner confirms root cause). AI output must remain traceable to approved FR/US/TC IDs."),
        ("T1.4 — Copilot, Tabnine and ChatGPT",
         "Copilot primarily offers code completion and chat in an IDE; Tabnine provides IDE-integrated completion and team-oriented controls depending on plan; ChatGPT is conversational and can reason over supplied prompts/files. Their data collection, retention, training use and enterprise controls depend on current product settings and contract; do not assume privacy from the product name. Ask: may submitted code or customer data be retained or used for training, where is it processed, and can the organisation disable retention? Use synthetic examples and approved accounts."),
        ("T1.5 — Limits of AI requirements analysis",
         "AI can turn interview notes into candidate requirements but may invent a default, hide conflict, or treat a phrase like “weekday slots go unused” as a proven demand pattern. The engineer should preserve source notes, tag each statement as confirmed/assumption/open, ask the relevant service manager and customer representative, define measurable acceptance criteria, and obtain sign-off. Conflicts about date cut-offs, reminder channel, promotional consent and workshop capacity cannot be resolved by a language model; escalate them and record the decision."),
        ("T2.1 — Agile values and principles",
         "The four values applied to this case are: individuals/interactions over process/tools (short workshop feedback); working software over exhaustive documentation (demonstrate booking and scheduler increments while retaining essential privacy and decision records); customer collaboration over contract negotiation (service staff review acceptance criteria); and responding to change over following a fixed plan (reorder backlog when a workshop identifies a real scheduling constraint). Six relevant principles: deliver valuable software early; welcome changing requirements; deliver frequently; business and developers collaborate; working software is the primary progress measure; and regularly reflect/adapt. Agile still needs tests, documentation and safety controls."),
        ("T2.2 — Scrum and its limitation",
         "Scrum defines Product Owner, Scrum Master and Developers; a Product Backlog, Sprint Backlog and Increment; and Sprint Planning, Daily Scrum, Sprint Review and Sprint Retrospective. Refinement is ongoing. For this student prototype, roles are simulated and may overlap; this is not evidence that a 40-workshop organisation has adopted Scrum. A limitation is dependency on stakeholder availability: if workshop representatives cannot attend reviews, the team may build a technically sound but operationally unsuitable flow. Handle this with a named proxy and recorded acceptance decisions."),
        ("T2.3 — Scrum artifacts and Definition of Done",
         "The Product Backlog is the ordered list of outcomes such as online booking, reminders, consent and weekday reporting. The Sprint Backlog is the chosen subset plus the plan for the current sprint (for example US-02 booking and TC-01 idempotency). The Increment is the integrated, usable, tested application change. Definition of Done should require code review, automated tests, safe synthetic data, acceptance criteria met, no unresolved high-severity issue, documentation updated and rollback understood. The prototype increment is not a production-ready release."),
        ("T2.4 — AI for planning and estimation",
         "AI can draft stories from scenario statements, identify missing edge cases and provide a rough relative estimate. It needs anonymised past story sizes, actual completed velocity, skill availability, dependencies, test effort, integration constraints, workshop blackout periods and uncertainty. This assignment has no real three-sprint history; the 24/27/26-point values are explicitly assumed. AI estimates should be shown beside human planning-poker estimates and accepted only after discussing why they differ."),
        ("T2.5 — Governance for AI in planning",
         "Use only approved accounts and synthetic or minimised meeting text. Do not upload contact details, vehicle registrations, customer complaints or confidential commercial plans. Tell participants when an AI assistant is recording/summarising; obtain consent; define access and deletion periods; label generated suggestions; retain decisions and accountable owners; let people correct the minutes; prohibit automated performance scoring; and review data-retention terms. A human Product Owner remains accountable for the backlog and estimates."),
        ("T3.1 — AI assistance across development",
         "Requirements: AI drafts candidate FRs; product owner verifies each. Design: AI suggests routes/schema; developer checks relationships and privacy. Code: IDE assistant proposes code; developer runs tests and checks licensing/security. Tests: AI drafts positive/negative cases; tester verifies expected outcomes and non-duplication. CI: AI may summarise failures; engineer reads the actual logs. Operations: AI can group alerts; on-call owner validates the incident and any action. AI is advisory and does not approve or deploy autonomously."),
        ("T3.2 — Prompt engineering",
         "A good prompt states role/context, task, input/output shape, constraints, examples and verification. Weak: “Make a reminder scheduler.” Strong: “For the synthetic Flask/SQLAlchemy models in this repository, propose a function that processes due Scheduled reminders once per booking, records a timestamp, handles an opted-out marketing contact according to the confirmed policy, returns processed/skipped counts, raises actionable errors, and includes PyTest cases for repeat invocation and date boundaries. Do not send messages or invent missing policy.” The strong version still needs a human-approved consent rule."),
        ("T3.3 — Git workflows",
         "GitHub Flow uses short feature branches and pull requests; GitFlow adds develop/release/hotfix branches; trunk-based development keeps branches very short and relies on frequent integration and strong automation. GitHub Flow is a simple fit for a small prototype if review and CI are available. In this repository, the code is pushed to a private GitHub repository, but protection could not be enabled on the available plan; this is a known control gap rather than a completed branch-protection exercise."),
        ("T3.4 — CI/CD risk reduction",
         "A pipeline can install pinned dependencies, run tests, report coverage, scan for secrets, build an image and publish a versioned artifact only after passing gates. This makes regressions visible and provides repeatable packaging. The repository has a Jenkinsfile with checkout/install/test stages, but a Jenkins job, push webhook and two build results were not produced. CI cannot prove the business rules are correct, prevent every supply-chain issue, or guarantee an external service is available."),
        ("T3.5 — Containers",
         "A Docker image is a built, immutable filesystem/configuration package; a container is a running instance; a Dockerfile expresses image-building steps; Docker Compose describes related services and local runtime settings. They can make Flask and dependencies more repeatable, but secrets must be injected securely and persistent data needs a volume. The Docker CLI was present, but its Linux engine pipe was unavailable, so no successful build or container run is claimed."),
        ("T4.1 — Verification, validation, bug and failure",
         "Verification asks whether the code conforms to requirements (for example, does a past-date request get rejected?). Validation asks whether the resulting workflow solves a customer/workshop need (can a customer complete an appropriate booking?). A bug is a defect in code or design; a failure is observable incorrect behaviour during execution. A date-parsing defect is a bug; accepting yesterday’s booking is a failure. A passing unit test does not alone validate the complete multi-workshop process."),
        ("T4.2 — Integration and system defects",
         "Integration defects include Flask form IDs not matching database foreign keys, transaction failures while creating a booking and reminder, or a scheduler status not appearing on the reminder page. System defects include consent not being checked before a real marketing sender dispatches a message, timezone mismatches at workshop boundaries, or concurrent jobs sending duplicates. The present prototype has no external message sender and its tests do not establish those integration behaviours; they remain release tests to add."),
        ("T4.3 — Selenium",
         "Selenium WebDriver drives a real browser. Locators identify elements by stable ID/name/CSS or accessible text; explicit waits handle asynchronous rendering; assertions check the expected page/state. First automate the booking-to-reminder path because it crosses UI, route, database and scheduler boundaries. The supplied script is a browser smoke run (dashboard, booking page, reminders and scheduler click), not the full assignment flow of booking, rescheduling and cancelling. The application does not implement reschedule/cancel yet."),
        ("T4.4 — AI testing and monitoring",
         "AI assistants can draft tests from requirements and help summarise logs; browser automation frameworks and observability tools provide the actual execution and metrics. Risks of unreviewed generated tests include wrong expected results, duplicated cases, missing boundaries, tests that pass without asserting behaviour, privacy leaks and brittle locators. Keep generated candidates traceable to FR/TC IDs, review them, run them and inspect failures. A written model response is not a test result."),
        ("T4.5 — Benefits and challenges of AI testing",
         "AI can accelerate test idea generation, explain stack traces and identify likely boundary conditions. It may hallucinate APIs, overfit to the prompt, create false positives or miss business-specific edge cases. Maintain trust through deterministic test data, clear assertions, reproducible commands, human review and a record of false alarms. For this case, prioritise date validation, duplicate safety, consent enforcement and weekday aggregation, then compare model suggestions with actual code and stakeholder-approved rules."),
        ("T5.1 — SQA and quality attributes",
         "Software Quality Assurance plans and evaluates processes so the product meets specified needs. Correctness: all critical booking/reminder acceptance tests pass. Reliability: target at least 98% successful scheduler invocations over 50 defined runs (not measured here). Maintainability: reviewable small functions, tests and no unresolved critical static-analysis issue. Usability: proposed target 90% of representative first-time users complete booking unaided (not tested). Efficiency: proposed 95th-percentile local health response under 500 ms on a defined 50-check run; current 10-sample run is only preliminary."),
        ("T5.2 — Quality gate",
         "A merge gate should run lint, unit/integration tests, coverage, dependency and secret checks. Start with the measured 72% coverage baseline and require no regression while a plan raises the threshold toward 90%; require all critical tests pass and zero exposed secrets; fail on new high-severity static-analysis findings; and require one independent review. The Jenkinsfile does not currently enforce these gates and no Sonar report exists, so this is a proposed control."),
        ("T5.3 — Safe AI code-review checklist",
         "For each AI suggestion: map it to an approved requirement; inspect input validation and authentication; verify parameterised database access; check consent/data minimisation; handle errors explicitly; consider duplicate/concurrent execution; review dependencies/licence; check secrets and logs; add positive, negative and boundary tests; run the code; and document accept/reject rationale. Apply the checklist to the scheduler’s duplicate check and the unimplemented consent enforcement; reject any claim that a database status update proves message delivery."),
        ("T5.4 — Responsible AI",
         "Fairness: names, gender and region must not change service eligibility or reminder status. Transparency: identify AI-assisted drafts and known limits. Accountability: a named product owner approves requirements and human reviewers approve code. Privacy: use synthetic customer/contact values and do not send production records to a model. Safety: validate inputs, restrict privileges, test failure and provide manual support. The local fairness table is only a small synthetic example, not a statistically valid bias audit."),
        ("T5.5 — Protected information",
         "Never submit passwords, access tokens, private keys, session cookies, production database exports, customer names/contact details, vehicle registration numbers tied to people, payment data, confidential workshop terms or identifiable meeting recordings to an unapproved AI tool. Replace them with synthetic IDs and fictional examples, remove metadata, use approved enterprise settings, restrict prompt access, and check retention rules. The prototype uses synthetic seed records; this does not make production data safe to upload."),
    ]

    for module in range(1, 6):
        page_title(doc, f"Part 2 — Module {module}: Theory Answers", "Part 2 • Theory questions")
        subset = theory[(module - 1) * 5:module * 5]
        for question, answer in subset:
            heading(doc, question, 2)
            add_body(doc, answer)

    labs_by_module = {
        1: [
            ("L1.1 — SDLC plan", "The plan moves through planning, requirements, design, implementation, testing, deployment and maintenance. Owners are proposed as Suhas (case/product owner), developer, test reviewer and operations owner; the small student project may have role overlap. Outputs include scope, approved FRs, architecture, code, test evidence, release package and monitoring backlog.", "Evidence: Case70_Lab_Outputs/L1.1_SDLC/L1.1_SDLC_Plan.csv."),
            ("L1.2 — Weighted process-model matrix", "Five criteria are scored 1–5 and weighted to 100: feedback speed 25, risk control 20, change handling 25, documentation 15 and prototype fit 15. Weighted contribution = criterion weight × rating/5. Totals: Incremental 93, Agile 86, Spiral 84, V-Model 61, Waterfall 40. These are transparent project judgements, not empirical measurements.", "Five-line conclusion: (1) Incremental scores highest at 93/100. (2) It releases a small booking/reminder slice early. (3) Reviews can adapt to workshop feedback. (4) Risk is limited by validating each increment. (5) It is consistent with the fixed model assigned in Case 70."),
            ("L1.3 — Handwritten and improved scheduler", "The handwritten draft checks whether a booking is eligible and appends a Scheduled record. The reviewed improvement validates the date type, rejects past dates, checks an existing Sent record and returns a Boolean outcome. A Boolean cannot replace transaction handling or idempotency under concurrent requests. The production Flask route also checks existing Sent reminders; the prototype does not deliver email.", "Change log: date type guard—valid; past-date guard—valid; duplicate Sent guard—valid but concurrency remains; observable return—valid; avoid claim of delivery—required scope correction."),
            ("L1.4 — Requirement analysis", "The final table contains ten functional and five non-functional testable statements tied to the scenario. Example: “The system shall reject a booking date earlier than the server’s local date and display an error without creating a booking.” A draft such as “weekday slots go unused” is not a requirement until the denominator, time window and target are agreed. FR-10 (status filtering) is not implemented in the current reminder view.", "Evidence: L1.4_Testable_Requirements.csv. The list is a requirements draft; stakeholder sign-off remains outstanding."),
            ("L1.5 — Generated test ideas and accuracy", "Five AI-origin candidates were assessed. Four were directly usable; the blank-vehicle suggestion lacked a precise expected response/validation rule. Candidate acceptance rate = 4/5 × 100 = 80%. This is review usefulness, not the software pass rate. Human additions cover boundary date (today), malformed date, repeated Sent reminder, consent false, empty queue and concurrent processing. Unknown IDs and concurrency are not fully handled by the current prototype.", "Evidence: L1.5_AI_Test_Accuracy.csv; test code: Case70_Vehicle_Service_Booking/tests/test_app.py."),
        ],
        2: [
            ("L2.1 — Agile working agreement", "Communication: written decisions in the repository; blockers raised at the daily check-in. Meetings: planning/refinement, 15-minute daily check, review and retrospective each two-week sprint. Definition of Done: acceptance criteria pass, tests pass, code reviewed, no secrets, evidence captured and README updated. Change handling: Product Owner reprioritises, team estimates impact, and changes enter the backlog rather than silently expanding the sprint.", "The existing agreement is a working proposal; ceremonies were not represented as meetings with a real team."),
            ("L2.2 — Scrum framework", "Proposed role names: Product Owner—Suhas; Scrum Master—Suhas (facilitator for solo project); Developer—Suhas; Stakeholder reviewer—dealership service manager role, simulated. Events: planning, daily check-in, review and retrospective. Artifacts: product backlog, sprint backlog, increment and Definition of Done. Role overlap is disclosed because this is an individual student prototype.", "Evidence: L2.2_Scrum_Framework.txt."),
            ("L2.3 — Epic into five stories", "Epic: Complete a vehicle booking and reminder flow. US-01 search 40 workshops (acceptance: list returns workshops); US-02 select vehicle and create booking (booking ID persists); US-03 validate future date (past date rejected without record); US-04 schedule one reminder (exactly one Scheduled row per new booking); US-05 opt out of marketing (preference persists and promotional sender must honour it). The last enforcement criterion remains a gap in code.", "Evidence: L2.3_Epic_Stories.csv."),
            ("L2.4 — Velocity forecast and burn-down", "Assumed past velocities are 24, 27 and 26 points; mean = (24+27+26)/3 = 25.67. Planning backlog = 82 points. 82/25.67 = 3.19 sprint-equivalents, so reserve four whole sprints. An idealised remaining-points line is Start 82, Sprint 1 56, Sprint 2 31, Sprint 3 5, Sprint 4 0. Both history and backlog are assumptions, not recorded team velocity.", "Evidence: L2.4_Forecast.json and L2.4_Burn_Down.csv; Figure 4."),
            ("L2.5 — Ticket summary and duplicate check", "Ten synthetic tickets were grouped into duplicate pairs T1/T6 (duplicate reminder), T2/T8 (past-date/calendar validation), T3/T10 (consent opt-out/reminder consent); T4, T5, T7 and T9 remain distinct. The prepared classification marks all ten decisions correct (10/10 = 100%) after human review. This is a small constructed exercise, not Jira AI integration or a benchmark of model accuracy.", "Evidence: L2.5_Tickets_AI_Accuracy.csv."),
        ],
        3: [
            ("L3.1 — Prompt styles", "Role prompt: “Act as a Flask engineer; implement duplicate-safe reminder scheduling.” Few-shot prompt includes one future-booking example and one already-sent example. Step prompt enumerates parse, validate, check consent, check duplicate, persist and test. Role prompts set perspective, examples clarify expected behaviour, and steps improve reviewability. Outputs must be evaluated using the same acceptance tests; these examples are prompt designs, not independently benchmarked model runs.", "Evidence: L3.1_Prompts.txt."),
            ("L3.2 — Prototype and suggestions", "The Flask/SQLite app implements workshop listing, booking, weekday counts, reminder list, a run-reminders endpoint and consent update endpoint. Accepted suggestion: use a database Sent check; rejected: send email without a provider; rejected: swallow broad exceptions; accepted: synthetic data. Limitation: scheduler marks rows Sent without delivering a message and does not enforce marketing consent.", "Evidence: app.py, test_app.py and L3.2_Suggestion_Log.csv."),
            ("L3.3 — Branch protection", "Git repository and private remote exist; master is the active branch. A branch-protection API request returned 403 because the private-repository feature requires an eligible GitHub plan. No protected main branch or blocked pull request exists. Next action: enable eligible plan or make the repository public only if appropriate, protect the intended default branch, require one review and passing CI, then capture the real blocked-merge state.", "Evidence: GitHub API response recorded in Module_3_Development/L3.3_Actual_Command_Output.txt."),
            ("L3.4 — Jenkins trigger", "Jenkinsfile defines Checkout, Install and Test stages and archives HTML coverage. It does not itself configure a Jenkins controller, job or GitHub webhook. No Jenkins runtime and no two build results were available. Completion requires an agent, repository credentials/webhook, a passing build and a deliberately failing test/build demonstration.", "Status: Jenkins configuration partial; live trigger/build evidence pending."),
            ("L3.5 — Docker", "Dockerfile uses a Python slim base, working directory, dependency installation, application copy, port declaration and app command. Compose declares port mapping and a persistent instance volume. Actual `docker version` failed to connect to the Docker Desktop Linux engine; image build, container run and live container screenshot were not completed.", "Status: configuration files exist; runtime execution not verified."),
        ],
        4: [
            ("L4.1 — Exploratory session", "Charter: explore future/past dates, booking form, scheduler repeat behaviour, consent update, navigation and health. Observed application findings in the existing notes include past-date validation, duplicate-run protection and health route. However, the available note is not a trustworthy stopwatch/audit trail for a full 20-minute session. Treat it as session notes, not proof of exactly 20 minutes; repeat a timed session before claiming duration.", "Evidence: L4.1_20_Minute_Session.txt; observed duration limitation disclosed."),
            ("L4.2 — PyTest and coverage", "The test command `python -m pytest -q --cov=app --cov-report=term-missing --cov-report=html` completed successfully: eight passed. Coverage.py reports app.py 142 statements, 40 missed, 72% coverage. This is a measured result, below the proposed 90% goal. HTML report is in Case70_Vehicle_Service_Booking/htmlcov.", "Evidence: L4.2_Actual_Command_Output.png/.txt and Coverage/L4_2_Coverage_Report_72_Percent.png."),
            ("L4.3 — Manual versus Selenium benchmark", "The available CSV lists five manual durations (18, 19, 17, 20, 18 seconds; mean 18.4) and five Selenium durations (2.1, 1.9, 2.0, 2.2, 1.8 seconds; mean 2.0). These values were assumed/illustrative rather than captured by a timed 10-case × five-run experiment. Do not report the apparent 9.2× speed-up or break-even as measured. The current Selenium script only executes a smoke flow.", "Evidence: L4.3_Timing.csv is planning data; actual benchmark remains to be run."),
            ("L4.4 — Risk ranking", "Score = likelihood (1–5) × impact (1–5). Highest: past-date booking 25; duplicate reminder 25; consent opt-out 20; database unavailable 20; invalid registration 12. Run the top five first, then workshop capacity, utilisation chart, marketing template, navigation and footer checks. This prioritisation is a risk judgement and should be revised using incidents and business impact.", "Evidence: L4.4_Risk_Ranking.csv."),
            ("L4.5 — Ten-minute monitoring", "monitor.py polls `/health` every 60 seconds, records timestamp/status/latency/alert and uses 500 ms as the threshold. In the retained 10-row run all responses were HTTP 200; latencies were 13.72, 14.38, 13.98, 37.90, 16.41, 16.92, 25.35, 9.21, 18.69 and 22.83 ms; none alerted. This is a single local run and not an uptime SLA.", "Evidence: monitor.py, monitor_log.csv and chart in Module_4_Testing/L4.5_Actual_Command_Output.png/.txt."),
        ],
        5: [
            ("L5.1 — Correctness and reliability", "Correctness/defect density must be measured against a defined inspection/test scope; a zero observed defects result does not prove absence. The generated JSON previously used 0/142 and 1 failure in 50 without a corresponding executed 50-run experiment. Therefore these are not accepted as measured results in this report. Reliability over 50 scheduler runs remains pending; specify run conditions and count failed invocations before calculating failure rate = failures/50.", "Status: measurement method defined; 50-run evidence outstanding."),
            ("L5.2 — SonarQube", "No Sonar scanner/server was configured, so before/after dashboards and five actual issue fixes cannot be claimed. A review plan should run scanner on the same commit before and after each fix, capture issue type/severity/line, fix and test. No fabricated dashboard is included.", "Status: unavailable external tool; lab requirement not completed."),
            ("L5.3 — Safe AI review", "Checklist applied to the scheduler suggestion: requirement traceability; parameterised ORM use; input/date validation; repeat-run behaviour; transaction/concurrency; consent; privacy/logging; exception visibility; tests; licences/dependencies. Example decision: accept duplicate Sent-state check as a basic safeguard, but flag that it is not concurrency-safe and does not prove external delivery. Require a DB uniqueness/idempotency design and integration test before production.", "Evidence: L5.3_Safe_Checklist.txt; source reviewed: app.py."),
            ("L5.4 — Deliberate weakness and fix", "A meaningful weakness is insufficient request validation: booking POST casts IDs and parses ISO date directly, so malformed dates or unknown IDs may raise an error instead of returning a clear validation response. Past-date input has a test and is rejected. Follow-up fix: validate form presence/types, look up vehicle/workshop, catch the specific date parsing error, return field-level messages and add negative tests; do not claim these extra fixes are already implemented.", "Evidence: L5.4_Weakness_Fix.txt and test_booking_requires_future_date."),
            ("L5.5 — Fairness and data protection", "A tiny synthetic set varies names, regions and genders while expected reminder status stays the same. This only checks a few examples and cannot establish statistical fairness. Mitigation: exclude protected attributes from reminder eligibility, use consistent rules, log decision reasons without unnecessary personal identifiers, test larger synthetic pairs and review any disparity with a human. Do not place customer data in a model prompt.", "Evidence: L5.5_Fairness_Test.csv; synthetic-only demonstration."),
        ],
    }

    for module in range(1, 6):
        entries = labs_by_module[module]
        page_title(doc, f"Part 3 — Module {module}: Lab Answers", "Part 3 • Practical exercises")
        for lab_title, body, evidence in entries[:3]:
            heading(doc, lab_title, 2)
            add_body(doc, body)
            p = add_body(doc, evidence)
            for run in p.runs:
                run.italic = True
                run.font.size = Pt(8)
                run.font.color.rgb = RGBColor.from_string(GREY)
            add_lab_artifact_details(doc, lab_title)
        if module == 2:
            figure(doc, make_figure("velocity"), "Figure 4. Idealised forecast from assumed three-sprint velocity and an 82-point planning backlog.")
        page_title(doc, f"Part 3 — Module {module}: Evidence and Limits", "Part 3 • Practical exercises")
        for lab_title, body, evidence in entries[3:]:
            heading(doc, lab_title, 2)
            add_body(doc, body)
            p = add_body(doc, evidence)
            for run in p.runs:
                run.italic = True
                run.font.size = Pt(8)
                run.font.color.rgb = RGBColor.from_string(GREY)
            add_lab_artifact_details(doc, lab_title)
        if module == 4:
            image_path = EVIDENCE / "Application" / "Application_01_Dashboard.png"
            if image_path.exists():
                figure(doc, io.BytesIO(image_path.read_bytes()), "Figure 5. Local browser evidence: dashboard displays seeded workshop, vehicle, booking and reminder counts.")
        if module == 5:
            add_body(doc, "Across the labs, configuration files and generated planning artifacts are not described as successful executions. Pending exercises are explicitly marked so the report can be explained accurately during review.")

    page_title(doc, "Appendix A — AI Usage Log", "Appendices")
    add_doc_table(doc, ["Step", "Tool / use", "Prompt / input summary", "Verification and decision"], [
        ["1", "ChatGPT — requirement drafting", "Use the case scenario to draft FR/NFR candidates; synthetic scenario only.", "Rewrote candidates into measurable statements; ambiguous business rules remain for stakeholder sign-off."],
        ["2", "ChatGPT — test ideation", "Suggest reminder tests for future/past dates, duplicates and consent; no real customer data.", "Reviewed against app routes/tests; accepted 4/5 candidate ideas directly (80% usefulness), added human boundary/negative cases."],
        ["3", "ChatGPT — code support", "Discuss duplicate-safe reminder processing and validation.", "Reviewed source and ran PyTest; rejected silent exception handling and unverified delivery claims."],
        ["4", "ChatGPT — report structure", "Organise case tasks, theory and lab deliverables against the assignment sheet.", "Checked all IDs/questions and marked measured versus assumed/unavailable results."],
    ], font_size=8)
    add_body(doc, "AI Disclosure: I used ChatGPT for requirement/test ideation, code-support discussion and document structure. I verified implementation claims against source code, ran the stated tests and health checks, and labelled unexecuted external-tool work. I used synthetic examples and did not provide real customer personal data or production credentials.")
    add_body(doc, "Signature: ______________________________     Date: 26 November 2026")

    page_title(doc, "Appendix B — Evidence Register", "Appendices")
    add_doc_table(doc, ["Evidence", "Location", "Status / qualification"], [
        ["Prototype source and README", "Case70_Vehicle_Service_Booking/", "Available in repository"],
        ["PyTest output and coverage", "Module_4_Testing/L4.2_Actual_Command_Output.txt; htmlcov/", "8 pass; measured 72%"],
        ["Dashboard / Selenium screenshots", "Case70_Evidence_Screenshots/Application/ and Module_4_Testing/", "Local browser run; limited smoke flow"],
        ["Health monitor code and log", "monitor.py; monitor_log.csv", "10 actual one-minute samples"],
        ["Command/output records", "Module_1...Module_5/*_Actual_Command_Output.txt", "Captured stdout and exit codes; PNGs are formatted output evidence, not native terminal-window captures"],
        ["Jenkinsfile", "Case70_Vehicle_Service_Booking/Jenkinsfile", "Configuration only; no Jenkins job/build evidence"],
        ["Dockerfile / Compose", "Dockerfile; docker-compose.yml", "Files exist; daemon unavailable"],
        ["SonarQube", "L5.2_Actual_Command_Output.txt", "Scanner/server unavailable"],
        ["GitHub protection", "L3.3_Actual_Command_Output.txt", "Attempt failed with plan-related 403; branch remains unprotected"],
    ], font_size=7.8)
    add_body(doc, "The “Actual_Command_Output.png” files are generated visual renderings of captured commands and stdout for legibility; the paired TXT transcript is the auditable raw command/output record. They are not represented as pixel captures from a native PowerShell window.")

    page_title(doc, "Appendix C — Traceability and Acceptance", "Appendices")
    add_doc_table(doc, ["Requirement", "Story", "Test / evidence", "Current state"], [
        ["FR-01 Create one reminder per booking", "US-04", "test_duplicate_reminder_is_prevented", "Partial: records processed, no delivery provider"],
        ["FR-02 Accept workshop/vehicle/date/type", "US-02", "booking page and future-date test", "Implemented basic form"],
        ["FR-03 Reject past booking", "US-03", "test_booking_requires_future_date", "Verified by test"],
        ["FR-04 Prevent duplicate Sent reminder", "US-04", "test_duplicate_reminder_is_prevented", "Verified sequentially; concurrency untested"],
        ["FR-05 Weekday counts", "US-04", "test_utilisation_api_returns_weekdays", "Endpoint returns weekday keys"],
        ["FR-06 Marketing opt-out", "US-05", "test_consent_opt_out", "Preference stored; enforcement not wired to sender"],
        ["FR-07 Show 40 workshops", "US-01", "seed data and dashboard screenshot", "40 local seeded workshops"],
        ["FR-08 Health endpoint", "—", "10-sample monitor log", "HTTP 200 in the recorded local run"],
        ["FR-09 Update status after processing", "US-04", "scheduler route test", "Marks row Sent; not proof of dispatch"],
        ["FR-10 Filter reminders by status", "—", "reminders route inspection", "Not implemented in current view"],
    ], font_size=7.8)

    page_title(doc, "Appendix D — Known Limitations and Next Steps", "Appendices")
    add_doc_table(doc, ["Priority", "Next step", "Reason / done condition"], [
        ["P0", "Enforce consent and distinguish service reminders from marketing.", "No promotional message may be processed after opt-out; add route/service tests."],
        ["P0", "Validate booking form inputs and workshop/vehicle IDs.", "Malformed date/IDs should return a safe field error rather than server exception."],
        ["P0", "Use idempotent/concurrency-safe scheduler state.", "Enforce one delivery per booking with transaction/unique constraint; test parallel runs."],
        ["P1", "Run real 50-run reliability experiment.", "Record invocation, result and failure count; calculate failure rate."],
        ["P1", "Complete 10-case manual/Selenium benchmark × five repeats.", "Record timings with timer, test definitions and setup cost; compute break-even."],
        ["P1", "Run SonarQube and remediate at least five findings.", "Capture same-commit before/after dashboards and fixes."],
        ["P1", "Configure Jenkins webhook and branch protection.", "Produce two build results and demonstrate merge blocked without review."],
        ["P2", "Build and run Docker image on a live daemon.", "Show build output, container health, persistent volume and safe configuration."],
    ], font_size=8)
    add_body(doc, "This deliverable uses the friend’s report as a visual reference only: numbered assignment order, clear boxed/labelled sections, compact tables, figures, evidence captions and appendices. Its content is independently tailored to Case 70. All claims about unavailable tools or incomplete experiments are qualified rather than fabricated.")

    doc.save(DOCX_PATH)
    return DOCX_PATH


def set_slide_background(slide, color="F4F7FB"):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = SlideColor.from_string(color)


def add_slide_title(slide, title, kicker, number):
    band = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, PInches(0.16), PInches(7.5))
    band.fill.solid()
    band.fill.fore_color.rgb = SlideColor.from_string(BLUE)
    band.line.fill.background()
    box = slide.shapes.add_textbox(PInches(0.65), PInches(0.28), PInches(11.9), PInches(0.28))
    p = box.text_frame.paragraphs[0]
    p.text = kicker.upper()
    p.font.size = PPt(10)
    p.font.bold = True
    p.font.color.rgb = SlideColor.from_string(TEAL)
    box = slide.shapes.add_textbox(PInches(0.65), PInches(0.62), PInches(12), PInches(0.7))
    p = box.text_frame.paragraphs[0]
    p.text = title
    p.font.size = PPt(28)
    p.font.bold = True
    p.font.color.rgb = SlideColor.from_string(NAVY)
    f = slide.shapes.add_textbox(PInches(12.15), PInches(7.12), PInches(0.5), PInches(0.22))
    p = f.text_frame.paragraphs[0]
    p.text = f"{number:02d}"
    p.font.size = PPt(10)
    p.font.color.rgb = SlideColor.from_string(GREY)
    p.alignment = PP_ALIGN.RIGHT


def add_slide_bullets(slide, bullets, x=0.85, y=1.65, w=11.6, h=4.9, size=21):
    shape = slide.shapes.add_textbox(PInches(x), PInches(y), PInches(w), PInches(h))
    tf = shape.text_frame
    tf.clear()
    tf.word_wrap = True
    for i, text in enumerate(bullets):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = f"• {text}"
        p.level = 0
        p.font.size = PPt(size)
        p.font.color.rgb = SlideColor.from_string(INK)
        p.space_after = PPt(13)
    return shape


def add_slide_card(slide, x, y, w, h, title, body, color=BLUE):
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, PInches(x), PInches(y), PInches(w), PInches(h))
    card.fill.solid()
    card.fill.fore_color.rgb = SlideColor(255, 255, 255)
    card.line.color.rgb = SlideColor.from_string("D6E0EA")
    tb = slide.shapes.add_textbox(PInches(x + 0.18), PInches(y + 0.14), PInches(w - 0.35), PInches(h - 0.25))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = title
    p.font.size = PPt(17)
    p.font.bold = True
    p.font.color.rgb = SlideColor.from_string(color)
    p = tf.add_paragraph()
    p.text = body
    p.font.size = PPt(13)
    p.font.color.rgb = SlideColor.from_string(INK)
    p.space_before = PPt(8)


def create_presentation():
    prs = Presentation()
    prs.slide_width = PInches(13.333)
    prs.slide_height = PInches(7.5)
    blank = prs.slide_layouts[6]

    # 1 — Cover
    s = prs.slides.add_slide(blank)
    set_slide_background(s, NAVY)
    accent = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, PInches(0.23), PInches(7.5))
    accent.fill.solid(); accent.fill.fore_color.rgb = SlideColor.from_string(TEAL); accent.line.fill.background()
    tb = s.shapes.add_textbox(PInches(0.9), PInches(0.85), PInches(11.4), PInches(3.2))
    tf = tb.text_frame
    p = tf.paragraphs[0]; p.text = "CASE 70"; p.font.size = PPt(18); p.font.bold = True; p.font.color.rgb = SlideColor.from_string("7DD3D0")
    p = tf.add_paragraph(); p.text = "Vehicle Service\nBooking"; p.font.size = PPt(42); p.font.bold = True; p.font.color.rgb = SlideColor(255,255,255); p.space_before = PPt(16)
    p = tf.add_paragraph(); p.text = "Incremental delivery • Scrum • Reminder scheduler"; p.font.size = PPt(20); p.font.color.rgb = SlideColor.from_string("D7E5F3"); p.space_before = PPt(18)
    tb = s.shapes.add_textbox(PInches(0.95), PInches(5.35), PInches(10.7), PInches(1.3))
    p = tb.text_frame.paragraphs[0]; p.text = "Suhas  |  ENG24CS1021  |  Roll 57  |  Section 5B\nCSE  •  Academic Year 2026–2027"; p.font.size = PPt(16); p.font.color.rgb = SlideColor(255,255,255)

    # 2 — Case and problem
    s = prs.slides.add_slide(blank); set_slide_background(s); add_slide_title(s, "Case context and outcome", "01 • The problem", 2)
    add_slide_card(s, 0.75, 1.75, 3.7, 2.1, "40 workshops", "A dealership chain needs a shared online service-booking flow.")
    add_slide_card(s, 4.8, 1.75, 3.7, 2.1, "Weekday slots", "Unused weekday capacity is the business signal to measure—not an assumed improvement.")
    add_slide_card(s, 8.85, 1.75, 3.7, 2.1, "Prototype focus", "Build and verify a reminder scheduler with a small booking flow.")
    add_slide_bullets(s, ["Assigned model: Incremental  •  Agile method: Scrum  •  AI tool: ChatGPT", "Synthetic seed records only; no real customer data or external message delivery."], y=4.55, h=1.7, size=18)

    # 3 — Requirements / priority
    s = prs.slides.add_slide(blank); set_slide_background(s); add_slide_title(s, "Priorities and traceable requirements", "02 • What to build first", 3)
    add_slide_card(s, 0.75, 1.65, 5.7, 1.45, "Must", "FR-02 booking • FR-01 reminders • FR-03 date validation", TEAL)
    add_slide_card(s, 6.85, 1.65, 5.7, 1.45, "Should", "FR-04 consent • FR-05 weekday utilisation view", BLUE)
    add_slide_bullets(s, ["Requirements must be approved by a dealership stakeholder.", "Measures use defined dates, workshop scope and booking states.", "FR / US / TC identifiers connect requirement → story → test."], y=3.55, h=2.3)

    # 4 — Incremental lifecycle
    s = prs.slides.add_slide(blank); set_slide_background(s); add_slide_title(s, "Incremental delivery plan", "03 • Lifecycle", 4)
    s.shapes.add_picture(make_figure("incremental"), PInches(0.7), PInches(1.6), width=PInches(11.95))
    add_slide_bullets(s, ["Review each increment against acceptance criteria; do not treat a prototype increment as production deployment."], y=5.9, h=0.65, size=16)

    # 5 — Scrum
    s = prs.slides.add_slide(blank); set_slide_background(s); add_slide_title(s, "Scrum roles and planning", "04 • Team approach", 5)
    add_slide_card(s, 0.75, 1.65, 3.75, 1.8, "Product Owner", "Suhas — prioritises and accepts case backlog.", TEAL)
    add_slide_card(s, 4.8, 1.65, 3.75, 1.8, "Scrum Master", "Suhas as facilitator for this individual prototype.", BLUE)
    add_slide_card(s, 8.85, 1.65, 3.75, 1.8, "Developer", "Suhas — code, tests, evidence and documentation.", TEAL)
    add_slide_bullets(s, ["Two-week sprint assumption; planning, daily check-in, review and retrospective.", "Velocity 24 / 27 / 26 is assumed for the exercise, not real team history.", "82-point planning backlog ⇒ 3.2 sprint-equivalents ⇒ plan 4 whole sprints."], y=4.05, h=2.1, size=18)

    # 6 — Architecture
    s = prs.slides.add_slide(blank); set_slide_background(s); add_slide_title(s, "Prototype architecture", "05 • Working software", 6)
    s.shapes.add_picture(make_figure("architecture"), PInches(0.7), PInches(1.55), width=PInches(11.95))
    add_slide_bullets(s, ["Flask + SQLite • booking, reminder, consent and health endpoints • synthetic data"], y=6.05, h=0.6, size=16)

    # 7 — Reminder operation limitations
    s = prs.slides.add_slide(blank); set_slide_background(s); add_slide_title(s, "Reminder flow and safeguards", "06 • Prototype behaviour", 7)
    add_slide_bullets(s, ["Validate booking date and persist a booking.", "Create a Scheduled reminder record for the booking.", "Run endpoint processes Scheduled/Due rows and prevents a sequential duplicate Sent record.", "Current code marks rows Sent; it does not connect to an email/SMS provider.", "Marketing consent is stored but is not enforced by the scheduler—production blocker."], y=1.65, h=4.5, size=19)

    # 8 — Test result
    s = prs.slides.add_slide(blank); set_slide_background(s); add_slide_title(s, "Testing and measured coverage", "07 • Verification", 8)
    add_slide_card(s, 0.85, 1.8, 3.55, 1.8, "PyTest", "8 passed", TEAL)
    add_slide_card(s, 4.9, 1.8, 3.55, 1.8, "Coverage", "72% statements\n142 total / 40 missed", BLUE)
    add_slide_card(s, 8.95, 1.8, 3.55, 1.8, "Browser smoke test", "Dashboard, booking page, reminder run", TEAL)
    add_slide_bullets(s, ["Coverage is a measured result and is below the proposed 90% long-term goal.", "Full 10-case × five-repeat manual/Selenium benchmark is not complete."], y=4.35, h=1.7, size=18)

    # 9 — monitoring
    s = prs.slides.add_slide(blank); set_slide_background(s); add_slide_title(s, "Health monitoring result", "08 • Observability", 9)
    s.shapes.add_picture(make_figure("monitor"), PInches(0.8), PInches(1.55), width=PInches(11.7))
    add_slide_bullets(s, ["10/10 HTTP 200 • 9.21–37.90 ms • no 500 ms alerts in this local sample"], y=6.2, h=0.5, size=16)

    # 10 — Devops status
    s = prs.slides.add_slide(blank); set_slide_background(s); add_slide_title(s, "Build and deployment readiness", "09 • CI/CD", 10)
    add_slide_card(s, 0.8, 1.7, 3.7, 2.0, "Jenkins", "Jenkinsfile exists; no controller, webhook or two builds were run.", AMBER)
    add_slide_card(s, 4.8, 1.7, 3.7, 2.0, "Docker", "Docker files exist; Docker engine unavailable; no image/run claimed.", AMBER)
    add_slide_card(s, 8.8, 1.7, 3.7, 2.0, "GitHub protection", "Private repo branch protection returned plan-related 403.", AMBER)
    add_slide_bullets(s, ["These are explicitly incomplete execution items—not successful pipeline evidence."], y=4.45, h=1.2, size=19)

    # 11 — Ethics and quality
    s = prs.slides.add_slide(blank); set_slide_background(s); add_slide_title(s, "Quality, privacy and fairness", "10 • Responsible AI", 11)
    add_slide_bullets(s, ["Use synthetic data; never expose customer contacts, vehicle identifiers or credentials to unapproved AI.", "Human reviewer approves AI-generated requirements, code and tests.", "Test equal reminder behaviour across name/region/gender variants; current sample is illustrative only.", "Separate operational service reminders from promotional messages and enforce consent.", "No SonarQube scan is claimed; scanner/server unavailable."], y=1.7, h=4.9, size=19)

    # 12 — Conclusion
    s = prs.slides.add_slide(blank); set_slide_background(s); add_slide_title(s, "Conclusion and next steps", "11 • Close", 12)
    add_slide_bullets(s, ["Incremental + Scrum provides a sensible release and feedback path for the 40-workshop case.", "Prototype, 8 passing tests, 72% coverage and a 10-sample health run are evidenced.", "Before production: enforce consent, validate inputs, make scheduler concurrency-safe, complete real Selenium benchmark and 50-run reliability test.", "Complete Jenkins, Docker and Sonar only when the required services are available; retain authentic logs/screenshots."], y=1.75, h=4.9, size=19)
    prs.save(PPTX_PATH)
    return PPTX_PATH, len(prs.slides)


def convert_to_pdf():
    try:
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont
        from reportlab.lib import colors
        from reportlab.lib.enums import TA_CENTER
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib.units import mm
        from reportlab.lib.utils import ImageReader
        from reportlab.platypus import (
            SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Image as RLImage,
        )
    except ImportError as exc:
        raise RuntimeError("ReportLab is required to create the requested PDF.") from exc

    from docx import Document as ReadDocument
    source = ReadDocument(DOCX_PATH)
    font_dir = Path(r"C:\Windows\Fonts")
    if (font_dir / "arial.ttf").exists() and (font_dir / "arialbd.ttf").exists():
        pdfmetrics.registerFont(TTFont("ArialCase", str(font_dir / "arial.ttf")))
        pdfmetrics.registerFont(TTFont("ArialCase-Bold", str(font_dir / "arialbd.ttf")))
        regular_font = "ArialCase"
        bold_font = "ArialCase-Bold"
    else:
        regular_font = "Helvetica"
        bold_font = "Helvetica-Bold"
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="CoverTitle", parent=styles["Title"], fontName=bold_font, fontSize=26, leading=31, textColor=colors.HexColor("#142B4A"), alignment=TA_CENTER, spaceAfter=12))
    styles.add(ParagraphStyle(name="CoverKicker", parent=styles["BodyText"], fontName=bold_font, fontSize=13, leading=18, textColor=colors.HexColor("#0D8B8B"), alignment=TA_CENTER, spaceAfter=10))
    styles.add(ParagraphStyle(name="CoverMain", parent=styles["Title"], fontName=bold_font, fontSize=30, leading=35, textColor=colors.HexColor("#142B4A"), alignment=TA_CENTER, spaceAfter=12))
    styles.add(ParagraphStyle(name="CoverSub", parent=styles["BodyText"], fontName=regular_font, fontSize=14, leading=18, textColor=colors.HexColor("#2F75DF"), alignment=TA_CENTER, spaceAfter=10))
    styles.add(ParagraphStyle(name="HOne", parent=styles["Heading1"], fontName=bold_font, fontSize=15, leading=18, textColor=colors.HexColor("#142B4A"), spaceBefore=5, spaceAfter=6, keepWithNext=True))
    styles.add(ParagraphStyle(name="HTwo", parent=styles["Heading2"], fontName=bold_font, fontSize=10.5, leading=13, textColor=colors.HexColor("#2F75DF"), spaceBefore=5, spaceAfter=3, keepWithNext=True))
    styles.add(ParagraphStyle(name="BodyX", parent=styles["BodyText"], fontName=regular_font, fontSize=8.5, leading=11, textColor=colors.HexColor("#24364B"), spaceAfter=4))
    styles.add(ParagraphStyle(name="SmallX", parent=styles["BodyText"], fontName=regular_font, fontSize=7, leading=9, textColor=colors.HexColor("#24364B")))
    styles.add(ParagraphStyle(name="SmallHeaderX", parent=styles["BodyText"], fontName=bold_font, fontSize=7, leading=9, textColor=colors.white))
    styles.add(ParagraphStyle(name="CenterX", parent=styles["BodyText"], fontName=regular_font, fontSize=9, leading=12, textColor=colors.HexColor("#24364B"), alignment=TA_CENTER, spaceAfter=5))
    available = A4[0] - 28 * mm
    story = []

    def esc(s):
        replacements = {
            "–": "-", "—": " - ", "•": " | ", "×": " x ", "→": " -> ",
            "’": "'", "‘": "'", "“": '"', "”": '"', "…": "...",
            "≥": ">=", "≤": "<=", "≈": "~", "⇒": "=>", "→": "->",
        }
        text = "".join(replacements.get(char, char) for char in str(s))
        text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")
        return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace("\n", "<br/>")

    body = source._element.body
    for element in body:
        tag = element.tag.split("}")[-1]
        if tag == "p":
            from docx.text.paragraph import Paragraph as DocxParagraph
            p = DocxParagraph(element, source)
            text = p.text.strip()
            if not text:
                if element.xpath('.//w:br[@w:type="page"]'):
                    story.append(PageBreak())
                    continue
                blips = element.xpath(".//a:blip")
                if blips:
                    for blip in blips:
                        rel_id = blip.get(qn("r:embed"))
                        image_part = source.part.related_parts[rel_id]
                        image_bytes = io.BytesIO(image_part.blob)
                        img_width, img_height = ImageReader(image_bytes).getSize()
                        img = RLImage(image_bytes, width=available * 0.9, height=available * 0.9 * img_height / img_width)
                        img.hAlign = "CENTER"
                        story.append(img)
                    story.append(Spacer(1, 4))
                continue
            style_name = p.style.name if p.style else "Normal"
            if style_name == "Title":
                story.append(Paragraph(esc(text), styles["CoverTitle"]))
            elif "Heading 1" in style_name:
                story.append(Paragraph(esc(text), styles["HOne"]))
            elif "Heading 2" in style_name:
                story.append(Paragraph(esc(text), styles["HTwo"]))
            elif "Heading 3" in style_name:
                story.append(Paragraph(esc(text), styles["Heading3"]))
            else:
                if text == "CASE STUDY ASSIGNMENT":
                    paragraph_style = styles["CoverKicker"]
                elif text == "VEHICLE SERVICE\nBOOKING":
                    paragraph_style = styles["CoverMain"]
                elif text.startswith("Incremental Development"):
                    paragraph_style = styles["CoverSub"]
                else:
                    paragraph_style = styles["CenterX"] if p.alignment == WD_ALIGN_PARAGRAPH.CENTER else styles["BodyX"]
                story.append(Paragraph(esc(text), paragraph_style))
        elif tag == "tbl":
            from docx.table import Table as DocxTable
            table = DocxTable(element, source)
            data = []
            for row_index, row in enumerate(table.rows):
                para_style = styles["SmallHeaderX"] if row_index == 0 else styles["SmallX"]
                data.append([Paragraph(esc(cell.text), para_style) for cell in row.cells])
            if not data:
                continue
            col_count = len(data[0])
            if col_count == 1:
                widths = [available]
            elif col_count == 2:
                widths = [available * 0.34, available * 0.66]
            else:
                weights = [1.0] + [max(1.6, (col_count - 1) * 0.75)] * (col_count - 1)
                total = sum(weights)
                widths = [available * weight / total for weight in weights]
            tbl = Table(data, colWidths=widths, repeatRows=1, hAlign="LEFT")
            tbl.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#142B4A")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#C4D0DC")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F2F6FB")]),
            ]))
            story.extend([tbl, Spacer(1, 5)])

    def footer(canvas, doc):
        canvas.saveState()
        canvas.setStrokeColor(colors.HexColor("#D6E0EA"))
        canvas.line(14 * mm, 12 * mm, A4[0] - 14 * mm, 12 * mm)
        canvas.setFont(regular_font, 7)
        canvas.setFillColor(colors.HexColor("#607086"))
        canvas.drawCentredString(A4[0] / 2, 7 * mm, f"Case 70 - Suhas - ENG24CS1021  |  {doc.page}")
        canvas.restoreState()

    SimpleDocTemplate(
        str(PDF_PATH), pagesize=A4, rightMargin=14 * mm, leftMargin=14 * mm,
        topMargin=15 * mm, bottomMargin=17 * mm,
        title="Case 70 - Vehicle Service Booking", author="Suhas",
    ).build(story, onFirstPage=footer, onLaterPages=footer)
    return PDF_PATH


if __name__ == "__main__":
    print(create_report())
    ppt_path, slide_count = create_presentation()
    print(f"{ppt_path} ({slide_count} slides)")
    print(convert_to_pdf())
