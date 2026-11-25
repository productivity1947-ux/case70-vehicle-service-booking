from docx import Document
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph as RLParagraph, Spacer, Table as RLTable, TableStyle, PageBreak
from reportlab.lib.enums import TA_CENTER
from pathlib import Path

root=Path(r"C:\Users\guest pc\Desktop\suhas")
docx=root/"Case70_Report"/"Case70_57_Suhas.docx"
pdf=root/"Case70_Report"/"Case70_57_Suhas.pdf"
source=Document(docx)
styles=getSampleStyleSheet()
styles.add(ParagraphStyle(name="ReportTitle",parent=styles["Title"],fontName="Helvetica-Bold",fontSize=22,leading=28,textColor=colors.HexColor("#13233A"),alignment=TA_CENTER))
styles.add(ParagraphStyle(name="ReportHeading",parent=styles["Heading1"],fontName="Helvetica-Bold",fontSize=15,leading=18,textColor=colors.HexColor("#13233A"),spaceBefore=8,spaceAfter=8))
styles.add(ParagraphStyle(name="ReportSubheading",parent=styles["Heading2"],fontName="Helvetica-Bold",fontSize=12,leading=15,textColor=colors.HexColor("#2F75DF"),spaceBefore=6,spaceAfter=5))
styles.add(ParagraphStyle(name="ReportBody",parent=styles["BodyText"],fontName="Helvetica",fontSize=9.5,leading=13,textColor=colors.HexColor("#1D2A3A"),spaceAfter=5))

def esc(text):
    return text.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")

story=[]
for block in source._element.body:
    tag=block.tag.split("}")[-1]
    if tag=="p":
        from docx.text.paragraph import Paragraph
        p=Paragraph(block,source)
        text=p.text.strip()
        if not text:
            continue
        style=p.style.name if p.style else "Normal"
        if "Title" in style or (len(source.paragraphs)==0):
            story.append(RLParagraph(esc(text).replace("\n","<br/>"),styles["ReportTitle"]))
        elif "Heading 1" in style:
            story.append(PageBreak()); story.append(RLParagraph(esc(text),styles["ReportHeading"]))
        elif "Heading 2" in style:
            story.append(RLParagraph(esc(text),styles["ReportSubheading"]))
        else:
            story.append(RLParagraph(esc(text).replace("\n","<br/>"),styles["ReportBody"]))
    elif tag=="tbl":
        from docx.table import Table
        table=Table(block,source)
        data=[[c.text for c in row.cells] for row in table.rows]
        t=RLTable(data,colWidths=[120 if i==0 else 180 for i in range(len(data[0]))])
        t.repeatRows=1
        t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#DFF3FF")),
                               ("TEXTCOLOR",(0,0),(-1,0),colors.HexColor("#13233A")),
                               ("GRID",(0,0),(-1,-1),0.5,colors.HexColor("#AAB8C6")),
                               ("FONTNAME",(0,0),(-1,-1),"Helvetica"),
                               ("FONTSIZE",(0,0),(-1,-1),8),
                               ("VALIGN",(0,0),(-1,-1),"TOP"),
                               ("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white,colors.HexColor("#F5F8FC")])]))
        story.append(t); story.append(Spacer(1,8))

SimpleDocTemplate(str(pdf),pagesize=A4,rightMargin=36,leftMargin=36,topMargin=36,bottomMargin=36,title="Case 70 - Vehicle Service Booking",author="Suhas").build(story)
print(pdf)
