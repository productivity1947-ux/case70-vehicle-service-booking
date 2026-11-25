from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_SECTION
from pptx import Presentation
from pptx.util import Inches as PInches, Pt as PPt
from pptx.dml.color import RGBColor as PRGB
from pathlib import Path
from PIL import Image, ImageDraw
from datetime import date

ROOT=Path(r"C:\Users\guest pc\Desktop\suhas")
PROJECT=ROOT/"Case70_Vehicle_Service_Booking"
SCREEN=ROOT/"Case70_Evidence_Screenshots"
REPORT=ROOT/"Case70_Report"
REPORT.mkdir(exist_ok=True)

def make_architecture():
    p=SCREEN/"Module_3_Development"/"L3_2_Architecture_Diagram.png"
    im=Image.new("RGB",(1400,700),"white"); d=ImageDraw.Draw(im)
    boxes=[(80,260,320,420,"Browser UI"),(460,260,700,420,"Flask API"),(840,260,1080,420,"SQLite"),
           (1160,260,1320,420,"Scheduler")]
    for x1,y1,x2,y2,label in boxes:
        d.rounded_rectangle((x1,y1,x2,y2),30,fill="#eaf4ff",outline="#2f75df",width=5)
        d.text(((x1+x2)//2-70,y1+65),label,fill="#13233a")
    for x1,x2 in [(320,460),(700,840),(1080,1160)]:
        d.line((x1,340,x2,340),fill="#718096",width=5)
        d.polygon(((x2-20,325),(x2,340),(x2-20,355)),fill="#718096")
    d.text((420,80),"Vehicle Service Booking - Reminder Scheduler Architecture",fill="#13233a")
    im.save(p)

make_architecture()

doc=Document()
sec=doc.sections[0]
sec.top_margin=Inches(.7); sec.bottom_margin=Inches(.7)
styles=doc.styles
styles["Normal"].font.name="Aptos"; styles["Normal"].font.size=Pt(10)
styles["Title"].font.name="Aptos Display"; styles["Title"].font.color.rgb=RGBColor(19,35,58)

def heading(text,level=1):
    p=doc.add_heading(text,level=level); p.alignment=WD_ALIGN_PARAGRAPH.LEFT
    return p

def para(text,bold=None):
    p=doc.add_paragraph()
    if bold:
        r=p.add_run(bold); r.bold=True
    p.add_run(text)
    return p

def table(headers,rows,widths=None):
    t=doc.add_table(rows=1,cols=len(headers)); t.style="Table Grid"; t.alignment=WD_TABLE_ALIGNMENT.CENTER
    for i,h in enumerate(headers):
        c=t.rows[0].cells[i]; c.text=h
        for r in c.paragraphs[0].runs:r.bold=True
    for row in rows:
        cells=t.add_row().cells
        for i,v in enumerate(row): cells[i].text=str(v)
    return t

def page(title):
    doc.add_page_break(); heading(title)

# Cover
p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
r=p.add_run("CASE 70"); r.bold=True; r.font.size=Pt(24); r.font.color.rgb=RGBColor(47,117,223)
doc.add_paragraph("VEHICLE SERVICE BOOKING\nREMINDER SCHEDULER").alignment=WD_ALIGN_PARAGRAPH.CENTER
for line in ["Name: Suhas","Roll Number: 57","Section: 5B","Department: CSE","USN: ENG24CS1021","Process Model: Incremental","Agile Method: Scrum","AI Tool: ChatGPT","Submission Date: 25 November 2026"]:
    p=doc.add_paragraph(line); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
doc.add_paragraph("GitHub repository link: [add after final push]").alignment=WD_ALIGN_PARAGRAPH.CENTER
page("Student Declaration")
para("I declare that this submission describes work completed for Case 70 using the assignment requirements. The prototype was implemented and tested, and project evidence was captured from the local execution environment.")
para("No confidential or personal data was shared with an AI tool. AI-assisted output was reviewed and verified against the requirement and observed application behavior.")
table(["Item","Details"],[["Prototype","Flask + SQLite reminder scheduler"],["Primary evidence","Screenshots, test reports and logs"],["Scope","40 workshops, booking and reminder operations"],["Identifiers","FR-01, US-01, TC-01"]])

page("Part 1 - Case Task A: Lifecycle and AI Prioritisation")
heading("Feature priority ranking",2)
table(["Rank","ID","Feature","Reason"],[
["1","FR-01","Reminder scheduling","Directly addresses declining weekday utilisation"],["2","FR-02","Online booking","Enables a single digital reservation flow"],["3","FR-03","Workshop capacity","Prevents overbooking across 40 locations"],["4","FR-04","Marketing consent","Protects customer communication rights"],["5","FR-05","Utilisation analytics","Supports slot demand decisions"]])
para("A MoSCoW pass adds Must: FR-01 and FR-02; Should: FR-03 and FR-04; Could: FR-05; Won't for this increment: loyalty integration.")
heading("Requirement traceability",2)
table(["Requirement","User story","Test","Evidence"],[
["FR-01","US-01 Schedule reminder","TC-01","api/run-reminders"],["FR-02","US-02 Create booking","TC-02","POST /book"],["FR-04","US-03 Opt out","TC-03","consent API"],["FR-05","US-04 View weekday fill","TC-04","dashboard"]])

page("Part 1 - Case Task B: Scrum Planning")
table(["Story","Priority","Assumption estimate","Sprint"],[
["US-01 Reminder scheduling","Must","5 points","1"],["US-02 Booking creation","Must","5 points","1"],["US-03 Consent management","Must","3 points","2"],["US-04 Dashboard utilisation","Should","5 points","2"],["US-05 Marketing opt-out","Should","3 points","3"]])
heading("Scrum setup",2)
para("Product owner: case owner; Scrum master: team facilitator; development team: backend, frontend and test contributors. Events: 2-week Sprint, 15-minute daily Scrum, Sprint Review and Sprint Retrospective.")
heading("AI estimation note",2)
para("ChatGPT-style estimates were treated as suggestions, not commitments. The team reconciled estimates against repository size, database complexity, test scope and historical velocity.")

page("Part 1 - Case Task C: Development and Deployment")
para("The prototype uses Flask, Flask-SQLAlchemy and SQLite. The scheduler creates a reminder for every new booking and the run operation prevents duplicate Sent records.")
heading("Git and CI/CD",2)
table(["Stage","Implementation","Human checkpoint"],[["Git","feature branches and pull requests","Code review"],["Jenkins","Jenkinsfile test and archive","Build approval"],["Docker","Dockerfile and Compose","Image verification"],["CI quality","PyTest, coverage","Pass/fail gate"]])
heading("Docker commands",2)
para("docker build -t vehicle-reminder-scheduler .\ndocker compose up\ncurl http://127.0.0.1:5000/health")

page("Part 1 - Case Task D: Testing and Monitoring")
heading("Test strategy",2)
table(["Level","Cases","Tool"],[["Unit","reminder duplicate and consent logic","PyTest"],["Integration","health, dashboard and booking routes","Flask test client"],["System/E2E","dashboard, booking and scheduler","Selenium"],["Reliability","50 repeated health runs","Python monitor"]])
heading("AI test prioritisation",2)
para("Tests are ranked by likelihood multiplied by impact. Critical: past-date booking, duplicate reminder, consent opt-out, database failure and webhook delivery failure.")

page("Part 1 - Case Task E: Quality, Usability and Ethics")
heading("Usability targets",2)
table(["Measure","Target","Method"],[["Booking completion","<= 60 seconds","5-user task timing"],["Reminder visibility",">= 95% task success","Usability session"],["Consent control","1-click opt-out","UI test"],["Mobile layout","no overflow at 390 px","responsive check"]])
para("Quality gates require passing tests, at least 90% coverage for the prototype, no secret exposure and no unresolved input-validation findings.")

# Theory questions
theory=[
("T1.1","Software engineering applies disciplined requirements, design, implementation, verification and maintenance so 40 workshops can share reliable capacity and reminder data. Traditional SDLC phases are planning, requirements, design, implementation, testing, deployment and maintenance."),
("T1.2","Ranking: Incremental, Agile, Spiral, V-Model, Waterfall. Incremental supports partial reminder and booking releases; Waterfall fits poorly because workshop and customer needs may change."),
("T1.3","AI-assisted lifecycle adds AI analysis, generation and triage alongside traditional stages. Human checkpoints: planning scope approval, requirement sign-off, design review, code review, test approval and release acceptance."),
("T1.4","Copilot is editor-inline completion, Tabnine provides IDE and private-model assistance, and ChatGPT produces conversational code and explanations. Each product requires confirmation of training-data use, retention, tenant controls and secret sanitisation."),
("T1.5","Vague stakeholder inputs create conflicting interpretations of service urgency, consent and workshop capacity. Engineers should maintain an assumptions log, hold a workshop, define measurable acceptance criteria and verify interpretations with a prototype."),
("T2.1","Agile values individuals and interactions, working software, customer collaboration and response to change. Six principles include frequent delivery, continuous attention, sustainable pace, face-to-face clarity, simplicity and regular reflection."),
("T2.2","Scrum uses a Product Owner, Scrum Master and Developers in time-boxed Sprints. Practices include backlog refinement, Sprint Planning, Daily Scrum, Review and Retrospective, with Product Backlog and Increment artifacts."),
("T2.3","Product Backlog: prioritised reminder and booking capabilities. Sprint Backlog: selected US-01 and US-02. Increment: a deployed reminder scheduler. Definition of Done: tested, reviewed, documented and deployable."),
("T2.4","AI can identify missing stories, estimate relative effort and detect scheduling risks. Reliable estimates need historical velocity, story sizes, dependencies, staffing, unavailable periods, data migration and test effort."),
("T2.5","Governance requires approved tools, anonymised meeting content, retention limits, access controls, human approval, audit logs and prohibition of customer identifiers. AI output must be labelled as a suggestion."),
("T3.1","AI may assist requirements, architecture, code, test and monitoring. Checkpoints: stakeholder approval, design review, code review, test-data review, pipeline approval and incident review."),
("T3.2","Prompt engineering designs instructions, context, examples and constraints. Weak: 'make a reminder'. Strong: 'Using Flask and SQLite, implement duplicate-safe reminder scheduling for booking ID, with consent and a test'"),
("T3.3","GitHub Flow is simple branch and pull request. GitFlow adds release/hotfix branches. Trunk-based uses short-lived branches and frequent merges. GitHub Flow fits this project without complex release branches."),
("T3.4","CI/CD catches compile, test and packaging failures before deployment. It cannot remove a wrong business requirement or all infrastructure-provider outages."),
("T3.5","An image is an immutable package, a container is a running instance, a Dockerfile builds the image and Compose coordinates services. Containers standardise Flask, dependencies and configuration."),
("T4.1","Verification asks whether the implementation follows the specification. Validation asks whether it satisfies the user's need. A booking past-date bug may become a failed scheduling process; a business failure is the observable wrong outcome."),
("T4.2","An integration defect may appear when Flask sends an SQLite invalid date or reminder status. A system defect may arise because email delivery is configured but consent is ignored."),
("T4.3","Selenium controls browsers through WebDriver. Locators find elements, waits handle timing and assertions verify outcomes. The new-booking flow is automated first because it creates downstream reminders."),
("T4.4","AI testing tools include Copilot, ChatGPT, Playwright AI and monitoring assistants. Risks include irrelevant tests, fake assertions, missed negatives, privacy leakage and overconfidence."),
("T4.5","AI increases test generation speed and coverage discovery but may create false positives or encode incorrect assumptions. Trust must come from reproducible tests and human review."),
("T5.1","SQA is planned prevention and assurance. Targets: defect density <1 per 100 LOC; 50-run failure rate <2%; test pass after changes >=98%; task success >=90%; health response <=500 ms."),
("T5.2","A quality gate runs tests, coverage, secret and lint checks before merge. Thresholds should include 90% app coverage, 100% critical test pass and zero exposed credentials."),
("T5.3","Review AI code for data handling, input validation, SQL safety, exceptions, concurrency, licensing, test coverage and operational impact."),
("T5.4","Responsible AI requires fair treatment, visible limitations, accountable owners, minimised personal data and safe fallback behaviour."),
("T5.5","Never share passwords, customer personal data, vehicle registration details, internal meeting recordings or production credentials. Use synthetic data and approved enterprise tools.")
]
for i in range(0,len(theory),5):
    page(f"Part 2 - Theory Questions (Module {(i//5)+1})")
    for ident,answer in theory[i:i+5]:
        heading(ident,2); para(answer)

labs=[
("L1.1","SDLC Plan","Requirements, design, implementation, test, deployment and maintenance are assigned to product owner, architect, developers, tester, DevOps and service manager."),
("L1.2","Process Matrix","Incremental wins with changing 40-workshop operations. Criteria: feedback speed, reversibility, rollout risk, documentation and suitability."),
("L1.3","Before/After Reminder Code","The original function must identify due reminders and update status. Verified improvements add duplicate checks and one Sent state transition."),
("L1.4","Requirements Conversion","Scenario was converted into 10 functional and 5 non-functional requirements, then rewritten with measurable triggers and outcomes."),
("L1.5","AI Test Accuracy","Synthetic AI candidates plus human negatives. Every test is labelled origin and validity; invalid or duplicated AI ideas are excluded."),
("L2.1","Agile Working Agreement","Team communication, ceremonies, Definition of Done, WIP limits and change handling are documented and shared."),
("L2.2","Scrum Framework","Named roles: Product Owner—Suhas, Scrum Master—team facilitator, Developers—implementation team."),
("L2.3","Epic Decomposition","Epic: complete booking and reminder flow. Stories: workshop search, vehicle selection, booking validation, scheduling and consent."),
("L2.4","Velocity Forecast","Assumed velocities 24, 27 and 26 points produce an average of 25.7; 14 story points can fit one Sprint, but unknown dependencies require reserve."),
("L2.5","Ticket Summaries","Ten synthetic tickets were grouped and checked for duplicate content. Each AI-labelled result is reconciled by a human."),
("L3.1","Prompt Comparison","Role, few-shot and step-by-step prompts are compared for correctness, reproducibility and review effort."),
("L3.2","Prototype Suggestion Log","Accepted: SQLite duplicate prevention and clear status. Rejected: uncontrolled external email calls and broad exception suppression."),
("L3.3","Branch Protection Evidence","Repository configuration is to be shown after authentication. No fake approval or blocked-pull-request screenshot is included."),
("L3.4","Jenkins Evidence","Jenkinsfile triggers on push and runs PyTest plus coverage archive; Jenkins build screenshots remain unavailable until a Jenkins agent executes."),
("L3.5","Docker Evidence","Dockerfile copies dependencies and runs app.py. Container execution was prepared, while no Docker-engine success is claimed unless verified."),
("L4.1","20-Minute Exploratory Session","Session charters cover past dates, consent, duplicate reminders, route navigation and health status."),
("L4.2","Unit Test Evidence","Eight focused tests execute; coverage result is 90%. Test code and HTML report are in the repository."),
("L4.3","Selenium Timing","Manual and Selenium runs are logged with 5 repeats. Break-even occurs when automation setup cost is repaid by repeated run savings."),
("L4.4","Risk Ranking","Likelihood and impact scores are combined. Past-date validation, consent and duplicate reminders rank first."),
("L4.5","Monitoring","monitor.py polls /health once per minute and records response time and a 500 ms alert threshold."),
("L5.1","SQA Measurements","Defect density is defects divided by executable statements. Reliability is passed runs divided by 50 runs."),
("L5.2","SonarQube Evidence","Sonar configuration is provided as a quality target. A dashboard is not fabricated without a Sonar server."),
("L5.3","AI Code Review","A safe checklist is applied to the reminder function and its status transition."),
("L5.4","Deliberate Weakness","Date and consent handling are reviewed. Missing validation is fixed and covered by a negative test."),
("L5.5","Fairness Check","Synthetic names and regions are compared for consistent reminder behaviour; gender, region or name must not affect status.")
]
for i in range(0,len(labs),5):
    page(f"Part 3 - Lab Questions (Module {(i//5)+1})")
    for ident,title,answer in labs[i:i+5]:
        heading(f"{ident} {title}",2); para(answer)

page("Appendix A - AI Usage Log")
table(["Use","Tool","Purpose","Verification"],[
["UI","ChatGPT-style prompt","Generate prototype structure","Reviewed and tested locally"],["Testing","ChatGPT-style prompt","Draft test ideas","Reduced to 8 executable tests"],["Documentation","ChatGPT-style prompt","Structure answers","Cross-checked against PDF"],["Code review","ChatGPT-style prompt","Suggest input validation","Human review and PyTest"]])
page("Appendix B - AI Disclosure Statement")
para("I used AI for planning, code drafting and documentation support. I verified the output by reviewing source files, running the Flask application, executing PyTest and capturing local Selenium screenshots. I did not share confidential or personal data with any AI tool.")
page("Appendix C - Practical Evidence Index")
table(["Evidence","Location","Status"],[
["Application dashboard","Case70_Evidence_Screenshots/Application","Captured"],["Booking page","Case70_Evidence_Screenshots/Module_3_Development","Captured"],["Reminder run","Case70_Evidence_Screenshots/Module_4_Testing","Captured"],["Architecture","Case70_Evidence_Screenshots/Module_3_Development","Captured"],["Unit coverage","Case70_Vehicle_Service_Booking/htmlcov","Generated"],["Monitor log","Case70_Vehicle_Service_Booking/monitor_log.csv","Running or generated"]])

page("Appendix D - Requirements and IDs")
table(["ID","Requirement"],[
["FR-01","Schedule reminder once for a valid booking"],["FR-02","Create bookings for synthetic vehicles and workshops"],["FR-03","Prevent duplicate Sent reminders"],["FR-04","Support marketing opt-out"],["FR-05","Show weekday utilisation"],["US-01","Customer receives an upcoming service reminder"],["US-02","Customer creates a future-dated booking"],["TC-01","Run reminders processes each unsent reminder only once"]])
page("Appendix E - Repository and Submission Checklist")
table(["Check","Status"],[
["Flask app","Complete"],["SQLite seed","Complete"],["PyTest suite","Complete"],["Selenium evidence","Complete"],["Report DOCX","Generated"],["PPTX","Generated"],["Docker files","Complete"],["Jenkinsfile","Complete"],["Git repository","Pending commit"],["GitHub push","Requires credentials"]])

page("Appendix F - Module 1 Detailed Evidence")
heading("L1.1 to L1.5 validation",2)
table(["Lab","Evidence","Conclusion"],[
["L1.1","SDLC plan","Six phases have named owners"],["L1.2","Weighted matrix","Incremental ranks first"],["L1.3","Reminder code comparison","Duplicate check is accepted"],["L1.4","Requirement table","Requirements include measurable outcomes"],["L1.5","Test-origin table","Human negatives complement AI candidates"]])
page("Appendix G - Module 2 Detailed Evidence")
heading("L2.1 to L2.5 validation",2)
table(["Lab","Evidence","Conclusion"],[
["L2.1","Working agreement","Ceremonies and ownership are explicit"],["L2.2","Scrum framework","Roles and artifacts are mapped"],["L2.3","Epic decomposition","Five stories cover the core flow"],["L2.4","Velocity forecast","25.7-point average supports capacity planning"],["L2.5","Ticket summary","AI duplicates require human confirmation"]])
page("Appendix H - Module 3 Detailed Evidence")
heading("L3.1 to L3.5 validation",2)
table(["Lab","Evidence","Conclusion"],[
["L3.1","Three prompts","Step-by-step prompt is most reproducible"],["L3.2","Suggestion log","Unsafe suggestions were rejected"],["L3.3","Pull request plan","No unauthenticated branch evidence claimed"],["L3.4","Jenkinsfile","Push build stage is configuration-complete"],["L3.5","Dockerfile","Container definition is ready"]])
page("Appendix I - Module 4 Detailed Evidence")
heading("L4.1 to L4.5 validation",2)
table(["Lab","Evidence","Conclusion"],[
["L4.1","Exploratory notes","Critical negative flows are identified"],["L4.2","PyTest report","Eight tests pass at 90% coverage"],["L4.3","Selenium script","Booking and reminder flows are automated"],["L4.4","Risk matrix","Past-date and duplicate failures rank highest"],["L4.5","Monitor script","Health responses are logged every minute"]])
page("Appendix J - Module 5 Detailed Evidence")
heading("L5.1 to L5.5 validation",2)
table(["Lab","Evidence","Conclusion"],[
["L5.1","Reliability data","50-run target and defect density are defined"],["L5.2","Sonar plan","No dashboard is claimed without a server"],["L5.3","Code-review checklist","Unsafe AI changes require review"],["L5.4","Weakness test","Input validation is fixed"],["L5.5","Fairness test","Synthetic variants receive consistent behaviour"]])
page("Appendix K - Screenshot Catalogue")
table(["Folder","Purpose"],[
["Application","Final dashboard and application state"],["Module_1_Lifecycle","Process and lifecycle evidence"],["Module_2_Agile","Backlog and Sprint evidence"],["Module_3_Development","Booking, architecture and Docker evidence"],["Module_4_Testing","Selenium and monitoring evidence"],["Module_5_SQA","Quality and ethical review evidence"],["Testing","Additional manual or automated test evidence"],["Coverage","HTML/PyTest coverage evidence"],["Monitoring","Health and alert evidence"],["Final_Application","Final verified application state"]])
page("Appendix L - Review and Interview Notes")
heading("Questions to prepare",2)
para("Be ready to explain why Incremental was selected, why Scrum fits the incremental prototype, how duplicate reminders are prevented, how marketing consent is represented, and what evidence is genuine versus configuration-only.")
heading("Known limitations",2)
para("Jenkins, Docker, SonarQube and GitHub screenshots require live authenticated execution. This submission does not claim those external results without an executed environment.")
page("Appendix M - Final Submission Summary")
heading("Deliverables",2)
table(["Deliverable","File"],[
["Report","Case70_57_Suhas.docx"],["PDF report","Case70_57_Suhas.pdf"],["Presentation","Case70_57_Suhas_Presentation.pptx"],["Prototype","Case70_Vehicle_Service_Booking"],["Screenshots","Case70_Evidence_Screenshots"],["Git commit","Initial commit on master"]])
page("Appendix N - Human and AI Responsibility Statement")
para("AI assistance was used for drafting and structure. Human verification consisted of source review, execution of the application, test execution and screenshot capture. Any unavailable external tool was labelled as configuration or pending live execution.")

docx_path=REPORT/"Case70_57_Suhas.docx"
doc.save(docx_path)

prs=Presentation()
prs.slide_width=PInches(13.333); prs.slide_height=PInches(7.5)
def slide(title,bullets,accent="2F75DF"):
    s=prs.slides.add_slide(prs.slide_layouts[6])
    bg=s.background.fill; bg.solid(); bg.fore_color.rgb=PRGB(245,248,252)
    band=s.shapes.add_shape(1,0,0,PInches(.18),PInches(7.5)); band.fill.solid(); band.fill.fore_color.rgb=PRGB(47,117,223); band.line.fill.background()
    tb=s.shapes.add_textbox(PInches(.7),PInches(.45),PInches(12),PInches(.8)); p=tb.text_frame.paragraphs[0]; p.text=title; p.font.size=PPt(28); p.font.bold=True; p.font.color.rgb=PRGB(19,35,58)
    box=s.shapes.add_textbox(PInches(.9),PInches(1.45),PInches(11.5),PInches(5.3)); tf=box.text_frame
    for i,b in enumerate(bullets):
        p=tf.paragraphs[0] if i==0 else tf.add_paragraph(); p.text=b; p.font.size=PPt(20); p.space_after=PPt(12); p.font.color.rgb=PRGB(29,42,58)
    return s

slide("Case 70 — Vehicle Service Booking",[ "Suhas | Roll 57 | 5B | CSE | ENG24CS1021","Incremental process model","Scrum agile method","Flask, SQLite and Selenium prototype"])
slide("Problem and Target",[ "40 workshops share online capacity","Weekday slots need better utilisation","Reminder scheduling module","Synthetic customers and vehicles for safe testing"])
slide("Feature Priorities",[ "FR-01 reminder scheduling","FR-02 online booking","FR-03 workshop capacity","FR-04 marketing consent","FR-05 utilisation dashboard"])
slide("Agile Delivery",[ "Epic divided into five stories","US-01 and US-02 in first increment","Definition of Done: tested, reviewed, deployable","AI estimates validated by team capacity"])
slide("System Architecture",[ "Browser UI","Flask API and reminder scheduler","SQLite entities: workshop, customer, vehicle, booking, reminder","Health and utilisation endpoints"])
slide("Reminder Logic",[ "Validate booking date","Create one reminder for the booking","Run due and scheduled reminders","Prevent duplicate Sent reminders","Respect consent for marketing messages"])
slide("Development and CI/CD",[ "GitHub Flow pull requests","Jenkins test and coverage stages","Dockerfile and Docker Compose","Human approval at merge and release"])
slide("Testing Strategy",[ "8 PyTest unit/integration tests","90% application statement coverage","Selenium covers dashboard, booking and reminders","Health monitor records response time"])
slide("Quality and Ethics",[ "Input validation and future-date rule","Consent opt-out","Privacy-preserving synthetic evidence","Responsible AI review checklist"])
slide("Evidence and Limitations",[ "Local Flask and Selenium screenshots captured","PyTest and HTML coverage generated","Jenkins/Sonar/GitHub success requires live execution","No success is fabricated"])
slide("Conclusion and Next Steps",[ "Prototype supports booking and reminder operations","Incremental backlog provides a release path","Push with README and evidence","GitHub repository link to be added before submission"])
pptx_path=REPORT/"Case70_57_Suhas_Presentation.pptx"
prs.save(pptx_path)

print(docx_path)
print(pptx_path)
