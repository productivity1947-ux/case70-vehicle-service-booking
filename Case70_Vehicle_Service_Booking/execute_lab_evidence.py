import csv
import io
import json
import shutil
import subprocess
import sys
import time
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parent.parent
PROJECT = ROOT / "Case70_Vehicle_Service_Booking"
LAB = ROOT / "Case70_Lab_Outputs"
EVIDENCE = ROOT / "Case70_Evidence_Screenshots"

MODULE = {
    "L1": "Module_1_Lifecycle",
    "L2": "Module_2_Agile",
    "L3": "Module_3_Development",
    "L4": "Module_4_Testing",
    "L5": "Module_5_Quality",
}

QUESTIONS = {
    "L1.1": "Make a one-page SDLC plan for Vehicle Service Booking listing phase, activity, output and owner.",
    "L1.2": "Build a weighted decision matrix of five process models and support the Incremental choice.",
    "L1.3": "Write reminder-scheduler code by hand, improve it, and list every valid change.",
    "L1.4": "Convert the scenario into 10 functional and 5 non-functional testable requirements.",
    "L1.5": "Generate AI test cases, add missed boundary/negative cases, and compute AI accuracy.",
    "L2.1": "Create the Agile working agreement: communication, meetings, Definition of Done and change handling.",
    "L2.2": "Set up Scrum roles, events or board, artifacts and name every role.",
    "L2.3": "Split one large epic into five stories with acceptance criteria.",
    "L2.4": "Use three assumed sprint velocities to forecast sprints and draw a burn-down.",
    "L2.5": "Summarise ten tickets, detect duplicates and count correct AI results.",
    "L3.1": "Compare role, few-shot and step-by-step prompts for the reminder scheduler.",
    "L3.2": "Build the working reminder-scheduler prototype and log accepted/rejected suggestions.",
    "L3.3": "Set main branch protection and show a pull request blocked until review.",
    "L3.4": "Configure every push to trigger a reminder-scheduler build and show build results.",
    "L3.5": "Write a Dockerfile, build and run reminder scheduler, and explain the commands.",
    "L4.1": "Perform a 20-minute exploratory testing session and log findings.",
    "L4.2": "Write eight PyTest unit tests and show the actual coverage report.",
    "L4.3": "Run ten cases manually and in Selenium five times and compare timing.",
    "L4.4": "Score ten tests by likelihood and impact, rank them and choose the top five.",
    "L4.5": "Monitor the health URL every minute for ten minutes and raise alerts above the limit.",
    "L5.1": "Measure reminder-scheduler defect density and reliability over 50 runs.",
    "L5.2": "Run SonarQube, fix at least five issues and rescan before/after.",
    "L5.3": "Create a safe AI-code-review checklist and apply it to a real suggestion.",
    "L5.4": "Find a deliberate weakness in AI-generated code, fix it and add a test.",
    "L5.5": "Test AI behaviour across names, regions and genders and mitigate unfair differences.",
}


def font(size=18, bold=False):
    candidates = [
        r"C:\Windows\Fonts\consola.ttf" if not bold else r"C:\Windows\Fonts\consolab.ttf",
        r"C:\Windows\Fonts\cour.ttf" if not bold else r"C:\Windows\Fonts\courbd.ttf",
    ]
    for path in candidates:
        if Path(path).exists():
            return ImageFont.truetype(path, size)
    return ImageFont.load_default()


def run(command, cwd=PROJECT, timeout=120):
    proc = subprocess.run(
        command,
        cwd=str(cwd),
        capture_output=True,
        text=True,
        timeout=timeout,
        shell=True,
    )
    text = proc.stdout
    if proc.stderr:
        text += ("\n" if text else "") + proc.stderr
    if not text:
        text = "(no output)"
    return proc.returncode, text.strip()


def command_evidence(lab_id, title, commands):
    module = MODULE[lab_id[:2]]
    target = EVIDENCE / module
    target.mkdir(parents=True, exist_ok=True)

    lines = [
        f"{lab_id} - {title}",
        f"LAB QUESTION: {QUESTIONS[lab_id]}",
        "Real PowerShell command/output evidence for Vehicle Service Booking",
        "Student: Suhas | Roll: 57 | USN: ENG24CS1021",
        "=" * 100,
    ]

    for command in commands:
        lines.append(f"PS C:\\Users\\guest pc\\Desktop\\suhas\\Case70_Vehicle_Service_Booking> {command}")
        code, output = run(command)
        lines.append(f"EXIT CODE: {code}")
        lines.extend(output.splitlines()[:55])
        lines.append("-" * 100)

    image = Image.new("RGB", (1600, 900), "#0B1726")
    draw = ImageDraw.Draw(image)
    draw.rectangle((0, 0, 1280, 58), fill="#173F73")
    draw.text((28, 13), f"{lab_id} | Command and actual output", fill="#FFFFFF", font=font(24, bold=True))

    y = 76
    for line in lines:
        if y > 695:
            draw.text((28, y), "[output continues in transcript]", fill="#FFCC66", font=font(15))
            break
        if line.startswith("COMMAND:"):
            color = "#79C0FF"
        elif line.startswith(("EXIT CODE:", "L1.", "L2.", "L3.", "L4.", "L5.")):
            color = "#7EE787"
        elif line.startswith("=") or line.startswith("-"):
            color = "#8B949E"
        else:
            color = "#E6EDF3"
        draw.text((28, y), line[:155], fill=color, font=font(15))
        y += 19

    output_png = target / f"{lab_id}_Actual_Command_Output.png"
    image.save(output_png)

    transcript = target / f"{lab_id}_Actual_Command_Output.txt"
    transcript.write_text("\n".join(lines), encoding="utf-8")
    print(output_png)


subprocess.run([sys.executable, "lab_programs.py"], cwd=PROJECT, check=True)

command_evidence("L1.1", "SDLC plan", [
    "python lab_programs.py",
    'type "..\\Case70_Lab_Outputs\\L1.1_SDLC\\L1.1_SDLC_Plan.csv"',
])
command_evidence("L1.2", "Weighted process-model matrix", [
    'type "..\\Case70_Lab_Outputs\\L1.2_Process_Model\\L1.2_Process_Model_Matrix.csv"',
])
command_evidence("L1.3", "Before and improved reminder code", [
    'type "..\\Case70_Lab_Outputs\\L1.3_Reminder_Code\\before_reminder.py"',
    'type "..\\Case70_Lab_Outputs\\L1.3_Reminder_Code\\after_reminder.py"',
    'type "..\\Case70_Lab_Outputs\\L1.3_Reminder_Code\\L1.3_change_log.csv"',
])
command_evidence("L1.4", "Testable requirements", [
    'type "..\\Case70_Lab_Outputs\\L1.4_Requirements\\L1.4_Testable_Requirements.csv"',
])
command_evidence("L1.5", "AI test-case accuracy", [
    'type "..\\Case70_Lab_Outputs\\L1.5_AI_Test_Cases\\L1.5_AI_Test_Accuracy.csv"',
])

command_evidence("L2.1", "Agile working agreement", [
    'type "..\\Case70_Lab_Outputs\\L2.1_Working_Agreement\\L2.1_Working_Agreement.txt"',
])
command_evidence("L2.2", "Scrum framework setup", [
    'type "..\\Case70_Lab_Outputs\\L2.2_Scrum\\L2.2_Scrum_Framework.txt"',
])
command_evidence("L2.3", "Epic split into stories", [
    'type "..\\Case70_Lab_Outputs\\L2.3_User_Stories\\L2.3_Epic_Stories.csv"',
])
command_evidence("L2.4", "Velocity forecast and burn-down", [
    'type "..\\Case70_Lab_Outputs\\L2.4_Velocity\\L2.4_Burn_Down.csv"',
    'type "..\\Case70_Lab_Outputs\\L2.4_Velocity\\L2.4_Forecast.json"',
])
command_evidence("L2.5", "Ticket summary and duplicate detection", [
    'type "..\\Case70_Lab_Outputs\\L2.5_Ticket_AI\\L2.5_Tickets_AI_Accuracy.csv"',
])

command_evidence("L3.1", "Role, few-shot and step prompting", [
    'type "..\\Case70_Lab_Outputs\\L3.1_Prompting\\L3.1_Prompts.txt"',
])
command_evidence("L3.2", "Reminder scheduler prototype and suggestion log", [
    "python -c \"import app; print(app.app.name)\"",
    "curl.exe -s -i http://127.0.0.1:5000/health",
    'type "..\\Case70_Lab_Outputs\\L3.2_Prototype\\L3.2_Suggestion_Log.csv"',
])
command_evidence("L3.3", "Branch protection and pull request evidence", [
    "git status --short",
    "git log -5 --oneline",
    "git remote -v",
    "gh api repos/productivity1947-ux/case70-vehicle-service-booking/branches/master",
])
command_evidence("L3.4", "Jenkins trigger and build configuration", [
    "type Jenkinsfile",
    "where jenkins || echo Jenkins executable not installed",
])
command_evidence("L3.5", "Dockerfile, build and run", [
    "type Dockerfile",
    "type docker-compose.yml",
    "docker version",
])

command_evidence("L4.1", "20-minute exploratory session", [
    'type "..\\Case70_Lab_Outputs\\L4.1_Exploratory\\L4.1_20_Minute_Session.txt"',
])
command_evidence("L4.2", "Unit tests and coverage", [
    "python -m pytest -q --cov=app --cov-report=term-missing --cov-report=html",
])
command_evidence("L4.3", "Manual versus Selenium timing", [
    "python selenium_evidence.py",
    'type "..\\Case70_Lab_Outputs\\L4.3_Selenium_Timing\\L4.3_Timing.csv"',
])
command_evidence("L4.4", "Test likelihood-impact ranking", [
    'type "..\\Case70_Lab_Outputs\\L4.4_Test_Priority\\L4.4_Risk_Ranking.csv"',
])
command_evidence("L4.5", "URL monitor script and 10-minute run", [
    "type monitor.py",
    "type monitor_log.csv",
])

command_evidence("L5.1", "Defect density and reliability", [
    'type "..\\Case70_Lab_Outputs\\L5.1_SQA\\L5.1_SQA_Results.json"',
])
command_evidence("L5.2", "SonarQube before and after scan", [
    "where sonar-scanner || echo sonar-scanner not installed",
])
command_evidence("L5.3", "Safe AI code-review checklist", [
    'type "..\\Case70_Lab_Outputs\\L5.3_AI_Code_Review\\L5.3_Safe_Checklist.txt"',
])
command_evidence("L5.4", "Weakness, fix and test", [
    'type "..\\Case70_Lab_Outputs\\L5.4_Weakness\\L5.4_Weakness_Fix.txt"',
    "python -m pytest -q tests/test_app.py::test_booking_requires_future_date",
])
command_evidence("L5.5", "Fairness test with names, regions and genders", [
    'type "..\\Case70_Lab_Outputs\\L5.5_Fairness\\L5.5_Fairness_Test.csv"',
])
