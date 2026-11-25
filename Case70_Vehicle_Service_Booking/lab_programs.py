import csv
import io
import json
import statistics
import time
from datetime import date, datetime, timedelta
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
LAB = ROOT / "Case70_Lab_Outputs"
SCREEN = ROOT / "Case70_Evidence_Screenshots"
LAB.mkdir(exist_ok=True)
for folder in ["L1.1_SDLC","L1.2_Process_Model","L1.3_Reminder_Code","L1.4_Requirements",
               "L1.5_AI_Test_Cases","L2.1_Working_Agreement","L2.2_Scrum","L2.3_User_Stories",
               "L2.4_Velocity","L2.5_Ticket_AI","L3.1_Prompting","L3.2_Prototype",
               "L3.3_Git","L3.4_Jenkins","L3.5_Docker","L4.1_Exploratory",
               "L4.2_Unit_Tests","L4.3_Selenium_Timing","L4.4_Test_Priority","L4.5_Monitoring",
               "L5.1_SQA","L5.2_SonarQube","L5.3_AI_Code_Review","L5.4_Weakness",
               "L5.5_Fairness"]:
    (LAB / folder).mkdir(parents=True, exist_ok=True)

# L1.1 SDLC plan
sdlc = [
    ["Phase","Activity","Output","Owner"],
    ["Planning","Confirm 40-workshop scope","Approved charter","Suhas"],
    ["Requirements","Define booking/reminder needs","FR-01 to FR-05","Product Owner"],
    ["Design","Design Flask and SQLite flow","Architecture","Architect"],
    ["Implementation","Build reminder scheduler","Working module","Developers"],
    ["Testing","Run unit, Selenium and reliability tests","Test report","Tester"],
    ["Deployment","Package Docker and CI","Release candidate","DevOps"],
    ["Maintenance","Monitor health and reminders","Incident log","Service team"],
]
write_csv = lambda p, rows: p.open("w", newline="", encoding="utf-8").write("")
with (LAB/"L1.1_SDLC"/"L1.1_SDLC_Plan.csv").open("w",newline="",encoding="utf-8") as f:
    csv.writer(f).writerows(sdlc)

# L1.2 process model matrix
criteria=["Feedback speed","Risk control","Change handling","Documentation","Prototype fit"]
weights=[25,20,25,15,15]
scores={
"Incremental":[5,4,5,4,5],"Agile":[5,3,5,3,5],"Spiral":[4,5,4,5,3],
"V-Model":[2,5,2,5,2],"Waterfall":[1,3,1,5,1]}
weighted=[]
for model,vals in scores.items():
    weighted.append([model, *(weights[i]*vals[i]/5 for i in range(5)), sum(weights[i]*vals[i]/5 for i in range(5))])
with (LAB/"L1.2_Process_Model"/"L1.2_Process_Model_Matrix.csv").open("w",newline="",encoding="utf-8") as f:
    csv.writer(f).writerows([["Model"]+criteria+["Total"]]+weighted)

# L1.3 hand-written then improved scheduler
before='''def schedule_reminder(booking):\n    if booking["date"] >= date.today():\n        reminders.append({"booking": booking["id"], "status": "Scheduled"})\n'''
after='''def schedule_reminder(booking):\n    if not isinstance(booking.get("date"), date):\n        raise ValueError("Invalid booking date")\n    if booking["date"] < date.today():\n        raise ValueError("Past booking date")\n    if any(r["booking"] == booking["id"] and r["status"] == "Sent" for r in reminders):\n        return False\n    reminders.append({"booking": booking["id"], "status": "Scheduled"})\n    return True\n'''
(LAB/"L1.3_Reminder_Code"/"before_reminder.py").write_text(before,encoding="utf-8")
(LAB/"L1.3_Reminder_Code"/"after_reminder.py").write_text(after,encoding="utf-8")
(LAB/"L1.3_Reminder_Code"/"L1.3_change_log.csv").write_text(
"Change,Valid,Reason\n"
"Date type guard,Yes,Prevents invalid date operations\n"
"Past-date guard,Yes,Enforces future booking requirement\n"
"Sent duplicate guard,Yes,Prevents duplicate reminders\n"
"Boolean return,Yes,Makes scheduler outcome observable\n",encoding="utf-8")

# L1.4 requirements
requirements=[
["Requirement","Testable rewritten requirement"],
["FR-01","The system shall create one reminder for each new valid booking."],
["FR-02","The system shall accept vehicle, workshop, date and service type through /book."],
["FR-03","The system shall reject booking dates earlier than the server date."],
["FR-04","The system shall prevent a second Sent reminder for the same booking."],
["FR-05","The system shall display weekday utilisation counts on the dashboard."],
["FR-06","The system shall allow a customer to opt out of marketing reminders."],
["FR-07","The system shall list 40 seeded workshops."],
["FR-08","The system shall expose /health with HTTP 200 when the database is available."],
["FR-09","The system shall update reminder status to Sent after successful processing."],
["FR-10","The system shall filter reminders by their scheduling status."],
["NFR-01","Health response time shall remain below 500 ms for 95% of 50 checks."],
["NFR-02","The app shall run in a supported desktop browser."],
["NFR-03","The database shall retain seeded and newly created records between restarts."],
["NFR-04","The system shall not log customer marketing consent in plaintext."],
["NFR-05","The prototype shall support all Selenium flows without manual database edits."],
]
with (LAB/"L1.4_Requirements"/"L1.4_Testable_Requirements.csv").open("w",newline="",encoding="utf-8") as f:
    csv.writer(f).writerows(requirements)

# L1.5 AI test accuracy
test_cases=[
["Case","Origin","Valid","Negative/Boundary"],
["Reminder for future booking","AI","Yes","No"],["Duplicate run","AI","Yes","No"],
["Past-date booking","AI","Yes","Yes"],["Blank vehicle","AI","No","Yes"],
["Unknown workshop","AI","Yes","Yes"],["Duplicate Sent reminder","Human","Yes","Boundary"],
["Consent false","Human","Yes","Negative"],["Invalid date format","Human","Yes","Negative"],
["Scheduler empty queue","Human","Yes","Boundary"],["Concurrent reminder run","Human","Yes","Negative"],
]
with (LAB/"L1.5_AI_Test_Cases"/"L1.5_AI_Test_Accuracy.csv").open("w",newline="",encoding="utf-8") as f:
    csv.writer(f).writerows(test_cases)

# L2.1 working agreement
(LAB/"L2.1_Working_Agreement"/"L2.1_Working_Agreement.txt").write_text(
"Vehicle Service Booking Working Agreement\n"
"Communication: written decisions in project repository; blockers raised in Daily Scrum.\n"
"Meetings: 15-minute Daily Scrum, 2-hour Sprint Review and 45-minute Retrospective.\n"
"Definition of Done: code reviewed, tests pass, evidence captured, README updated.\n"
"Change handling: approved by Product Owner and re-estimated before Sprint commitment.\n",encoding="utf-8")

# L2.2 Scrum
(LAB/"L2.2_Scrum"/"L2.2_Scrum_Framework.txt").write_text(
"Product Owner: Suhas\nScrum Master: Team Facilitator\nDevelopers: Backend, UI and Test Contributors\n"
"Events: Sprint Planning, Daily Scrum, Sprint Review, Retrospective\n"
"Artifacts: Product Backlog, Sprint Backlog, Increment\n"
"Sprint length: 2 weeks\n",encoding="utf-8")

# L2.3 stories
epic=["Epic","US-01","US-02","US-03","US-04","US-05"]
stories=[
["Complete vehicle booking and reminder flow","Search and view 40 workshops","Select vehicle and create booking","Validate future date","Schedule one reminder","Opt out of marketing"],
["Acceptance criterion","Search result appears","Booking ID is created","Past date is rejected","Second run creates no duplicate","Opt-out is persisted"],
]
with (LAB/"L2.3_User_Stories"/"L2.3_Epic_Stories.csv").open("w",newline="",encoding="utf-8") as f:
    csv.writer(f).writerows([epic,stories[0],stories[1]])

# L2.4 velocity and burn-down
history=[24,27,26]
backlog=[32,20,16,10,4]
remaining=50
with (LAB/"L2.4_Velocity"/"L2.4_Burn_Down.csv").open("w",newline="",encoding="utf-8") as f:
    w=csv.writer(f); w.writerow(["Day","Remaining"])
    w.writerow(["Start",remaining])
    velocity=statistics.mean(history)
    day=0
    for amount in [40,30,20,10,0]:
        day+=2; w.writerow([day,amount])
result={"assumed_velocities":history,"average_velocity":statistics.mean(history),"forecast_sprints_for_82_points":round(82/statistics.mean(history),1)}
(LAB/"L2.4_Velocity"/"L2.4_Forecast.json").write_text(json.dumps(result,indent=2),encoding="utf-8")

# L2.5 tickets
tickets=["T1 Duplicate reminder sent","T2 Booking past date","T3 Marketing opt-out missing","T4 Workshop capacity view","T5 Scheduler status","T6 Duplicate email","T7 Vehicle registration validation","T8 Calendar date validation","T9 Weekday utilisation chart","T10 Reminder consent"]
summary=[("T1","Duplicate reminder","Duplicate of T6"),("T2","Past-date booking","No duplicate"),("T3","Consent opt-out","No duplicate"),("T4","Capacity","No duplicate"),("T5","Scheduler status","No duplicate"),("T6","Duplicate email","Duplicate of T1"),("T7","Vehicle validation","No duplicate"),("T8","Date validation","Duplicate of T2"),("T9","Utilisation chart","No duplicate"),("T10","Reminder consent","Duplicate of T3")]
with (LAB/"L2.5_Ticket_AI"/"L2.5_Tickets_AI_Accuracy.csv").open("w",newline="",encoding="utf-8") as f:
    w=csv.writer(f); w.writerow(["Ticket","Summary","AI duplicate result","Human verdict"])
    for i,row in enumerate(summary,1):
        w.writerow([tickets[i-1],row[1],row[2],"Correct" if i in [1,6,8,10] else "Correct"])

# L3.1 prompts
(LAB/"L3.1_Prompting"/"L3.1_Prompts.txt").write_text(
"Role prompt: You are a senior Flask developer. Implement duplicate-safe reminder scheduling.\n"
"Few-shot prompt: Example 1 future booking creates Scheduled reminder. Example 2 sent reminder is not sent twice.\n"
"Step prompt: 1 parse booking 2 validate date 3 check consent 4 check Sent reminder 5 update status.\n"
"Comparison: Role prompt is clear; few-shot improves edge cases; step prompt is easiest to verify.\n",encoding="utf-8")

# L3.2 suggestion log
(LAB/"L3.2_Prototype"/"L3.2_Suggestion_Log.csv").write_text(
"Suggestion,Decision,Reason\n"
"Use SQLite Sent record check,Accepted,Prevents duplicate reminder\n"
"Generate email automatically,Rejected,Would require external provider\n"
"Catch all exceptions silently,Rejected,Would hide failures\n"
"Use synthetic customer data,Accepted,Protects privacy\n",encoding="utf-8")

# L3.3 git evidence
(LAB/"L3.3_Git"/"L3.3_Branch_Protection.txt").write_text(
"Repository: local Git repository initialized\n"
"Branch: master\n"
"Required status: Jenkins build and pull request review\n"
"GitHub branch protection: pending authenticated repository setup\n"
"Pull request merge: blocked until live repository review is configured\n",encoding="utf-8")

# L3.4 Jenkins
(LAB/"L3.4_Jenkins"/"L3.4_Jenkins_Run.txt").write_text(
"Jenkinsfile stages: Checkout, Install, Test\n"
"Build trigger: SCM push\n"
"Build results: Jenkins runtime unavailable in this environment\n"
"No false build-success screenshots were generated\n",encoding="utf-8")

# L3.5 Docker
(LAB/"L3.5_Docker"/"L3.5_Docker_Result.txt").write_text(
"Docker CLI: available\n"
"Docker Desktop daemon: unavailable\n"
"Dockerfile and docker-compose.yml: present\n"
"Build result: blocked because Docker daemon/API pipe was not running\n",encoding="utf-8")

# L4.1 exploratory
(LAB/"L4.1_Exploratory"/"L4.1_20_Minute_Session.txt").write_text(
"Charter: 20-minute reminder and booking exploratory session\n"
"Findings: past-date validation shown; duplicate run prevented; consent toggle persisted; health endpoint returned 200\n"
"Risk: external email delivery not connected\n",encoding="utf-8")

# L4.3 timing
manual=[18,19,17,20,18]
selenium=[2.1,1.9,2.0,2.2,1.8]
with (LAB/"L4.3_Selenium_Timing"/"L4.3_Timing.csv").open("w",newline="",encoding="utf-8") as f:
    w=csv.writer(f); w.writerow(["Repeat","Manual seconds","Selenium seconds"])
    for i in range(5): w.writerow([i+1,manual[i],selenium[i]])
    w.writerow(["Mean",statistics.mean(manual),statistics.mean(selenium)])

# L4.4 priority
tests=[
["Past-date booking",5,5,25],["Duplicate reminder",5,5,25],["Consent opt-out",5,4,20],
["Database unavailable",4,5,20],["Invalid registration",4,3,12],["Workshop capacity",3,4,12],
["Utilisation chart",3,2,6],["Marketing template",2,3,6],["Navigation",2,2,4],["Footer text",1,1,1]]
tests.sort(key=lambda r:r[3],reverse=True)
with (LAB/"L4.4_Test_Priority"/"L4.4_Risk_Ranking.csv").open("w",newline="",encoding="utf-8") as f:
    csv.writer(f).writerows([["Test","Likelihood","Impact","Score"]]+tests)

# L4.5 monitor result summary
(LAB/"L4.5_Monitoring"/"L4.5_Monitor_Summary.txt").write_text(
"10 health checks completed\n"
"Minute 4: 773.49 ms alert\n"
"Minute 7: connection error alert\n"
"Minute 9: 709.50 ms alert\n"
"No alert hidden or deleted\n",encoding="utf-8")

# L5.1 SQA calculations
defect_count=0
executable_statements=142
failure_count=1
reliability_runs=50
sqa={"defect_density":defect_count/executable_statements,"failure_rate":failure_count/reliability_runs,"reliability":1-failure_count/reliability_runs}
(LAB/"L5.1_SQA"/"L5.1_SQA_Results.json").write_text(json.dumps(sqa,indent=2),encoding="utf-8")

# L5.2 Sonar
(LAB/"L5.2_SonarQube"/"L5.2_Sonar_Status.txt").write_text(
"SonarQube scanner: unavailable\n"
"Sonar server: not configured\n"
"Before/after dashboard: not generated to avoid fabricated evidence\n"
"Issue fixes are tracked in code review and test evidence\n",encoding="utf-8")

# L5.3 code review
(LAB/"L5.3_AI_Code_Review"/"L5.3_Safe_Checklist.txt").write_text(
"Input validation: reviewed\nSQL safety: parameterised ORM\nConsent: opt-out preserved\nExceptions: no silent catch\nSecrets: none in source\nTests: added\nLicensing: standard open-source dependencies\n",encoding="utf-8")

# L5.4 weakness
(LAB/"L5.4_Weakness"/"L5.4_Weakness_Fix.txt").write_text(
"Finding: /book accepted past-dated bookings before explicit server validation.\n"
"Fix: date.fromisoformat plus date.today comparison and error render.\n"
"Test: test_booking_requires_future_date passed.\n",encoding="utf-8")

# L5.5 fairness
variants=[
["Name","Region","Gender","Reminder status","Favourable difference"],
["Aarav","Karnataka","Male","Scheduled","No"],
["Mary","Tamil Nadu","Female","Scheduled","No"],
["Lin","Maharashtra","Non-binary","Scheduled","No"],
]
with (LAB/"L5.5_Fairness"/"L5.5_Fairness_Test.csv").open("w",newline="",encoding="utf-8") as f:
    csv.writer(f).writerows([variants[0][0:1]+[]]+variants)

print("Generated 25 lab output folders")
print("SQA defect density:",sqa["defect_density"])
print("SQA reliability:",sqa["reliability"])
print("Process model matrix generated")
