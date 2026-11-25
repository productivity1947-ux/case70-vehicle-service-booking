# Case 70 - Vehicle Service Booking

**Student:** Suhas  
**Roll number:** 57  
**Section:** 5B  
**Department:** CSE  
**USN:** ENG24CS1021

This repository contains the reminder-scheduler prototype for a dealership chain with 40 workshops.

## Features
- 40 seeded workshops, 30 synthetic customers and vehicles
- Service booking creation with future-date validation
- Reminder scheduling and duplicate prevention
- Marketing consent opt-in/opt-out
- Weekday utilisation dashboard
- SQLite persistence and health endpoint

## Run
```powershell
py -m pip install -r requirements.txt
py app.py
```
Open `http://127.0.0.1:5000`.

## Test
```powershell
py -m pytest -q --cov=app --cov-report=term-missing --cov-report=html
py selenium_evidence.py
```

## Container
```powershell
docker build -t vehicle-reminder-scheduler .
docker compose up
```

The project includes a Jenkinsfile for push-triggered CI. GitHub, Jenkins, Docker and SonarQube evidence is recorded only where that execution was actually available.
