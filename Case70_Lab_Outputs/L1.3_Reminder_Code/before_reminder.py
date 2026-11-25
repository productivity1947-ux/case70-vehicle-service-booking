def schedule_reminder(booking):
    if booking["date"] >= date.today():
        reminders.append({"booking": booking["id"], "status": "Scheduled"})
