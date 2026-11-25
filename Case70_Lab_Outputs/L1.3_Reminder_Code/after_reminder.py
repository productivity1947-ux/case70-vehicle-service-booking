def schedule_reminder(booking):
    if not isinstance(booking.get("date"), date):
        raise ValueError("Invalid booking date")
    if booking["date"] < date.today():
        raise ValueError("Past booking date")
    if any(r["booking"] == booking["id"] and r["status"] == "Sent" for r in reminders):
        return False
    reminders.append({"booking": booking["id"], "status": "Scheduled"})
    return True
