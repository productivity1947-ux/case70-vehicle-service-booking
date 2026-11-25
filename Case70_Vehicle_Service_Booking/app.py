from flask import Flask, render_template, request, jsonify, redirect, url_for
from flask_sqlalchemy import SQLAlchemy
from datetime import datetime, date, timedelta
from pathlib import Path

app = Flask(__name__)
app.config["SQLALCHEMY_DATABASE_URI"] = "sqlite:///service_booking.db"
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
db = SQLAlchemy(app)


class Workshop(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False)
    city = db.Column(db.String(80), nullable=False)
    capacity = db.Column(db.Integer, default=8)


class Customer(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False)
    email = db.Column(db.String(120), nullable=False)
    phone = db.Column(db.String(20), nullable=False)
    marketing_consent = db.Column(db.Boolean, default=False)


class Vehicle(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    registration = db.Column(db.String(20), unique=True, nullable=False)
    customer_id = db.Column(db.Integer, db.ForeignKey("customer.id"), nullable=False)
    make = db.Column(db.String(50), nullable=False)
    model = db.Column(db.String(50), nullable=False)
    next_service = db.Column(db.Date, nullable=False)


class Booking(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    vehicle_id = db.Column(db.Integer, db.ForeignKey("vehicle.id"), nullable=False)
    workshop_id = db.Column(db.Integer, db.ForeignKey("workshop.id"), nullable=False)
    booking_date = db.Column(db.Date, nullable=False)
    service_type = db.Column(db.String(80), nullable=False)
    status = db.Column(db.String(20), default="Scheduled")


class ReminderTemplate(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(80), unique=True, nullable=False)
    subject = db.Column(db.String(120), nullable=False)
    body = db.Column(db.Text, nullable=False)


class Reminder(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    booking_id = db.Column(db.Integer, db.ForeignKey("booking.id"), nullable=False)
    template_id = db.Column(db.Integer, db.ForeignKey("reminder_template.id"), nullable=False)
    scheduled_for = db.Column(db.DateTime, nullable=False)
    status = db.Column(db.String(20), default="Scheduled")
    sent_at = db.Column(db.DateTime)


def seed_database():
    if Workshop.query.first():
        return
    cities = ["Bengaluru", "Mysuru", "Mangaluru", "Hubballi", "Belagavi"]
    for i in range(1, 41):
        db.session.add(Workshop(name=f"Workshop-{i:02d}", city=cities[(i - 1) % len(cities)],
                                capacity=6 + (i % 5)))
    db.session.flush()
    customers = []
    for i in range(1, 31):
        customer = Customer(name=f"Customer {i:02d}", email=f"customer{i:02d}@example.com",
                            phone=f"+91 90000{i:05d}"[-15:], marketing_consent=(i % 4 != 0))
        db.session.add(customer)
        customers.append(customer)
    db.session.flush()
    makes = ["Toyota", "Honda", "Hyundai", "Maruti", "Tata"]
    for i, customer in enumerate(customers, 1):
        vehicle = Vehicle(registration=f"KA-{i:02d}-AB-{1000+i}",
                          customer_id=customer.id, make=makes[(i - 1) % 5],
                          model=f"Model-{i % 8}", next_service=date.today() + timedelta(days=(i % 21)))
        db.session.add(vehicle)
    db.session.flush()
    db.session.add_all([
        ReminderTemplate(name="Service Due", subject="Your vehicle service is due",
                         body="Dear customer, please schedule your vehicle service."),
        ReminderTemplate(name="Booking Reminder", subject="Upcoming service booking",
                         body="Dear customer, your service appointment is approaching."),
        ReminderTemplate(name="Marketing Offer", subject="Weekday service offer",
                         body="Enjoy priority service on available weekday slots."),
    ])
    db.session.flush()
    vehicles = Vehicle.query.all()
    workshops = Workshop.query.all()
    for i, vehicle in enumerate(vehicles):
        booking = Booking(vehicle_id=vehicle.id,
                          workshop_id=workshops[i % 40].id,
                          booking_date=date.today() + timedelta(days=(i % 14) - 3),
                          service_type=["General Service", "Oil Change", "Brake Inspection"][i % 3],
                          status=["Scheduled", "Completed", "Cancelled"][i % 7 == 0])
        db.session.add(booking)
    db.session.flush()
    template = ReminderTemplate.query.filter_by(name="Booking Reminder").first()
    for i, booking in enumerate(Booking.query.all()):
        db.session.add(Reminder(booking_id=booking.id, template_id=template.id,
                                scheduled_for=datetime.combine(booking.booking_date, datetime.min.time()) - timedelta(days=2),
                                status="Sent" if i % 3 == 0 else "Scheduled"))
    db.session.commit()


with app.app_context():
    db.create_all()
    seed_database()


@app.route("/")
def dashboard():
    return render_template("dashboard.html",
                           workshops=Workshop.query.count(),
                           vehicles=Vehicle.query.count(),
                           bookings=Booking.query.count(),
                           reminders=Reminder.query.count(),
                           weekday_fill=weekday_utilisation(),
                           recent=Booking.query.order_by(Booking.id.desc()).limit(8).all(),
                           customers=Customer.query.all())


@app.route("/book", methods=["GET", "POST"])
def book():
    if request.method == "POST":
        vehicle_id = int(request.form["vehicle_id"])
        workshop_id = int(request.form["workshop_id"])
        booking_date = date.fromisoformat(request.form["booking_date"])
        if booking_date < date.today():
            return render_template("book.html", vehicles=Vehicle.query.all(), workshops=Workshop.query.all(),
                                   error="Booking date cannot be in the past.")
        booking = Booking(vehicle_id=vehicle_id, workshop_id=workshop_id,
                          booking_date=booking_date,
                          service_type=request.form["service_type"])
        db.session.add(booking)
        db.session.flush()
        template = ReminderTemplate.query.filter_by(name="Booking Reminder").first()
        db.session.add(Reminder(booking_id=booking.id, template_id=template.id,
                                scheduled_for=datetime.combine(booking_date, datetime.min.time()) - timedelta(days=2)))
        db.session.commit()
        return redirect(url_for("dashboard"))
    return render_template("book.html", vehicles=Vehicle.query.all(), workshops=Workshop.query.all())


@app.route("/reminders")
def reminders():
    rows = db.session.query(Reminder, Booking, Vehicle, Customer).join(
        Booking, Reminder.booking_id == Booking.id).join(
        Vehicle, Booking.vehicle_id == Vehicle.id).join(
        Customer, Vehicle.customer_id == Customer.id).all()
    return render_template("reminders.html", rows=rows)


@app.route("/api/run-reminders", methods=["POST"])
def run_reminders():
    now = datetime.now()
    processed = 0
    duplicate_prevented = 0
    for reminder in Reminder.query.filter(Reminder.status.in_(["Scheduled", "Due"])).all():
        booking = db.session.get(Booking, reminder.booking_id)
        vehicle = db.session.get(Vehicle, booking.vehicle_id)
        customer = db.session.get(Customer, vehicle.customer_id)
        if reminder.status == "Sent":
            duplicate_prevented += 1
            continue
        existing = Reminder.query.filter_by(booking_id=booking.id, status="Sent").count()
        if existing:
            duplicate_prevented += 1
            continue
        reminder.status = "Sent"
        reminder.sent_at = now
        processed += 1
    db.session.commit()
    return jsonify({"processed": processed, "duplicate_prevented": duplicate_prevented})


@app.route("/api/consent/<int:customer_id>", methods=["POST"])
def set_consent(customer_id):
    customer = db.session.get(Customer, customer_id)
    customer.marketing_consent = request.json.get("marketing_consent", False)
    db.session.commit()
    return jsonify({"customer_id": customer_id, "marketing_consent": customer.marketing_consent})


@app.route("/api/utilisation")
def utilisation_api():
    return jsonify(weekday_utilisation())


def weekday_utilisation():
    labels = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    counts = {label: 0 for label in labels}
    for booking in Booking.query.filter_by(status="Scheduled").all():
        counts[labels[booking.booking_date.weekday()]] += 1
    return counts


@app.route("/health")
def health():
    return jsonify({"status": "healthy", "database": "sqlite", "workshops": Workshop.query.count()})


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
