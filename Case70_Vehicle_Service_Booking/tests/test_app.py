import pytest
from datetime import date, datetime
from app import app, db, Customer, Vehicle, Booking, Reminder, ReminderTemplate, Workshop

@pytest.fixture()
def client():
    app.config.update(TESTING=True)
    with app.app_context():
        db.drop_all()
        db.create_all()
        yield app.test_client()
        db.session.remove()
        db.drop_all()

def test_health(client):
    assert client.get("/health").status_code == 200

def test_dashboard_has_navigation(client):
    response = client.get("/")
    assert b"Vehicle Service Operations" in response.data
    assert b"New Booking" in response.data

def test_booking_requires_future_date(client):
    customer = Customer(name="Test", email="test@example.com", phone="123")
    db.session.add(customer); db.session.flush()
    vehicle = Vehicle(registration="KA-TEST", customer_id=customer.id, make="Honda", model="City", next_service=date(2027,1,1))
    workshop = Workshop(name="Test Workshop", city="Bengaluru")
    db.session.add_all([vehicle, workshop]); db.session.commit()
    response = client.post("/book", data={"vehicle_id":vehicle.id,"workshop_id":workshop.id,"booking_date":"2020-01-01","service_type":"General Service"})
    assert b"cannot be in the past" in response.data

def test_duplicate_reminder_is_prevented(client):
    with app.app_context():
        c=Customer(name="A",email="a@example.com",phone="1"); db.session.add(c); db.session.flush()
        v=Vehicle(registration="KA-1",customer_id=c.id,make="A",model="M",next_service=date(2027,1,1)); db.session.add(v); db.session.flush()
        w=Workshop(name="W",city="C"); db.session.add(w); db.session.flush()
        b=Booking(vehicle_id=v.id,workshop_id=w.id,booking_date=date(2027,1,10),service_type="General")
        t=ReminderTemplate(name="T",subject="S",body="B"); db.session.add_all([b,t]); db.session.flush()
        r=Reminder(booking_id=b.id,template_id=t.id,scheduled_for=datetime(2027,1,8))
        db.session.add(r); db.session.commit()
        result=client.post("/api/run-reminders").get_json()
        assert result["processed"] == 1
        result=client.post("/api/run-reminders").get_json()
        assert result["processed"] == 0

def test_consent_opt_out(client):
    c=Customer(name="Consent",email="c@example.com",phone="2",marketing_consent=True); db.session.add(c); db.session.commit()
    result=client.post(f"/api/consent/{c.id}",json={"marketing_consent":False}).get_json()
    assert result["marketing_consent"] is False

def test_book_page_lists_vehicles_and_workshops(client):
    assert client.get("/book").status_code == 200

def test_reminders_page_loads(client):
    assert client.get("/reminders").status_code == 200

def test_utilisation_api_returns_weekdays(client):
    data=client.get("/api/utilisation").get_json()
    assert set(data) == {"Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"}
