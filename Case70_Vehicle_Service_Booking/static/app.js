async function setConsent(id, checked) {
  await fetch('/api/consent/' + id, {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({marketing_consent:checked})});
}
async function runReminders() {
  const response = await fetch('/api/run-reminders', {method:'POST'});
  const data = await response.json();
  document.getElementById('scheduler-result').textContent = `Processed ${data.processed}; prevented ${data.duplicate_prevented} duplicate reminders.`;
  setTimeout(()=>location.reload(),700);
}
const dateInput=document.getElementById('booking_date');
if(dateInput){const d=new Date();d.setMinutes(d.getMinutes()-d.getTimezoneOffset());dateInput.min=d.toISOString().slice(0,10);}
