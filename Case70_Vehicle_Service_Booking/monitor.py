import time, json, requests
from datetime import datetime
from pathlib import Path

url = "http://127.0.0.1:5000/health"
limit_ms = 500
log = Path("monitor_log.csv")
if not log.exists():
    log.write_text("timestamp,status_code,response_ms,alert\n", encoding="utf-8")

for minute in range(10):
    started=time.perf_counter()
    try:
        response=requests.get(url,timeout=3)
        elapsed=(time.perf_counter()-started)*1000
        alert=elapsed > limit_ms
    except Exception as exc:
        elapsed=-1
        alert=True
        response=None
    status=response.status_code if response else "ERROR"
    with log.open("a",encoding="utf-8") as f:
        f.write(f"{datetime.now().isoformat()},{status},{elapsed:.2f},{alert}\n")
    print(json.dumps({"minute":minute+1,"status":status,"response_ms":round(elapsed,2),"alert":alert}))
    if minute < 9:
        time.sleep(60)
