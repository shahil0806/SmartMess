import requests
from datetime import datetime

URL = "https://script.google.com/macros/s/AKfycbzRjb3Xh_O95l9N18VJKhUVs6-4qb99Ybr60zmyxIp-amMOtBPnFzArpyJo2tlyl1zE/exec"

data = {
    "date": datetime.now().strftime("%d-%m-%Y"),
    "time": datetime.now().strftime("%H:%M:%S"),
    "name": "Shahil Raj",
    "roll": "101",
    "branch": "AI & ML",
    "room": "A-12",
    "meal": "LUNCH",
    "status": "TEST"
}

try:
    response = requests.post(URL, json=data, timeout=15)

    print("Status Code:", response.status_code)
    print("Response:", response.text)

except Exception as e:
    print("ERROR:", e)