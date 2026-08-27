# SmartMess Premium

Premium hostel mess coupon and attendance system built with Flask, SQLite and secure one-time QR coupons.

## Main links

- Student panel: `/student`
- Admin login: `/admin`
- Scanner: `/admin/scanner`
- Health check: `/health`

## Run on Windows

```powershell
py -m venv venv
venv\Scripts\activate
python -m pip install -r requirements.txt
$env:SECRET_KEY="change-this-secret"
$env:ADMIN_PASSWORD="your-new-password"
python app.py
```

Open `http://127.0.0.1:5000/student` in the browser.

## Render settings

The included `render.yaml` supplies the build/start commands. Add `ADMIN_PASSWORD` and `GOOGLE_SHEET_URL` in Render Environment settings. SQLite and uploaded photos need a persistent disk in production; set `DATABASE_PATH` and `PHOTO_DIR` to that disk's paths.

## Important

- Never publish the real admin password or secret key in GitHub.
- The coupon remains valid for five minutes and becomes unusable after a successful scan.
- The server uses India time so daily meal records stay on the correct date.
