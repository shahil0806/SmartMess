# SmartMess Premium

Premium hostel mess coupon and attendance system built with Flask, MongoDB Atlas and secure one-time QR coupons.

## Phase 1 features

- Registration Number based student verification
- Boy/Girl and Boys Hostel/Girls Hostel classification
- Hostel block and room details
- Separate Boys/Girls dashboard counts
- Student search and Gender/Hostel filters
- Student detail/photo editing
- Active/Inactive student control
- Fixed branch dropdown: AI & ML, Civil (Construction Technology), Electronics (Robotics), Mechanical (CAD/CAM)
- Fixed hostel-block dropdown: BH-1 and BH-2, with dashboard counts and filters

## Phase 2 features

- Permanent MongoDB Atlas storage for students, photos, coupons, settings and scan records
- Registration Number + secure 4-digit PIN student login
- Admin can reset a student's PIN from Edit Student
- Breakfast, Lunch and Dinner Open/Close controls
- Admin-controlled start and end time for every meal
- Student panel shows live meal availability
- Local SQLite fallback when `MONGODB_URI` is not configured

Existing records remain available while using the same SQLite database. For an existing SQLite student without a PIN, the temporary PIN is the last four characters of the Registration Number; admin should reset it from Edit Student.
SQLite records are not automatically copied into a new MongoDB database. Migrate or re-register them before switching if the old database contains real data.

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
$env:MONGODB_URI="your-private-mongodb-atlas-connection-string"
$env:MONGODB_DB="smartmess"
python app.py
```

Open `http://127.0.0.1:5000/student` in the browser.

## Render settings

The included `render.yaml` supplies the build/start commands. Add `ADMIN_PASSWORD`, `GOOGLE_SHEET_URL` and the private `MONGODB_URI` in Render Environment settings. Set `MONGODB_DB` to `smartmess`. Student photos are stored inside MongoDB, so they remain available after Render restarts.

See `PHASE2_SETUP_HINDI.md` for a short step-by-step deployment guide.

## Important

- Never publish the real admin password or secret key in GitHub.
- Never put the real MongoDB connection string in GitHub or send it in chat.
- The coupon remains valid for five minutes and becomes unusable after a successful scan.
- The server uses India time so daily meal records stay on the correct date.
