# SmartMess Phase 2 — आसान Setup

## इस update में क्या है

- MongoDB Atlas में permanent student, photo, coupon और scan data
- Registration Number + 4-digit PIN login
- Admin से student PIN reset
- Breakfast, Lunch और Dinner Open/Close
- हर meal का start/end time
- Student panel पर live meal status

## Files copy करने के बाद

SmartMess folder में PowerShell खोलें और ये commands एक-एक करके चलाएँ:

```powershell
git status
git add .
git commit -m "Add SmartMess Phase 2 MongoDB and PIN security"
git push origin main
```

## Render में MongoDB जोड़ें

Render Dashboard → SmartMess → Environment → Add Environment Variable:

1. Key: `MONGODB_URI`
   Value: Notepad में रखी पूरी private MongoDB connection string
2. Key: `MONGODB_DB`
   Value: `smartmess`

Save Changes करें। Render का नया deploy पूरा होने दें। MongoDB connection string को GitHub या chat में कभी share न करें।

## सही connection कैसे check करें

Browser में खोलें:

`https://smartmess-cmv8.onrender.com/health`

सही setup पर response में यह दिखेगा:

```json
{"database":"mongodb","service":"SmartMess","status":"ok"}
```

फिर Admin → Register Student में test student बनाएँ और उसका 4-digit PIN याद रखें। Student Panel में Registration Number और उसी PIN से login करें।

