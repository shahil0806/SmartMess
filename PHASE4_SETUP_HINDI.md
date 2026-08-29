# SmartMess Phase 4 Setup (Hindi)

## Phase 4 में क्या नया है

- नया premium Admin और Student dashboard
- Gender automatic: Boys Hostel = BOY, Girls Hostel = GIRL
- Girls Hostel में block नहीं; Boys Hostel में केवल BH-1/BH-2
- Meal Skip Cancel, Weekly Menu और Complaint/Feedback
- Student Change PIN, Forgot PIN Request और Admin PIN Reset
- Admin login protection, Activity Log और Main Admin Forgot Password
- Expired QR को उसी दिन केवल एक बार 5 मिनट extension
- Installable PWA; browser version भी हमेशा काम करेगा

## Render में केवल एक नया secret जोड़ें

Render > SmartMess > Environment > Add variable:

- Key: `ADMIN_RECOVERY_KEY`
- Value: अपनी कम-से-कम 16 character की private recovery key

Recovery key को GitHub, WhatsApp या screenshot में कभी share न करें। इसके बाद **Save, rebuild, and deploy** दबाएँ।

## जरूरी बातें

- MongoDB Atlas का `MONGODB_URI` वही रहने दें। पुराना data सुरक्षित रहेगा।
- Main Admin दूसरे admins और Student PIN reset कर सकता है।
- Boys/Girls Admin केवल अपने hostel के students देखता है।
- Scanner Operator Boys और Girls दोनों QR scan कर सकता है।
- Offline mode में login, coupon और scanner नहीं चलता; internet आने पर फिर काम करेगा।
