GuardianComms - Communications Module (Starter)
==============================================

Recommended install location on Windows:
  C:\GoogleSync\Automation\GuardianComms

Files in this folder:
- comm_module.py              (the module)
- Emails\templates\simple_reminder.html  (starter template)
- *.bat launchers (init, preview, send, history)

1) Install dependency (one time):
   pip install openpyxl

2) Initialize your workbook (creates sheets + default template):
   init_comm.bat

3) Add test rows to Clients.xlsx (Clients sheet) with: CauseNo, WardFirst, WardLast, PrimaryEmail, Status=A, ConsentToEmail=Yes

4) Preview (no emails are sent; HTML previews created + logs written):
   preview_comm.bat

5) Send for real (optional; requires Gmail App Password):
   - Create a Gmail App Password (with 2FA).
   - Run send_comm.bat after setting EMAIL_USER and EMAIL_APP_PASSWORD.