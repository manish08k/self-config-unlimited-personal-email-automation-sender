================================================================
  RESUME EMAIL AUTOMATION — SETUP GUIDE
================================================================

WHAT THIS DOES
--------------
Automatically sends your resume PDF to all HR contacts in your
Excel file. Handles 1800+ emails with:
  ✓ Personalised emails (name, company, title per contact)
  ✓ Resume PDF attached to every email
  ✓ Auto-retry on failures
  ✓ Progress saved — resume if interrupted (no re-sends)
  ✓ Rate limiting to avoid Gmail spam blocks
  ✓ Full log of every send attempt
  ✓ Failed emails saved separately for review


FILES IN THIS FOLDER
--------------------
  email_sender.py   — main script (do not edit)
  config.py         — YOUR settings (edit this)
  requirements.txt  — Python packages needed
  README.txt        — this file


STEP 1 — Install Python
-----------------------
Download Python 3.10+ from https://python.org/downloads
During install, check "Add Python to PATH"

Check it works:
  python --version


STEP 2 — Install required packages
-----------------------------------
Open Terminal (Mac/Linux) or Command Prompt (Windows),
navigate to this folder, then run:

  pip install -r requirements.txt


STEP 3 — Set up Gmail App Password
------------------------------------
You CANNOT use your normal Gmail password. You need an App Password.

  1. Go to: https://myaccount.google.com/security
  2. Make sure "2-Step Verification" is ON
  3. Search for "App passwords" at the top
  4. Click "App passwords"
  5. Choose app: Mail | Choose device: Windows/Mac
  6. Click Generate
  7. Copy the 16-character password (e.g. "abcd efgh ijkl mnop")
  8. Paste it into config.py as EMAIL_PASSWORD

⚠ If you don't see "App passwords", 2FA is not enabled. Enable it first.


STEP 4 — Edit config.py
------------------------
Open config.py in any text editor (Notepad, VS Code, etc.)
Fill in:

  EMAIL_ADDRESS  = "yourgmail@gmail.com"
  EMAIL_PASSWORD = "abcd efgh ijkl mnop"   ← App Password
  SENDER_NAME    = "Your Full Name"

  RESUME_PDF_PATH    = "resume.pdf"         ← filename of your resume
  CONTACTS_FILE_PATH = "HR_Lists.xlsx"      ← filename of your contacts file

  Customise EMAIL_BODY_TEMPLATE with your real details (experience, phone, LinkedIn)


STEP 5 — Place your files in this folder
-----------------------------------------
  • Your resume PDF   → rename it to match RESUME_PDF_PATH in config.py
  • Your contacts file → rename it to match CONTACTS_FILE_PATH in config.py

Your contacts file must have these columns (check config.py for exact names):
  Email | Name | Company | Title

Your spreadsheet already has these columns — just make sure the
column name in config.py exactly matches your file header.


STEP 6 — Run the script
------------------------
Open Terminal/Command Prompt in this folder, then:

  python email_sender.py

It will show you a summary and ask you to type YES to confirm.


STEP 7 — Monitor progress
--------------------------
While running, you will see live output like:

  [1/1800] Sending to: akanksha.puri@sourcefuse.com (SourceFuse Technologies)
    ✓ Sent
  [2/1800] Sending to: akanksha.sogani@perennialsys.com (Perennial Systems)
    ✓ Sent
  ...
  Progress: 50/1800 | Success: 49 | Failed: 1 | ETA: 210.0 min

  ✓ Full log saved to: email_log.txt
  ✓ Progress saved to: progress.json


RESUME IF INTERRUPTED
---------------------
If the script stops (power cut, internet drop, etc.), just run it again:

  python email_sender.py

It reads progress.json and skips already-sent emails automatically.
No duplicates will be sent.


GMAIL LIMITS
------------
Gmail free account: 500 emails/day
Gmail Workspace (paid): 2000 emails/day

With 1800 contacts, you may need 4 days on a free Gmail account.
The script handles this automatically — just run it daily.
Progress is saved, so it picks up where it left off.

To send faster: upgrade to Google Workspace (~$6/month).


FAILED EMAILS
-------------
If any emails fail after all retries, they are saved to:
  failed_emails.csv

You can review them and retry separately.


COMMON ERRORS
-------------
"Username and Password not accepted"
  → You used your normal Gmail password. Use the App Password instead.
  → Make sure 2FA is enabled on your Gmail account.

"SMTPAuthenticationError"
  → Same as above — wrong password type.

"Column 'Email' not found"
  → The column name in your Excel doesn't match config.py
  → Open your Excel, check the exact header name, update config.py

"resume.pdf not found"
  → Make sure your PDF filename matches RESUME_PDF_PATH in config.py
  → File must be in the same folder as email_sender.py

