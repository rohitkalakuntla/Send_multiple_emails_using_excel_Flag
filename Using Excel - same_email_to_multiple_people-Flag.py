import os
import win32com.client
from datetime import datetime
import json
import pandas as pd
import tkinter as tk
from tkinter import messagebox
import shutil


# Set paths
base_dir = os.getcwd()

# Load json config
config_path = os.path.join(base_dir, "Need_to_update_details.json")
try:
    with open(config_path, "r", encoding="utf-8") as f:
        config = json.load(f)
except Exception as e:
    log_message(f"Error reading config file: {e}")
    exit()

# Access json config values
subject = config.get("subject", "Default Subject")
cc_email = config.get("cc_email", "")
attachments = config.get("attachments", [])
excel_file = config.get("excel_file", "")
excel_path = os.path.join(base_dir, excel_file)
body_file = config.get("body_file", "")
body_path = os.path.join(base_dir, body_file)
log_file = config.get("log_file", "email_log.txt")
log_path = os.path.join(base_dir, log_file) # Single persistent log file
status_log_recipient = config.get("status_log_report_email", "")

# Logging function
def log_message(message):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(log_path, "a", encoding="utf-8") as log_file:
        log_file.write(f"[{timestamp}] {message}\n")
    print(message)

# Log section starting point
with open(log_path, "a", encoding="utf-8") as log_file:
    log_file.write("\n\n")  # Leave two blank lines before this run
    
log_message(f"---- Starting the logging for sending the emails ---------")

# Ask user test mode or not - LD Flag
def ask_test_mode():
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    result = messagebox.askyesno("Run Mode Selection", "Do you want to run in TEST MODE?")
    return result  # True if Yes, False if No

# Test mode usage
TEST_MODE = ask_test_mode()
log_message(f"TEST_MODE is {'ON' if TEST_MODE else 'OFF'}")

# Confirmation if Not test mode - will send emails.
if not TEST_MODE:
    proceed = messagebox.askyesno("Confirm", "You are about to Send Real Emails.\n\n Do you want to continue?")
    if proceed:
        backup_message = messagebox.askyesno("Confirm", "Do you want to take a backup of your existing excel file, before sending emails ?? ")
        if backup_message:
            log_message("Backup of the excel workbook will be taken and stored in same folder path")
            # Generate backup filename with timestamp
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            file_name, file_ext = os.path.splitext(excel_path)
            backup_file = f"{file_name}_backup_{timestamp}{file_ext}"
            # Copy the Excel file
            try:
                shutil.copy(excel_path, backup_file)
                log_message(f"📁 Backup created: {backup_file}")
            except Exception as e:
                log_message(f"⚠️ Failed to create backup: {e}")
    if not proceed:
        log_message("Aborted by user before sending real emails.")
        exit()

# Confirm with user all the details in the json config file.
def confirm_config(config):
    # Prepare the message
    message = (
        f"📋 Please confirm the following configuration:\n\n"
        f"📧 Subject: {subject}\n"
        f"📩 CC Email: {cc_email}\n"
        f"📎 Attachments: {', '.join(attachments) if attachments else 'None'}\n"
        f"📗 Excel File : {excel_file}\n"
        f"🧾 Please make sure the body is updated in : {body_file}\n"
        f"📬 Status Log Report Email: {status_log_recipient}\n\n"
        f"Do you want to continue?"
    )

    root = tk.Tk()
    root.withdraw()
    return messagebox.askyesno("Confirm Email Configuration", message)


# To send the status of this script via email ( Test mode status / real mode status 
def send_log_email_status(recipient_email, log_path, is_test_mode):
    try:
        with open(log_path, "r", encoding="utf-8") as f:
            log_lines = f.readlines()
            # Only get the most recent session (after the last double line break)
            recent_log = []
            for line in reversed(log_lines):
                if line.strip() == "":
                    break
                recent_log.insert(0, line)
            log_content = ''.join(recent_log)

        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.To = recipient_email
        mail.Subject = "📬 Email Automation Log - " + ("TEST MODE" if is_test_mode else "REAL MODE")
        mail.Body = f"Hello,\n\nAttached is the log summary of the last email run.\n\n" + log_content
        mail.Send()
        log_message(f"📤 Log email sent to {recipient_email}")
    except Exception as e:
        log_message(f"❌ Failed to send log email: {e}")

       
# Read html body content
try:
    with open(body_path, "r", encoding="utf-8") as f:
        body_template = f.read()
except Exception as e:
    log_message(f"Error reading email body file: {e}")
    exit()

# calling the confirm_config by the user 

value = confirm_config(config)
if value:
    log_message("User has confirmed all the details")
else:
    exit()

# Read recipient list
recipients = []
skipped_count = 0
log_message(f"***************** ⚠️ Skipping these people as status is mentioned as sent ********** ")
try:
    df = pd.read_excel(excel_path, engine='openpyxl')

    total_rows_in_excel = df.shape[0]

    for idx, row in df.iterrows():
        
        first_name = str(row.get("first_name") or "").strip()
        last_name = str(row.get("last_name") or "").strip()
        email = str(row.get("email") or "").strip()
        

        # Skip if already marked Sent
        if str(row.get("status")).strip().lower().startswith("sent"):
            log_message(f"⚠️ Skipped , due to status was already sent: {row.to_dict()}")
            skipped_count += 1
            continue

        # Skip if missing first name or last name or email
        if first_name.lower() in ("", "nan") or last_name.lower() in ("", "nan") or email.lower() in ("", "nan"):

            log_message(f"⚠️ Skipped row due to missing field(s): {row.to_dict()}")
            skipped_count += 1
            continue
        
        # Add only valid recipients
        recipients.append({
            "first_name": first_name,
            "last_name": last_name,
            "email": email,
            "row_index": idx  # track the Excel row index
        })

    #with open(log_path, "a", encoding="utf-8") as log_file:
        #log_file.write("\n\n")  # Leave two blank lines before this run
    log_message(f"**************************** ")
   
except Exception as e:
    log_message(f"Error reading recipients Excel file: {e}")
    log_message(f"--------------------------------- DONE -----------------------------------")
    exit()

# Send emails

success_count = 0
failed_emails = []


try:
    log_message("Starting multiple email sending process...")

    outlook = win32com.client.Dispatch("Outlook.Application")

    for r in recipients:
        
        idx = r['row_index']  # Use the correct row index
        
        try:
            personalized_body = f"""
            Hi {r['first_name']} {r['last_name']},<br><br>
            {body_template}
            """
    
            mail = outlook.CreateItem(0)
            mail.To = r['email']
            mail.CC = cc_email
            mail.Subject = subject
            mail.HTMLBody = personalized_body
    
            for filename in attachments:
                full_path = os.path.join(base_dir, filename)
                if os.path.exists(full_path):
                    mail.Attachments.Add(full_path)
                else:
                    print(f"Attachment not found: {filename}")
    
            if TEST_MODE:
                log_message(f"🧪 [TEST MODE] Would send email to {r['first_name']} {r['last_name']} ({r['email']})")
            else:
                mail.Send()
                # ✅ Update with date/time
                timestamp = datetime.now().strftime("%Y-%m-%d %I:%M %p")
                df.loc[r['row_index'], 'status'] = f"sent - {timestamp}"
                log_message(f"✅ Email sent to {r['first_name']} {r['last_name']} ({r['email']})")
            
            success_count += 1
            
        except Exception as e:
            log_message(f"❌ Failed to send email to {r['first_name']} {r['last_name']} ({r['email']}): {e}")
            failed_emails.append(f"{r['first_name']} {r['last_name']} ({r['email']})")

    # Saving the excel workbook with the status sent details. It will not update in the test mode.
    if not TEST_MODE:
        df.to_excel(excel_path, index=False, engine="openpyxl")
    else:
        log_message("🧪 [TEST MODE] Skipped writing status updates to Excel.")
    
except Exception as e:
    log_message(f"❌ Global error occurred: {e}")
    
# Final summary
total_expected = len(recipients)

log_message(f"📊 Summary: Total rows in excel: {total_rows_in_excel} | Emails Planned to send: {total_expected} | Successfully Sent: {success_count} | Skipped: {skipped_count}")


if success_count == total_expected:
    if TEST_MODE:
        log_message("Test mode: Emails that it will send.")
    else:
        log_message("✅ All emails sent successfully.")

else:
    log_message("⚠️ Some emails failed. Please check the logs above.")

if failed_emails:
    log_message("❌ Failed email recipients:")
    for entry in failed_emails:
        log_message(f"   - {entry}")

if status_log_recipient:
    send_log_email_status(status_log_recipient, log_path, TEST_MODE)
    
log_message(f"--------------------------------- DONE -----------------------------------")
