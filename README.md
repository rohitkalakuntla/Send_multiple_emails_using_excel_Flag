# individual-email-to-multiple-people-with-names
Sending Individual emails to multiple people with addressing with their names.

**Goal:** When the Python script is executed, it will send individual emails to all the people in csv file. 

Before Executing, need to make sure and update all these 

**Python Script:**
  Nothing needs to be updated here. Everything is parameterized

**Need_to_update_details.json:**
Need to update 
	"subject",  "cc_email", "attachments", "excel_file", "body_file", "status_log_report_email"

**email_body.html**
  Need to update the body as per the requirement.

**recipients.xlxs workbook **
    Need to add all the list of members we would need to send email. 
    Format needs to be exactly as     first_name,last_name,email

**Attachments**
    The attachments which are needed to send should be in the same folder where the python script is executed. 

**email_log.txt**
    This will provide the logs for the activity with date and time and also summary. 


**Execution Steps**

1. Once the script is executed, it will ask the user ( test mode ( yes or no)
	2. Once yes is selected: It will run in test mode and not send emails. 
	2. Once No is selected: It will ask another confirmation, that it will send emails ( yes or no)
		3. yes: It will show confirmation of all the items in the json file. ( yes or no)
			4. If yes: It will send emails and send the status report via email.
		3. if no: It wil abort and exit the script. 
	4. no: It will break and exit the script.
	
NOTE: IT will also ask if backup is needed or not in the real run mode only. The backup will be saved in the same folder. 
		