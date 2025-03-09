1️⃣ Add a New Task
To schedule an email task, you need to specify:

--add <interval> (Time interval)
--unit <time_unit> (seconds, minutes, hours, days)
--email-list <file_path> (CSV/XLSX file with email addresses)
--message-file <file_path> (TXT file with email content)
--subject <email_subject>
(Optional) --attachments <file1> <file2> (List of file paths)

========================================
Minute
*******************
python email_automation.py --add 10 --unit minutes --email-list emails.csv --message-file message.txt --subject "Meeting Reminder"
===============================
Hours
**************
python email_automation.py --add 1 --unit hours --email-list contacts.xlsx --message-file content.txt --subject "Monthly Newsletter" --attachments report.pdf image.jpg
====================
2️⃣ List All Scheduled Tasks

python email_automation.py --list

==============
3️⃣ Remove a Scheduled Task
python email_automation.py --remove task_1
===================
pythonw email_automation.py

===================
Stop process

Get-Process pythonw | Stop-Process -Force
