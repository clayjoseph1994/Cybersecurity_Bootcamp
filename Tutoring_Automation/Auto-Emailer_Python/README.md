# Trilogy-Tutor-Auto-Emailer
This is a Python program that I created to automate the sending of reminder emails for my tutoring sessions the next day. Students often times don't show up for sessions if they are not reminded. Avoiding this issues requires some outside-the-box thinking. The program works by using my Google APIs in a credentials.json file on my machine that adds support for Google Calendar, Gmail, and Google Sheets. Google Sheets is where the list I keep of students that I am tutoring resides. This roster sheet contains student's names, emails, zoom links, etc. 

The program will check my Google Calendar for events tomorrow matching specific parameters in the description field, and if it finds the ones corresponding to a tutorial session, it then checks for the person's email associated with the event. The next task would be to search the student roster for that email address, specified by column and row. The program then fills out an email template inputting the student's name, date/time of the session, and a few other variables. 

An email is then sent to the students, reminding them of our scheduled tutoring session tomorrow.

This program is set up to be run via a batch file, utilizing Windows Task Scheduler.
