PATH_TO_MSG_TEMPLATE = (r"C:\Users\puzzl\Trilogy-Tutor-Auto-Emailer\msg_template.html") 
#String - file path to msg_template.html (modify as needed)

SUBJECT_TEMPLATE = 'Cybersecurity Boot Camp - Tutorial Confirmation <$date ${starttime}-${endtime}>'
#SUBJECT_TEMPLATE = #String - desired subject line containing the placeholder variables $date, $starttime, and $endtime
#^^The braces prevent confusion with the hyphen and angle brackets. Use as needed.

TUTOR_EMAIL = "mytestemail@gmail.com"

TUTOR_SENDER = '(Joseph Clay) <myemail@gmail.com>'
#TUTOR_SENDER = '(John Smith) <john.smith@gmail.com>'
#^^THE PARENS ARE NECESSARY FOR SOME DISPLAY NAMES TO SHOW!

EVENT_DESCRIPTION = 'Boot Camp Tutorial Session'

SHEET_NAME = 'my_sheet_ID'
#String - the Sheet ID of the Google Sheet containing your students' data (see here: https://help.form.io/assets/img/googlesheet/googlesheet-spreadsheet.png)

RANGE_NAME = 'Student Roster!C2:G'
#String - the Excel-formatted Worksheet name & Range of the data containing the students' names, emails, timezones, and Zoom links (Example: 'Student Roster!C5:G')

NAME_COLUMN = 0
#Int - number specifying the index of the Student Name column in the range specified in RANGE_NAME

EMAIL_COLUMN = 1
#Int - number specifying the index of the Student Email column in the range specified in RANGE_NAME

TZ_COLUMN = 3
#Int - number specifying the index of the Student Timezone column in the range specified in RANGE_NAME

ZOOM_COLUMN = 4
#See guide for variables NAME_COLUMN-ZOOM_COLUMN here: https://imgur.com/a/q4LClAE

ZOOM_PASSWORD = "Not required- the link does not require a password"
#String - the password for your Zoom sessions (by default, Zoom sets same password for all meetings)

TUTOR_TIMEZONE = "US/Eastern"
#String - pytz recognizeable timezone of tutor (for list of valid values, see here: https://gist.github.com/heyalexej/8bf688fd67d7199be4a1682b3eec7568)

RESENDING_EMAILS = False 
#Boolean - Flag indicating whether script should resend emails that have already been sent for next calendar day

NIGHTLY_CUTOFF = '06:00' #String - Time after which the script should send confirmation emails for the next calendar.
#^^For the night owls that want to send emails after midnight for the 'same' day.

IS_IN_TEST_MODE = False 
#Boolean - Flag indicating whether script is in test mode (Emails will send to TEST_EMAIL instead of student and Central Support. 
#Emails will be resent regardless of RESENDING_EMAILS.)

TEST_EMAIL = "mytestemail@gmail.com"
#String - Email address for receiving test emails. Change the value of IS_IN_TEST_MODE to toggle test mode.
