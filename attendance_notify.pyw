"""Script to send notifications to certain emails when an attendance letter has been sent.

https://github.com/Philip-Greyson/D118-Attendance-Letter-Notification

Needs the google-api-python-client, google-auth-httplib2 and the google-auth-oauthlib:
pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib
also needs oracledb: pip install oracledb --upgrade
"""

import base64
import json
import os  # needed for environement variable reading
from datetime import *

# importing module
import acme_powerschool
import oracledb  # needed for connection to PowerSchool server (ordcle database)
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from email.message import EmailMessage

# setup db connection
DB_UN = os.environ.get('POWERSCHOOL_READ_USER')  # username for read-only database user
DB_PW = os.environ.get('POWERSCHOOL_DB_PASSWORD')  # the password for the database account
DB_CS = os.environ.get('POWERSCHOOL_PROD_DB')  # the IP address, port, and database name to connect to
print(f'DBUG: Database Username: {DB_UN} |Password: {DB_PW} |Server: {DB_CS}')  # debug so we can see where oracle is trying to connect to/with

d118_client_id = os.environ.get("POWERSCHOOL_API_ID")
d118_client_secret = os.environ.get("POWERSCHOOL_API_SECRET")

# Google API Scopes that will be used. If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/gmail.compose']

EMAIL_GROUP_SUFFIX = '-attendance-notifications@d118.org'  # a suffix to be appended to the school abbreviations and will make up the group email


if __name__ == '__main__':
    with open('attendance_notification_log.txt', 'w') as log:
        startTime = datetime.now()
        startTime = startTime.strftime('%H:%M:%S')
        print(f'INFO: Execution started at {startTime}')
        print(f'INFO: Execution started at {startTime}', file=log)
        creds = None
        # The file token.json stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.
        if os.path.exists('token.json'):
            creds = Credentials.from_authorized_user_file('token.json', SCOPES)
        # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'credentials.json', SCOPES)
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open('token.json', 'w') as token:
                token.write(creds.to_json())
        service = build('gmail', 'v1', credentials=creds)

        ps = acme_powerschool.api('d118-powerschool.info', client_id=d118_client_id, client_secret=d118_client_secret) # create ps object via the API to do requests on

        # create the connecton to the PowerSchool database
        with oracledb.connect(user=DB_UN, password=DB_PW, dsn=DB_CS) as con:
            with con.cursor() as cur:  # start an entry cursor
                print(f'INFO: Connection established to PS database on version: {con.version}')
                print(f'INFO: Connection established to PS database on version: {con.version}', file=log)
                cur.execute('SELECT stu.dcid, stu.first_name, stu.last_name, stu.grade_level, stu.schoolid, absent.chronicletter_sem1_sent, absent.chronicletter_sem2_sent, absent.chronicletter_sem1_notified, absent.chronicletter_sem2_notified, schools.abbreviation, stufields.custom_counselor_email, stufields.custom_deans_house_email, stufields.custom_social_email, stufields.custom_psych_email, stu.student_number FROM students stu LEFT JOIN u_chronicabsenteeism absent ON stu.dcid = absent.studentsdcid LEFT JOIN schools ON stu.schoolid = schools.school_number LEFT JOIN u_studentsuserfields stufields ON stu.dcid = stufields.studentsdcid WHERE absent.chronicletter_sem1_sent = 1 OR absent.chronicletter_sem2_sent = 1')
                students = cur.fetchall()
                for student in students:
                    try:
                        # print(student)  # debug
                        dcid = int(student[0])
                        stuNum = int(student[14])
                        firstName = str(student[1])
                        lastName = str(student[2])
                        grade = int(student[3])
                        school = int(student[4])
                        schoolAbbrev = str(student[9])
                        semester1LetterSent = True if student[5] == 1 else False
                        semester2LetterSent = True if student[6] == 1 else False
                        semester1Notified = True if student[7] == 1 else False
                        semester2Notified = True if student[8] == 1 else False
                        guidanceCounselorEmail = str(student[10])
                        deansEmail = str(student[11])
                        socialWorkerEmail = str(student[12])
                        psychologistEmail = str(student[13])
                        print(f'DBUG: Starting student {stuNum} at building {school}, 1st semester letter {semester1LetterSent} and notified {semester1Notified}, 2nd semester letter {semester2LetterSent} and notified {semester2Notified}')
                        print(f'DBUG: Starting student {stuNum} at building {school}, 1st semester letter {semester1LetterSent} and notified {semester1Notified}, 2nd semester letter {semester2LetterSent} and notified {semester2Notified}', file=log)
                        # start processing if we need to send emails
                        if (semester1LetterSent and not semester1Notified):  # if the letter has been sent for semester 1 but we havent sent the notification
                            toEmail = schoolAbbrev + EMAIL_GROUP_SUFFIX
                            if school == 5:
                                toEmail = f'{toEmail},{guidanceCounselorEmail},{deansEmail},{socialWorkerEmail},{psychologistEmail}'  # if we are at the high school, need to add their specific student service team
                            # print(toEmail)  # debug
                            print(f'INFO: {firstName} {lastName} in grade {grade} at building {school} has not had the notification sent for semester 1, will send to {toEmail}')
                            print(f'INFO: {firstName} {lastName} in grade {grade} at building {school} has not had the notification sent for semester 1, will send to {toEmail}', file=log)
                            try:
                                mime_message = EmailMessage()  # create an email message object
                                # define headers
                                mime_message['To'] = toEmail
                                mime_message['Subject'] = f'[TEST]Chronic Absence Letter Sent to {stuNum} - {firstName} {lastName} for semester 1'  # subject line of the email
                                mime_message.set_content(f'[TEST]This email is to inform you that a Chronic Absence Letter has been sent home to {stuNum} - {firstName} {lastName} for semester 1. Please follow up with building administration if more details are needed.')  # body of the email
                                # encoded message
                                encoded_message = base64.urlsafe_b64encode(mime_message.as_bytes()).decode()
                                create_message = {'raw': encoded_message}
                                send_message = (service.users().messages().send(userId="me", body=create_message).execute())
                                print(f'DBUG: Email sent, message ID: {send_message["id"]}') # print out resulting message Id
                                print(f'DBUG: Email sent, message ID: {send_message["id"]}', file=log)
                                # update the notification box via API. See https://groups.io/g/PSUG/message/197045 for details on updating fields in extension tables
                                data = {
                                    'students' : {
                                        'student': [{
                                            '@extensions': 'u_chronicabsenteeism',
                                            'id' : str(dcid),
                                            'client_uid' : str(dcid),
                                            'action' : 'UPDATE',
                                            '_extension_data': {
                                                '_table_extension': [{
                                                    'name': 'u_chronicabsenteeism',
                                                    '_field': [{
                                                        'name': 'chronicletter_sem1_notified',
                                                        'value': True
                                                    }]
                                                }]
                                            }
                                        }]
                                    }
                                }
                                result = ps.post('ws/v1/student?extensions=u_chronicabsenteeism', data=json.dumps(data))
                                statusCode = result.json().get('results').get('result').get('status')
                                if statusCode != 'SUCCESS':
                                    print(f'ERROR: Could not update semester 1 notified field for student {stuNum}, status {result.json().get('results').get('result')}')
                                    print(f'ERROR: Could not update semester 1 notified field for student {stuNum}, status {result.json().get('results').get('result')}', file=log)
                                else:
                                    print(f'INFO: Successfully marked {stuNum} as having sent the notification')
                                    print(f'INFO: Successfully marked {stuNum} as having sent the notification', file=log)
                            except HttpError as er:   # catch Google API http errors, get the specific message and reason from them for better logging
                                status = er.status_code
                                details = er.error_details[0]  # error_details returns a list with a dict inside of it, just strip it to the first dict
                                print(f'ERROR {status} from Google API while sending semester 1 email: {details["message"]}. Reason: {details["reason"]}')
                                print(f'ERROR {status} from Google API while sending semester 1 email: {details["message"]}. Reason: {details["reason"]}', file=log)
                            except Exception as er:
                                print(f'ERROR while sending or updating semester 1 notification for student {stuNum}: {er}')
                                print(f'ERROR while sending or updating semester 1 notification for student {stuNum}: {er}', file=log)

                        if (semester2LetterSent and not semester2Notified):  # if the letter has been sent for semester 2 but we havent sent the notification
                            toEmail = schoolAbbrev + EMAIL_GROUP_SUFFIX
                            if school == 5:
                                toEmail = f'{toEmail},{guidanceCounselorEmail},{deansEmail},{socialWorkerEmail},{psychologistEmail}'  # if we are at the high school, need to add their specific student service team
                            # print(toEmail)  # debug
                            print(f'INFO: {firstName} {lastName} in grade {grade} at building {school} has not had the notification sent for semester 2, will send to {toEmail}')
                            print(f'INFO: {firstName} {lastName} in grade {grade} at building {school} has not had the notification sent for semester 2, will send to {toEmail}', file=log)
                            try:
                                mime_message = EmailMessage()  # create an email message object
                                # define headers
                                mime_message['To'] = toEmail
                                mime_message['Subject'] = f'[TEST]Chronic Absence Letter Sent to {stuNum} - {firstName} {lastName} for semester 2'  # subject line of the email
                                mime_message.set_content(f'[TEST]This email is to inform you that a Chronic Absence Letter has been sent home to {stuNum} - {firstName} {lastName} for semester 2. Please follow up with building administration if more details are needed.')  # body of the email
                                # encoded message
                                encoded_message = base64.urlsafe_b64encode(mime_message.as_bytes()).decode()
                                create_message = {'raw': encoded_message}
                                send_message = (service.users().messages().send(userId="me", body=create_message).execute())
                                print(f'DBUG: Email sent, message ID: {send_message["id"]}') # print out resulting message Id
                                print(f'DBUG: Email sent, message ID: {send_message["id"]}', file=log)
                                # update the notification box via API. See https://groups.io/g/PSUG/message/197045 for details on updating fields in extension tables
                                data = {
                                    'students' : {
                                        'student': [{
                                            '@extensions': 'u_chronicabsenteeism',
                                            'id' : str(dcid),
                                            'client_uid' : str(dcid),
                                            'action' : 'UPDATE',
                                            '_extension_data': {
                                                '_table_extension': [{
                                                    'name': 'u_chronicabsenteeism',
                                                    '_field': [{
                                                        'name': 'chronicletter_sem2_notified',
                                                        'value': True
                                                    }]
                                                }]
                                            }
                                        }]
                                    }
                                }
                                result = ps.post('ws/v1/student?extensions=u_chronicabsenteeism', data=json.dumps(data))
                                statusCode = result.json().get('results').get('result').get('status')
                                if statusCode != 'SUCCESS':
                                    print(f'ERROR: Could not update semester 2 notified field for student {stuNum}, status {result.json().get('results').get('result')}')
                                    print(f'ERROR: Could not update semester 2 notified field for student {stuNum}, status {result.json().get('results').get('result')}', file=log)
                                else:
                                    print(f'INFO: Successfully marked {stuNum} as having sent the notification')
                                    print(f'INFO: Successfully marked {stuNum} as having sent the notification', file=log)
                            except HttpError as er:   # catch Google API http errors, get the specific message and reason from them for better logging
                                status = er.status_code
                                details = er.error_details[0]  # error_details returns a list with a dict inside of it, just strip it to the first dict
                                print(f'ERROR {status} from Google API while sending semester 2 email: {details["message"]}. Reason: {details["reason"]}')
                                print(f'ERROR {status} from Google API while sending semester 2 email: {details["message"]}. Reason: {details["reason"]}', file=log)
                            except Exception as er:
                                print(f'ERROR while sending or updating semester 2 notification for student {stuNum}: {er}')
                                print(f'ERROR while sending or updating semester 2 notification for student {stuNum}: {er}', file=log)
                    except Exception as er:
                        print(f'ERROR while processing overall student {stuNum}')
                        print(f'ERROR while processing overall student {stuNum}', file=log)