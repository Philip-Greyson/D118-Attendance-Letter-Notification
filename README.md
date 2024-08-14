# # D118-Attendance-Letter-Notification

This is an extremely specific script D118 uses to send an email after a field has been checked for a letter being sent home, so that the administration teams can be notified.

## Overview

The purpose of this script is to send an email to a group when a student in their building has had an attendance letter sent home. When the letter is sent, a box is checked on the student's page, and this script looks at that field, as well as some other custom fields to see if a notification has already been sent about the letter for each semester. If the letter has been sent but the notification has not been sent, it will send an email to the xyz-attendance-notifications group, using the school abbreviation the student is enrolled in. It will then use the ACME PowerSchool API plugin to update the custom field denoting that the notification has been sent so it will not send it more than once per semester per student.

## Requirements

The following Environment Variables must be set on the machine running the script:

- POWERSCHOOL_READ_USER
- POWERSCHOOL_DB_PASSWORD
- POWERSCHOOL_PROD_DB
- POWERSCHOOL_API_ID
- POWERSCHOOL_API_SECRET

These are fairly self explanatory, and just relate to the usernames, passwords, and host IP/URLs for PowerSchool, as well as the API ID and secret you can get from creating a plugin in PowerSchool. If you wish to directly edit the script and include these credentials or to use other environment variable names, you can.

Additionally, the following Python libraries must be installed on the host machine (links to the installation guide):

- [Python-oracledb](https://python-oracledb.readthedocs.io/en/latest/user_guide/installation.html)
- [Python-Google-API](https://github.com/googleapis/google-api-python-client#installation)
- [ACME PowerSchool](https://easyregpro.com/acme/pythonAPI/README.html)

In addition, an OAuth credentials.json file must be in the same directory as the overall script. This is the credentials file you can download from the Google Cloud Developer Console under APIs & Services > Credentials > OAuth 2.0 Client IDs. Download the file and rename it to credentials.json. When the program runs for the first time, it will open a web browser and prompt you to sign into a Google account that has the permissions to send emails. Based on this login it will generate a token.json file that is used for authorization. When the token expires it should auto-renew unless you end the authorization on the account or delete the credentials from the Google Cloud Developer Console. One credentials.json file can be shared across multiple similar scripts if desired.
There are full tutorials on getting these credentials from scratch available online. But as a quickstart, you will need to create a new project in the Google Cloud Developer Console, and follow [these](https://developers.google.com/workspace/guides/create-credentials#desktop-app) instructions to get the OAuth credentials, and then enable APIs in the project (the Admin SDK API is used in this project).

## Customization

This script is an extremely niche and specific one for our use cases, so there is not much you would likely want to customize. You could generalize it and change much of the SQL query to suit fields of your needs while sending specific emails and updating the custom fields, but it is going to be an extensive rewrite of the script.
