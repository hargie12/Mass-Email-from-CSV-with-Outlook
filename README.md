# Mass Email Sender With Outlook Application

This script sends an email using Microsoft Outlook and a list of email addresses stored in a CSV file.

## Dependencies
- win32com module
- csv module

## How it works
1. The script opens the CSV file located at `C:\Users\filepath` and reads the values in the `email` column using the `csv` module.
2. The values are joined together with a semicolon and stored in the `result` variable.
3. The `win32com` module is used to create a new Outlook email object.
4. The email's subject, recipients, and body are set, and the `result` variable is used to set the BCC field.
5. The email is sent.
