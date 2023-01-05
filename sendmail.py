import csv
import win32com.client

# Open the CSV file
with open(r'C:\Users\filepath') as f:
  # Read the CSV file
  reader = csv.DictReader(f)
  # Get the values from the 'column_name' column, in the CSV I have named the column email
  values = [row['email'] for row in reader]

# Join the values together with a semicolon
#Outlook uses the semicolon separator to identify multiple emails
result = ';'.join(values)
print(result)



ol=win32com.client.Dispatch("outlook.application")
olmailitem=0x0 #size of the new email
newmail=ol.CreateItem(olmailitem)
newmail.Subject= 'Testing Mail'
newmail.To=""
newmail.CC=''
newmail.BCC= result
#Type your text into the Body variable
newmail.Body= 'Hello, this is a test email to showcase how we can automate emails using Python and Outlook.'
#You can uncomment the next two lines and specify a file path to add attachment
#attach='C:\\Users\\person\\attachment.xlsx'
#newmail.Attachments.Add(attach)

# To display the mail before sending it
#newmail.Display()

newmail.Send()
