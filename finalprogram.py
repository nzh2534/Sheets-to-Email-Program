import gspread
from oauth2client.service_account import ServiceAccountCredentials

import email, smtplib, ssl

import time

from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

#Google Sheets Information and Start Up Info
print("\nHello! This is an application that accesses Google Sheets to send mass amounts of emails. This process is meant to be simple and relatively quick.")
print("You can expect each email to take about 4 seconds to send, so Google Sheets with a couple thousand rows may take an hour or so. Considering this, ensure you have")
print("a relatively stable internet connection.")
print("This can include but is not limited to: typos, information input problems, etc...")

first_time = str(input("\nIs this your first time using the app? Or are you seeking help getting started? If so, type y and hit enter, otherwise just hit enter to begin: "))
if first_time == 'y':
    print("\nIn order to ensure this process proceeds without errors, you need to properly format the linked Google Sheet, and other required files.")

    print("\nFor the Google Sheet, there should only be two columns, one titled Email and the other titled Full Name exactly.")
    print("Input the word Email into cell A2, the words Full Name into cell B2, and paste the required full names and emails below them accordingly.")

    print("\nThis application requires that you also have the email you want to send typed out and saved as a .txt file in this application's folder.")
    print("In order to do this, go to your computer's homescreen, right click on your wallpaper, scroll over new, and then click Text Document.")
    print("Type your email on this page exactly as you would in a regular email, starting from the BODY of the email (This application will retain the file's formatting as well).")
    print("Please do not include the email header (Ex: Dear, such and such) as the application pulls this information from Google Sheets")

    print("\nAdditionally, ensure there is an email signature included in the .txt file. This includes Sincerely, [Your Name] AND your base email signature if you have one.")
    print("This program will not retain your base email signature from your email and you need to input it in the .txt file.")
    print("At this time, only input text into the .txt file, as this program can't input an email signature image.")
    print("However, this may be a part of future development if needed.")

    print("\nHave you created and saved a .txt file? If so, type y and hit enter to continue.")
    created_file = str(input("Otherwise, please do as instructed above, merely hit enter to skip this and begin, or quit the program and contact Noah at nhaglund2015@gmail.com for further assistance: "))
    while created_file == 'y':
        print("\nGreat! Now we just have to save the file in the correct location [INPUT MORE HERE]")
        print("If you also have a .pdf file to attach to the emails, put this file here too.")
        print("\nLast, but not least!, if at anytime the program ends because your laptop dies or a similiar problem occurs, merely start the program again and")
        print("ensure the people who were sent emails (check your email sent folder) are removed from the Google Sheet accordingly.")
        program_continue = str(input("\nPlease type y and hit enter to begin the program, otherwise just hit enter to quit the program: "))
        if program_continue == 'y':
            first_time = 'y'
            created_file = ' '
        elif program_continue != 'y':
            quit()

#API and Sheet Input
creds_json = input("\nPlease generate a Google Drive API Service Account JSON Credentials File. Input the file name (something.json): ")
sheet_name = input("Input the name of the Google Sheet (the exact name as it appears in sheets, no file extension): ")
print("\nNOTE: Use the first tab only of the indicated sheet; Title cell A1 'Email' and cell B2 'Full Name' exactly.")
print("Under cell A1 input all recipient emails and then their corresponding full names under cell B2")

#Access Google Sheets
scope = ["https://spreadsheets.google.com/feeds","https://www.googleapis.com/auth/spreadsheets","https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name(creds_json,scope)
client = gspread.authorize(creds)
sheet = client.open(sheet_name).sheet1
lyst1 = sheet.get_all_records()

#Email Input
sender_email = str(input("\nType the email address of the sender here: "))
password = str(input("Type the email's application password here. (Enable this by turning on Two Step Verification in Gmail Settings and then enabling an App Password): "))
email_subject = str(input("\nType the subject of the email here: "))
email_greeting = str(input("Type the email greeting here (ex: Dear or Hi): "))

#Email Body from Text File
require_body = ''
while require_body != 'q':
    try:
        textfile_name = str(input("\nPlease input the name of a text file for the body of the email followed by .txt: "))
        with open (textfile_name, "r+") as f:
            email_body = f.read()
            break
    except FileNotFoundError:
        print("\nLooks like the file can't be found. Don't forget to put .txt after the file name and make sure the file is listed in the application's folder.")
        print("Please try to enter the name of the file again. Otherwise, exit the program and contact the system administrator Noah at nhaglund2015@gmail.com")

#See if Email requires PDF Attachement
if_text_file = str(input("Does this email require a pdf attachment? If it doesn't, please type x and hit enter. Otherwise, just press enter to input a pdf file: "))

#If there is an attachment, open PDF file in binary mode
while if_text_file != 'x':
    filename = str(input("Input the attachment file's name followed by .pdf (Ex: attachment.pdf): "))
    try:
        with open(filename, "rb") as attachment:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())
        # Encode file in ASCII characters to send by email
        encoders.encode_base64(part)
        # Add header as key/value pair to attachment part
        part.add_header(
            "Content-Disposition",
            f"attachment; filename= {filename}",
        )
        break

    except:
        print("\nThe pdf file can't be found. Make sure the file is in your application folder and was input correctly")
        print("Please try to enter the name of the file again. Otherwise, exit the program and contact the system administrator Noah at nhaglund2015@gmail.com")

if if_text_file == 'x':
    part = ''

#Email Loop Function:
def email_loop (sender_email,email_subject,lyst1,email_body,part,password):
    global x
    global sheet
    try:
        message = MIMEMultipart("alternative")
        message["From"] = sender_email
        message["Subject"] = email_subject
        y = lyst1[x]
        receiver_email = str(y["Email"])
        full_name = str(y["Full Name"])
        message["To"] = receiver_email

        text = """\
{greeting} {name},

{body}""".format(name=full_name, body=email_body, greeting=email_greeting)

        html = """\
        <html>
          <body>
            <p>{greeting} {name},<br>
               {body}<br>
            </p>
          </body>
        </html>
        """.format(name=full_name, body=email_body, greeting=email_greeting)

        # Turn these into plain/html MIMEText objects
        part1 = MIMEText(text, "plain")
        part2 = MIMEText(html, "html")

        # Add HTML/plain-text parts to MIMEMultipart message
        # The email client will try to render the last part first
        message.attach(part1)
        message.attach(part2)

        # Add attachment to message and convert message to string
        if if_text_file != 'x':
            message.attach(part)

        # Log in to server using secure context and send email
        context = ssl.create_default_context()

        time.sleep(1)

        with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
            server.login(sender_email, password)
            server.sendmail(sender_email, receiver_email, message.as_string())

    except smtplib.SMTPServerDisconnected:
        time.sleep(110)
        sheet = client.open("sfdata").sheet1
        lyst1 = sheet.get_all_records()
        email_loop(sender_email, email_subject, lyst1, email_body, part, password)

    except smtplib.SMTPDataError:
        time.sleep(200)
        sheet = client.open("sfdata").sheet1
        lyst1 = sheet.get_all_records()
        email_loop(sender_email, email_subject, lyst1, email_body, part, password)

    except smtplib.SMTPRecipientsRefused:
        print("\nError on row", x + 2, "of your Google Sheet")
        print("The listed email recipient is incorrect or there is a black Google Sheet row.")
        print("Please correct this by updating the email or deleting the row and then resume the program.")
        send_error = str(input("\nPress enter to resume the program: "))
        if send_error != 'quit':
            sheet = client.open("sfdata").sheet1
            lyst1 = sheet.get_all_records()
            email_loop(sender_email, email_subject, lyst1, email_body, part, password)
        elif send_error == 'quit':
            quit()

    except smtplib.SMTPAuthenticationError:
        print("\nYou entered your email or password wrong, please restart the program and enter the sender email and password correctly.")
        final_response = str(input("Please press Enter to now exit the program."))
        quit()

    else:
        x = x + 1

#Optional Email Test Before Official Email Loop
print("\nThis program recommends you send a test email first, especially to check the HTML Format.")
email_preview = str(input("Hit ENTER to continue with a test email, otherwise type 'continue' and hit enter to continue: "))
while email_preview != "continue":

    test_recipient = str(input("Please input the email test recipient: "))

    with open(textfile_name, "r+") as f:
        email_body = f.read()

    try:
        message = MIMEMultipart("alternative")
        message["From"] = sender_email
        message["Subject"] = email_subject
        full_name = "Test Subject"
        message["To"] = test_recipient

        text = """\
    {greeting} {name},
    
    {body}""".format(name=full_name, body=email_body, greeting=email_greeting)

        html = """\
        <html>
          <body>
            <p>{greeting} {name},<br>
               {body}<br>
            </p>
          </body>
        </html>
        """.format(name=full_name, body=email_body, greeting=email_greeting)

        # Turn these into plain/html MIMEText objects
        part1 = MIMEText(text, "plain")
        part2 = MIMEText(html, "html")

        # Add HTML/plain-text parts to MIMEMultipart message
        # The email client will try to render the last part first
        message.attach(part1)
        message.attach(part2)

        # Add attachment to message and convert message to string
        if if_text_file != 'x':
            message.attach(part)

        # Log in to server using secure context and send email
        context = ssl.create_default_context()

        with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
            server.login(sender_email, password)
            server.sendmail(sender_email, test_recipient, message.as_string())

    except smtplib.SMTPRecipientsRefused:
        print("The listed email recipient is incorrect.")
        test_recipient = str(input("Please correct this by updating the email."))
        continue

    except smtplib.SMTPAuthenticationError:
        print("\nYou entered your email or password wrong, please restart the program and enter the sender email and password correctly.")
        final_response = str(input("Please press Enter to now exit the program."))
        quit()

    else:
        print("\nWas the test email formatted as desired?")
        email_preview = str(input("If not, edit the text file, press Enter to re-enter the email test recipient, and have another test sent. Otherwise, type 'continue' to start the official program: "))



#Read txt file one last time to ensure official email blast retains final formatting
with open (textfile_name, "r+") as f:
    email_body = f.read()

#Official Loop, will send email from Google Sheets
x = 0
while x != len(lyst1):
    email_loop(sender_email, email_subject, lyst1, email_body, part, password)


print("\nCompleted! Please see your email's sent folder to ensure that the program sent the required emails. ")

final_response = str(input("Please press Enter to now exit the program."))