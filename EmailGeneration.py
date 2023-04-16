from PyPDF2 import PdfReader
from os import listdir
from os.path import isfile, join
import requests
import json
import imaplib
import email
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

input_path = "C:\\Users\\vivan\\OneDrive\\Documents\\Course Folder\\CS 4375\\001"

onlyfiles = [f for f in listdir(input_path) if isfile(join(input_path, f))]
print(onlyfiles)

host = 'outlook.com'
username = 'jakekawalski@outlook.com'
password = 'Pass123word'

profession = "Professor"
name = "Karen Mazidi"

def extract_doc(input_dir):
    reader = PdfReader(input_dir)

    page_num = len(reader.pages)
    
    for i in range(page_num):
        page = reader.pages[i]
        text = page.extract_text()
        text = text.split('\n')
        text = map(lambda s: s.rstrip(), text)
        text = filter(lambda s: s != '', text)
        text = '\n'.join(text)
        yield text
        
def get_inbox():
    mail = imaplib.IMAP4_SSL(host)
    mail.login(username, password)
    mail.select("inbox")
    _, search_data = mail.search(None, 'UNSEEN')
    my_message = []
    for num in search_data[0].split():
        email_data = {}
        _, data = mail.fetch(num, '(RFC822)')
        # print(data[0])
        _, b = data[0]
        email_message = email.message_from_bytes(b)
        for header in ['subject', 'to', 'from', 'date']:
            # print("{}: {}".format(header, email_message[header]))
            email_data[header] = email_message[header]
        for part in email_message.walk():
            if part.get_content_type() == "text/plain":
                body = part.get_payload(decode=True)
                email_data['body'] = body.decode()
            elif part.get_content_type() == "text/html":
                html_body = part.get_payload(decode=True)
                email_data['html_body'] = html_body.decode()
        my_message.append(email_data)
    return my_message

def send_mail(text='Email Body', subject='hello world', from_email=None, to_emails=None, html=None):
    assert isinstance(to_emails, list)
    msg = MIMEMultipart('alternative')
    msg['From'] = from_email
    msg['To'] = ", ".join(to_emails)
    msg['Subject'] = subject
    txt_part = MIMEText(text, 'plain')
    msg.attach(txt_part)
    if html != None:
        html_part = MIMEText(html, 'html')
        msg.attach(html_part)
    msg_str = msg.as_string()
    # login to my smtp server
    server = smtplib.SMTP(host=host, port=587)
    server.ehlo()
    server.starttls()
    server.login(username, password)
    server.sendmail(from_email, to_emails, msg_str)
    server.quit()
    # with smtplib.SMTP() as server:
    #     server.login()
    #     pass
    
def generateEmail(gottenEmail, resources):
    emailRes = requests.post(
    "https://api.respell.ai/v1/run",
    headers={
        # This is your API key
        'Authorization': 'Bearer 9d65caeb-469e-4db9-b25b-de5281c2fb24',
        'Accept': 'application/json',
        'Content-Type': 'application/json'
    },
    data=json.dumps({
        "spellId": "44PzkK572qEohHItfQ5my",
        # This field can be omitted to run the latest published version
        # "spellVersionId": 'OAGwtWuPAb0CYGOcjMwtq',
        # Fill in dynamic values for each of your 4 input blocks
        "inputs": {
        "email": gottenEmail,
        "name": name,
        "resources": str(resources),
        "profession": profession,
        }
    }),
    )
    
    return emailRes.json()['outputs']['email_response']

def getResources(gottenEmail):
    onlyfiles = [f for f in listdir(input_path) if isfile(join(input_path, f))]
    pages = extract_doc(input_path + "\\" + onlyfiles[0])
    resources = []

    for page in pages:
        response = requests.post(
        "https://api.respell.ai/v1/run",
        headers={
            # This is your API key
            'Authorization': 'Bearer 9d65caeb-469e-4db9-b25b-de5281c2fb24',
            'Accept': 'application/json',
            'Content-Type': 'application/json'
        },
        data=json.dumps({
            "spellId": "SoTzMuMR21F-zUOvF_fBR",
            # This field can be omitted to run the latest published version
            # "spellVersionId": 'xX0v-ZCQ6KhcMesMCj5mM',
            # Fill in dynamic values for each of your 2 input blocks
            "inputs": {
            "email": gottenEmail,
            "text": page,
            
            # "instructions": "Using the following email, select lines from the text that are relevant to the email's queries. Each line or set of lines should be surrounded by square brackests, and the entire output should be surrounded by curly braces. Do not output anything if there are no relevant sentences in the text to the query. Here is the email-\n\n" + email,
            }
        }),
        )
        resources.append(response.json()['outputs']['edited_text'])
    
    return resources

if __name__ == "__main__":
    # while (True):
        emails = get_inbox()
        
        if len(emails) > 0:
            newemail = emails[0]
            emailSubj = newemail['subject']
            emailFrom = newemail['from']
            emailBody = newemail['body']

            gottenEmail = "Sender: " + emailFrom + "\nSubject: " + emailSubj + "\n\n" + emailBody
            
            resources = getResources(gottenEmail)
            genEmail = generateEmail(gottenEmail, resources)
            
            send_mail(text=genEmail, subject= "RE: " + emailSubj, from_email= username, to_emails=[emailFrom])