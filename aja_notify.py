'''
This program waits for a certain window to pop up, and then sends an SMS/email
to a number/address of your choice.  Can be closed by replying 'q' to the
notification.

Trigger window is identified only by its title and its location and shape on the
screen.  Specify it by creating a trigger_window.dat file using learn_window.py

A dedicated google account with a google voice number needs to be set up.  Give
the login credentials in a plaintext file 'google.dat'.  First line should have
the email address and second line the password

Tyler Hennen 2016
'''

from time import sleep
from win32gui import GetWindowText, GetForegroundWindow, GetWindowPlacement
from googlevoice import Voice
from sys import argv, path
import imaplib
import smtplib
from email import message_from_string
from email.MIMEMultipart import MIMEMultipart
from email.MIMEText import MIMEText
import re

# TODO:
#        Identify failed deposition and notify
#        Run all the time and activate by sms or email

print('This program sends an SMS when deposition completes.')

# Load trigger window information from file
try:
    with open(path[0] + '\\trigger_window.dat', 'r') as f:
        trig_title = f.readline().strip('\n')
        trig_loc = f.readline()
except:
    raise Exception('trigger_window.dat not found.  Use learn_window.py to create it')

# Get recipient email address or phone number
# Try to identify input as email or phone number with regular expressions
sms = None
recipient = None
email_re = re.compile('^([a-zA-Z0-9_\-\.]+)@([a-zA-Z0-9_\-\.]+)\.([a-zA-Z]{2,5})$')
phone_re = re.compile('(\d{3})\D*(\d{3})\D*(\d{4})$')
# Take email/phone number from command line argument if there is one
if len(argv) > 1:
    recipient = argv[1]
else:
    recipient = raw_input('Enter recipient phone number or email address:')
while 1:
    if email_re.match(recipient):
        sms = False
        break
    elif phone_re.match(recipient):
        sms = True
        phonestr = ''.join(phone_re.search(recipient).groups())
        break
    # Ask again
    print('I don\'t think that\'s a valid email address or phone number.')
    recipient = raw_input('Enter recipient phone number or email address:')

# Read user name and password from plaintext file google.dat
with open(path[0] + '\\google.dat', 'r') as f:
    emailaddr = f.readline().strip('\n')
    password = f.readline()

# Initialize receiving email/sms
imap = imaplib.IMAP4_SSL('imap.gmail.com')
imap.login(emailaddr, password)
rv, data = imap.select('INBOX')
nmsgs0 = int(data[0])

# Initialize sending sms or email
if sms:
    # Connect to google voice for sending SMS
    print('Logging in to google account...')
    voice = Voice()
    voice.login(emailaddr, password)
    # sms to be sent
    text = 'Deposition finished. Reply \'q\' to stop notifications.'
else:
    # email to be sent
    subject = 'Deposition finished.'
    body = 'Reply \'q\' to terminate program.'
    smtp_msg = MIMEMultipart()
    smtp_msg['From'] = emailaddr
    smtp_msg['To'] = recipient
    smtp_msg['Subject'] = subject
    smtp_msg.attach(MIMEText(body, 'plain'))
    smtp = smtplib.SMTP('smtp.gmail.com', 587)
    smtp.starttls()
    smtp.login(emailaddr, password)
    text = smtp_msg.as_string()

prev_trig_handle = long(0)
def trigger():
    global prev_trig_handle
    # Return True on trigger event -- new trigger window in foreground
    active_window = GetForegroundWindow()
    active_title = GetWindowText(active_window)
    active_loc = str(GetWindowPlacement(active_window))
    title_match = (active_title == trig_title)
    loc_match = (active_loc == trig_loc)
    handle_match = (prev_trig_handle == active_window)
    if title_match and loc_match and not handle_match:
        prev_trig_handle = active_window
        return True
    return False

def stop():
    # Check sms/email for stop condition -- recipient replies 'q'
    global nmsgs0
    rv, data = imap.select('INBOX')
    # Number of messages in inbox
    nmsgs = int(data[0])
    if nmsgs > nmsgs0:
        # Maybe you got more than one email
        for num in range(nmsgs0 + 1, nmsgs + 1):
            rv, data = imap.fetch(num, '(RFC822)')
            msg = message_from_string(data[0][1])
            sender = msg['From']
            subject = msg['Subject']
            body = msg.get_payload()
            issms = subject.startswith('SMS from ')
            if sms and issms:
                # Make sure it came from the same number that we are spamming
                fromphone = ''.join(phone_re.search(subject).groups())
                if fromphone == phonestr and body.strip() == 'q':
                    return True
            elif not sms and not issms and recipient in sender:
                email_lines = body.split('\n')
                if len(email_lines) > 0 and email_lines[0].strip() == 'q':
                    return True
        nmsgs0 = nmsgs

def main_loop():
    # This is a function so that I can use return to break from nested loop
    while 1:
        print('Waiting for trigger window.')
        while not trigger():
            # Loop until new trigger window is in foreground
            # or received sms to terminate
            if stop(): return
            sleep(5)

        if sms:
            print('Trigger window detected. Sending SMS to {}'.format(recipient))
            voice.send_sms(recipient, text)
        else:
            print('Trigger window detected. Sending email to {}'.format(recipient))
            smtp.sendmail(emailaddr, recipient, text)


main_loop()
if sms:
    voice.send_sms(recipient, 'Program Terminated.')
    voice.logout()
else:
    term_text = text.replace(subject, 'Program Terminated.')
    term_text = term_text.replace(body, '')
    smtp.sendmail(emailaddr, recipient, term_text)
    smtp.close()
imap.close()
imap.logout()
