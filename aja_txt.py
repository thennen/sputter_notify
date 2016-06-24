'''
This program waits for a certain window to pop up, and then sends an SMS to a
number of your choice.  Can be closed by replying 'q' to the sms.

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
from googlevoice.util import input
from sys import argv, path
import imaplib
import email
import re

# TODO:
#        Identify failed deposition and notify
#        Run all the time and activate by sms or email
#        Merge email and sms program into one

print('This program sends an SMS when deposition completes.')

# Load trigger window information from file
try:
    with open(path[0] + '\\trigger_window.dat', 'r') as f:
        trig_title = f.readline().strip('\n')
        trig_loc = f.readline()
except:
    raise Exception('trigger_window.dat not found.  Use learn_window.py to create it')

if len(argv) > 1:
    phoneNumber = argv[1]
else:
    phoneNumber = input('Enter phone number of SMS recipient:')
# Try to identify input phone number with regular expression
phonePattern = re.compile('(\d{3})\D*(\d{3})\D*(\d{4})$')
phonestr = ''.join(phonePattern.search(phoneNumber).groups())
# This will be regular expression for getting phone number from received sms
voicephonePattern = re.compile('\.1(\d{3})\D*(\d{3})\D*(\d{4})\.')

# Read user name and password from plaintext file google.dat
with open(path[0] + '\\google.dat', 'r') as f:
    emailaddr = f.readline().strip('\n')
    password = f.readline()

print('Logging in to google account...')
# Connect to google voice for sending SMS
voice = Voice()
voice.login(emailaddr, password)
# Connect to IMAP for receiving emails and SMS
M = imaplib.IMAP4_SSL('imap.gmail.com')
M.login(emailaddr, password)
rv, data = M.select('INBOX')
nmsgs0 = int(data[0])

# Message to be sent
text = 'Deposition finished. Reply \'q\' to stop notifications.'

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
    # Check sms for stop condition -- recipient replies 'q'
    global nmsgs0
    rv, data = M.select('INBOX')
    nmsgs = int(data[0])
    if nmsgs > nmsgs0:
        # Maybe you got more than one email
        for num in range(nmsgs0 + 1, nmsgs + 1):
            rv, data = M.fetch(num, '(RFC822)')
            msg = email.message_from_string(data[0][1])
            sender = msg['From']
            subject = msg['Subject']
            body = msg.get_payload()
            issms = subject.startswith('SMS from ')
            if issms:
                # Make sure it came from the same number that we are spamming
                fromphone = ''.join(voicephonePattern.search(sender).groups())
                if fromphone == phonestr and body.strip() == 'q':
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

        print('Trigger window detected. Sending SMS to {}'.format(phoneNumber))
        voice.send_sms(phoneNumber, text)

main_loop()
voice.send_sms(phoneNumber, 'Program Terminated.')
M.close()
M.logout()
voice.logout()
