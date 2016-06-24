from time import sleep
from win32gui import GetWindowText, GetForegroundWindow, GetWindowPlacement
from googlevoice import Voice
from googlevoice.util import input
from sys import argv, path
import imaplib
import email
import re

# TODO:  Let text message deactivate/activate program
#        Remember handle of trigger window so it can't trigger twice
#        Identify failed deposition and notify
#        Run all the time and activate by sms
#        Allow multiple recipients?

# Load trigger window information from file
try:
    with open(path[0] + '\\window.dat', 'r') as f:
        trig_title = f.readline().strip('\n')
        trig_loc = f.readline()
except:
    raise Exception('window.dat not found.  Run train.py')

print('This program sends an SMS when deposition completes.')

if len(argv) > 1:
    phoneNumber = argv[1]
else:
    phoneNumber = input('Enter phone number of SMS recipient:')

# Try to identify input phone number with regular expression
phonePattern = re.compile('(\d{3})\D*(\d{3})\D*(\d{4})$')
# This will be regular expression for getting phone number from received sms
voicephonePattern = re.compile('\.1(\d{3})\D*(\d{3})\D*(\d{4})\.')

phonestr = ''.join(phonePattern.search(phoneNumber).groups())

with open(path[0] + '\\google.dat', 'r') as f:
    emailaddr = f.readline().strip('\n')
    password = f.readline()

print('Logging in to google voice account...')
voice = Voice()
voice.login(emailaddr, password)
M = imaplib.IMAP4_SSL('imap.gmail.com')
M.login(emailaddr, password)
rv, data = M.select('INBOX')
nmsgs0 = int(data[0])

# Message to be sent
text = 'Deposition finished. Reply \'q\' to stop notifications.'

def trigger():
    active_window = GetForegroundWindow()
    active_title = GetWindowText(active_window)
    active_loc = str(GetWindowPlacement(active_window))
    if (active_title == trig_title) and (active_loc == trig_loc):
        return True
    return False

def stop():
    global nmsgs0
    # Check sms for stop message
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
                    voice.send_sms(phoneNumber, 'Quitting.')
                    return True
        nmsgs0 = nmsgs

def main_loop():
    # This is a function so that I can use return to break from nested loop
    while 1:
        print('Waiting for trigger window.')
        while not trigger():
            # Loop until window is in foreground
            if stop(): return
            sleep(5)

        print('Trigger window detected. Sending SMS to {}'.format(phoneNumber))
        voice.send_sms(phoneNumber, text)

        while trigger():
            # Loop until different window in foreground
            if stop(): return
            sleep(5)

        print('Trigger window left foreground')

main_loop()
