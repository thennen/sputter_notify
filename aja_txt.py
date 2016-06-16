from time import sleep
from win32gui import GetWindowText, GetForegroundWindow, GetWindowPlacement
from googlevoice import Voice
from googlevoice.util import input
from sys import argv, path

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

with open(path[0] + '\\google.dat', 'r') as f:
    email = f.readline().strip('\n')
    password = f.readline()

print('Logging in to google voice account...')
voice = Voice()
voice.login(email, password)

# Message to be sent
text = 'Deposition finished.'

def trigger():
    active_window = GetForegroundWindow()
    active_title = GetWindowText(active_window)
    active_loc = str(GetWindowPlacement(active_window))
    if (active_title == trig_title) and (active_loc == trig_loc):
        return True
    return False

while 1:
    print('Waiting for trigger window.')
    while not trigger():
        # Loop until window is in foreground
        sleep(5)

    print('Trigger window detected. Sending SMS to {}'.format(phoneNumber))
    voice.send_sms(phoneNumber, text)

    while trigger():
        # Loop until different window in foreground
        sleep(5)

    print('Trigger window left foreground')
