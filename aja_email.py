from time import sleep
from win32gui import GetWindowText, GetForegroundWindow, GetWindowPlacement
import smtplib
from email.MIMEMultipart import MIMEMultipart
from email.MIMEText import MIMEText
from sys import argv, path

print('This program sends an email when deposition completes.')

# Load trigger window information from file
try:
    with open(path[0] + '\\trigger_window.dat', 'r') as f:
        trig_title = f.readline().strip('\n')
        trig_loc = f.readline()
except:
    raise Exception('trigger_window.dat not found.  Use learn_window.py to create it')

with open(path[0] + '\\google.dat', 'r') as f:
    email = f.readline().strip('\n')
    password = f.readline()


if len(argv) > 1:
    toaddr = argv[1]
else:
    toaddr = raw_input('Enter email of recipient:')

# Message to be sent
subject = 'Deposition finished.'
body = ''

msg = MIMEMultipart()
msg['From'] = email
msg['To'] = toaddr
msg['Subject'] = subject
msg.attach(MIMEText(body, 'plain'))

server = smtplib.SMTP('smtp.gmail.com', 587)
server.starttls()
server.login(email, 'fu11ert0n')
text = msg.as_string()

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

    print('Trigger window detected. Sending email to {}'.format(toaddr))
    server.sendmail(email, toaddr, text)

    while trigger():
        # Loop until different window in foreground
        sleep(5)

    print('Trigger window left foreground')
