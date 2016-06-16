from win32gui import GetWindowText, GetForegroundWindow, GetWindowPlacement
from time import sleep
from sys import path

# Learn what window we're looking for
# Goes by window title and location of window on the screen
print('Bring trigger window to the foreground in 10 seconds')
sleep(10)
active_window = GetForegroundWindow()
active_title = GetWindowText(active_window)
active_location = GetWindowPlacement(active_window)
print('Saving to window.dat:\nTitle:\n {}\nPlacement:\n{}'.format(active_title, active_location))
with open(path[0] + '\\window.dat', 'w') as f:
    f.write('{}\n{}'.format(active_title, active_location))
