# Auto GUI
# This is a program that grabs information from a specified 
# Excel column and then automates the input of that list into a data entry form
# using autogui 
import pyautogui, openpyxl, time, pyperclip, datetime, calendar
from openpyxl import load_workbook
from pyautogui import write, press, keyUp, keyDown, hotkey
from time import sleep
from datetime import datetime, timedelta, date

from tkinter import Tk, simpledialog
from tkinter.filedialog import askopenfilename

def copy_clipboard():
    hotkey('ctrl', 'c')
    time.sleep(.5)
    return pyperclip.paste()

def highlight_line():
    press('numlock')
    keyDown('shiftleft')
    press('end')
    keyUp('shiftleft')
    sleep(1)
    press('numlock')
    sleep(1)

# Create the Tkinter root window
root = Tk()
root.withdraw()  # Hide the root window
completed_through= simpledialog.askinteger("Invoice Prompt", "Enter the amount billed without retention taken out:")

sleep(1)

today = date.today()
print(today)
res = calendar.monthrange(today.year, today.month)[1]
completed_date = f"Completed through {today.month}/{res}/{today.year}"


# starts once you have created a copy of the invoice, which will
# start with highlighting the customer job
# 8 tabs to the first item
# 10 tabs to the first price cell

# get the new invoice number
press('tab', presses=3)
new_inv = copy_clipboard()



press('tab', presses=7)
press('down', presses=1)
sleep(2)
prev_billed = copy_clipboard()
sleep(2)
prev_billed = prev_billed.replace(',', '')
prev_billed = prev_billed[0:-3]

press('down', presses=2)
last_period = copy_clipboard()
sleep(1)

last_period = last_period.replace(',', '')
last_period = last_period[0:-3]

new_prev_billed = int((prev_billed)) + int((last_period))
new_prev_retained = new_prev_billed * .1
press('up', presses=2)
sleep(1)
press('backspace')
write(str(new_prev_billed))
press('down')
press('backspace')
write(str(new_prev_retained))
press('tab', presses=4)
highlight_line()
# sleep(2)
press('backspace')
write(completed_date)
press('tab')
write(str(completed_through * .9))
press('down')
write(str(completed_through * -.1))