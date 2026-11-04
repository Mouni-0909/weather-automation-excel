import pyautogui
import time
import pyperclip
import openpyxl
import os
from datetime import datetime

pyautogui.FAILSAFE = True
pyautogui.PAUSE = 1

print("Step 1 : Open the chrome browser")    
time.sleep(1)
#now we are using hotkey to open the run dialog box
pyautogui.hotkey('win', 'r')  # Opens the Run dialog
time.sleep(2)
#now we are using write to type the name of the application we want to open
pyautogui.write('chrome')  # Types 'chrome' in the Run dialog
time.sleep(1)
#now we are using press to simulate pressing the Enter key
pyautogui.press('enter')  # Presses Enter to launch the application
time.sleep(3)
#verifying if the application is opened or not by taking a screenshot
pyautogui.screenshot('chrome_opened.png')  # Takes a screenshot after opening Chrome

print("Step 2 : Open the weather.com website")
time.sleep(2)
#trying to open the website using hotkey in a newtab
pyautogui.hotkey('ctrl', 't')  # Opens a new tab in Chrome
time.sleep(2)
#now we are using write to type the name of the website we want to open
pyautogui.typewrite('https://www.accuweather.com/en/in/chennai/206671/weather-forecast/206671'
, interval=0.3)  # Types 'weather.com' in the address bar
time.sleep(1)
#now we are using press to simulate pressing the Enter key
pyautogui.press('enter')  # Presses Enter to navigate to the website
time.sleep(1)
#taking a screenshot to verify if the website is opened or not
pyautogui.screenshot('weather_opened.png')  # Takes a screenshot after opening the website

print("Step 3 : copy the full data of the website")
#using hotkey to select all the content of the webpage and copy it
pyautogui.hotkey('ctrl', 'a')  # Selects all content on the webpage
time.sleep(1)
#using hotkey to copy the selected content
pyautogui.hotkey('ctrl', 'c')  # Copies the selected content

print("Step 4 : Open the excel application and paste the copied data")
time.sleep(2)
#using hotkey to open the run dialog box
pyautogui.hotkey('win', 'r')  # Opens the Run dialog
time.sleep(2)
#using typewrite to type the name of the application we want to open
pyautogui.write('excel')  # Types 'excel' in the Run dialog
time.sleep(1)
#using press to simulate pressing the Enter key
pyautogui.press('enter')  # Presses Enter to launch Excel
time.sleep(2)
pyautogui.hotkey('ctrl', 'n')  # Opens a new workbook in Excel
time.sleep(2)
pyautogui.press('enter')  # Presses Enter to confirm creating a new workbook
time.sleep(1)
pyautogui.press('enter')  # Presses Enter to select the cell where the data will be pasted
#using hotkey to paste the copied content into excel
pyautogui.hotkey('ctrl', 'v')  # Pastes the copied content
# Step 5: Add today's date, time, weather data and comment in a new row

time.sleep(2)

# Get weather data from clipboard
weather_data = pyperclip.paste()

# Get today's date and time
today = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

# Add a comment
comment = "Nalla icecream sapdalam macha!!"

# Create Excel workbook
workbook = openpyxl.Workbook()
sheet = workbook.active

# Add headings
sheet.append(["Date & Time", "Weather Data", "Comment"])

# Add data in a new row
sheet.append([today, weather_data[:500], comment])

# Save Excel file with today's date
filename = f"daily_report_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
workbook.save(filename)

print("Excel file saved successfully")

# Open Excel file
os.startfile(filename)

time.sleep(5)

# Take screenshot of Excel sheet
pyautogui.screenshot("daily_report_screenshot.png")

print("Screenshot saved successfully")