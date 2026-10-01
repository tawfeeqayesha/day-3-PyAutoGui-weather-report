
import pyautogui
import pyperclip
import time
import os
import subprocess
from datetime import datetime
from openpyxl import Workbook

# Safety settings
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.3

# Current date and time
now = datetime.now()
date_str = now.strftime("%Y-%m-%d_%H-%M-%S")
date_time = now.strftime("%Y-%m-%d %H:%M:%S")

# File paths
folder = os.path.expanduser("~/Desktop")
excel_path = os.path.join(folder, "Dubai_Weather_" + date_str + ".xlsx")
screenshot_path = os.path.join(folder, "Dubai_Weather_" + date_str + ".png")

# Weather website
weather_url = "https://www.timeanddate.com/weather/united-arab-emirates/dubai"

print("Daily Weather Report Started")

# Open the weather website in Chrome
subprocess.Popen(["open", "-a", "Google Chrome", weather_url])
time.sleep(8)

# Create Excel workbook
workbook = Workbook()
sheet = workbook.active
sheet.title = "Weather Report"

# Add headings
headings = ["Date & Time", "Weather Information", "Comment"]

for column, heading in enumerate(headings, start=1):
    sheet.cell(row=1, column=column, value=heading)

# Add report information
sheet["A2"] = date_time
sheet["B2"] = "Dubai weather website: " + weather_url
sheet["C2"] = "Check the forecast before outdoor activities."

# Adjust column widths
sheet.column_dimensions["A"].width = 25
sheet.column_dimensions["B"].width = 65
sheet.column_dimensions["C"].width = 45

# Save Excel file
workbook.save(excel_path)
print("Excel file saved:", excel_path)

# Open the report in Excel
subprocess.Popen(["open", "-a", "Microsoft Excel", excel_path])
time.sleep(8)

# Take screenshot
pyautogui.screenshot().save(screenshot_path)

print("Screenshot saved:", screenshot_path)
print("Report completed!")