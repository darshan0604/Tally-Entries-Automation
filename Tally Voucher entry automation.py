import pandas as pd
import pyautogui as pg
import cv2
import numpy as np
import time
import keyboard  # Import the keyboard module

try:
    error_image = cv2.imread('Screenshotnew.png')
    if error_image is None:
        raise FileNotFoundError("Error image not found or cannot be read.")
except Exception as e:
    print("Error loading error image:", e)

# Function to check if error image is present on the screen
def check_for_error():
    try:
        # Take a screenshot of the entire screen
        screenshot = pg.screenshot()
        
        # Convert the screenshot to OpenCV format (BGR)
        screenshot_cv = np.array(screenshot)
        screenshot_cv = cv2.cvtColor(screenshot_cv, cv2.COLOR_RGB2BGR)
        
        # Find the error image within the screenshot
        result = cv2.matchTemplate(screenshot_cv, error_image, cv2.TM_CCOEFF_NORMED)
        threshold = 0.8
        locations = np.where(result >= threshold)
        
        # If matches are found, return True (error detected)
        if locations[0].size > 0:
            return True
        
        # If no matches found, return False (no error detected)
        return False
    except Exception as e:
        print("Error checking for error:", e)

# Load data from Excel file
excel_path = r"C:\Users\Dell\.ipynb_checkpoints\Tally\Raag - HDFC Apr to Jan 24 (1) - Copy.xls"
df = pd.read_excel(excel_path, sheet_name="Summary")
date = df["Date"].values
account = df["Account"].values
particular = df["Head"].values
naration = df["Particulars"].values
Withdrawal = df["Withdrawal"].values
Deposit = df["Deposit"].values
Under = df["Under"].values
Bank = df["Bank"].values

zipped = zip(date, account, particular, naration, Withdrawal, Deposit, Under)

time.sleep(3)

# Initialize a list to store incorrect entries
lost_entries = []
# Write The contra bank for contra entry
contra_bank = df["Bank"].astype(str)

# Initialize control variables
paused = False
stopped = False

# Function to handle key presses for control
def on_key_event(e):
    global paused, stopped
    if keyboard.is_pressed('ctrl') and keyboard.is_pressed('<'):
        paused = not paused
        if paused:
            print("Paused")
        else:
            print("Resumed")
    elif keyboard.is_pressed('ctrl') and keyboard.is_pressed('>'):
        paused = False
        stopped = False
        print("Started")
    elif keyboard.is_pressed('ctrl') and keyboard.is_pressed('s'):
        stopped = True
        print("Exit")

# Set up keyboard event listener
keyboard.on_press(on_key_event)


# Start of the Loop:
for (da, ac, pa, na, wi, de, un) in zipped:
    while paused:
        time.sleep(1)  # Wait while the script is paused

    if stopped:
        break  # Stop the script if stopped

    head = pa  # Assign the value from "Head" column to the head variable
    
    if any(contra_bank == head):  
        try:
            pg.press("F4")
            pg.press("F2")
            
            # Convert to Pandas Timestamp
            date_obj = pd.Timestamp(da)
            # Format date as "1-4-2023"
            formatted_date = date_obj.strftime("%#d-%#m-%Y")
            # Typing date
            pg.typewrite(formatted_date)
            pg.press("enter")
            if pd.isna(wi):
                pg.typewrite(ac)
                pg.press("enter")
                pg.typewrite(pa)
                pg.press("enter")
                time.sleep(3)
                pg.typewrite(str(de))  # Convert withdrawal to string before typing

            elif pd.isna(de):
                pg.typewrite(pa)
                pg.press("enter")
                pg.typewrite(ac)
                pg.press("enter")
                time.sleep(3)
                pg.typewrite(str(wi))

            pg.press("enter")
            pg.press("enter")  # Pressing enter based on your Tally input steps
            pg.press("enter")
            pg.press("enter")
            pg.press("enter")
            pg.press("enter")
            pg.press("enter")
            pg.press("enter")
            pg.press("enter")
            pg.press("enter")
            pg.press("enter")
            pg.press("enter")
            pg.typewrite(na)  # Typing narration
            pg.hotkey("ctrl", "a")
        
        except Exception as e:
            print("Error:", e)
    
    elif pd.Series(pd.isna(de)).any() & pd.Series(contra_bank != head).any(): # Check if any values in wi are NaN and any values in contra_bank don't match head

        try:
            pg.press("F5")
            pg.press("F2")
            # Convert to Pandas Timestamp
            date_obj = pd.Timestamp(da)
            # Format date as "1-4-2023"
            formatted_date = date_obj.strftime("%#d-%#m-%Y")
            # Typing date
            pg.typewrite(formatted_date)
            pg.press("enter")
            pg.typewrite(ac)
            pg.press("enter")
            pg.typewrite(pa)
            
            time.sleep(2)

            if check_for_error():
                # If error detected, create ledger and reattempt voucher entry
                time.sleep(1)
                pg.press("esc")
                pg.press("esc")
                pg.press("esc")
                pg.press("esc")
                pg.press("esc")
                
                # Create ledger
                pg.press("c")
                pg.typewrite("Ledger")
                pg.press("enter")
                pg.typewrite(pa)
                pg.press("enter")
                pg.press("enter")
                pg.typewrite(un)
                pg.hotkey("ctrl", "a")
                time.sleep(2)

                pg.press("esc")
                pg.press("esc")
                pg.press("esc")
                pg.press("esc")
                time.sleep(3)
                pg.press("V")
                
                # Store current entry information
                lost_entries.append((da, ac, pa, na, wi, de, un))
            else:
                pg.press("enter")
                time.sleep(1)
                pg.typewrite(str(wi))  # Convert deposit to string before typing
                pg.press("enter")
                pg.press("enter")  # Pressing enter based on your Tally input steps
                pg.press("enter")
                pg.press("enter")
                pg.press("enter")
                pg.press("enter")
                pg.press("enter")
                pg.typewrite(na)  # Typing narration
                pg.hotkey("ctrl", "a")
        
        except Exception as e:
            print("Error:", e)

    elif pd.Series(pd.isna(wi)).any() & pd.Series(contra_bank != head).any():
    # Check if any values in wi are NaN and any values in contra_bank don't match head

        try:
            pg.press("F6")
            pg.press("F2")
            # Convert to Pandas Timestamp
            date_obj = pd.Timestamp(da)
            # Format date as "1-4-2023"
            formatted_date = date_obj.strftime("%#d-%#m-%Y")
            # Typing date
            pg.typewrite(formatted_date)
            pg.press("enter")
            pg.typewrite(ac)
            pg.press("enter")
            pg.typewrite(pa)

            time.sleep(2)

            if check_for_error():
                # If error detected, create ledger and reattempt voucher entry
                time.sleep(1)
                pg.press("esc")
                pg.press("esc")
                pg.press("esc")
                pg.press("esc")
                pg.press("esc")
                
                # Create ledger
                pg.press("c")
                pg.typewrite("Ledger")
                pg.press("enter")
                pg.typewrite(pa)
                pg.press("enter")
                pg.press("enter")
                pg.typewrite(un)
                pg.hotkey("ctrl", "a")
                time.sleep(2)

                pg.press("esc")
                pg.press("esc")
                pg.press("esc")
                pg.press("esc")
                time.sleep(3)
                pg.press("V")
                
                # Store current entry information
                lost_entries.append((da, ac, pa, na, wi, de, un))
            else:
                pg.press("enter")
                time.sleep(1)
                pg.typewrite(str(de))  # Convert withdrawal to string before typing
                pg.press("enter")
                pg.press("enter")  # Pressing enter based on your Tally input steps
                pg.press("enter")
                pg.press("enter")
                pg.press("enter")
                pg.press("enter")
                pg.press("enter")
                pg.press("enter")
                pg.typewrite(na)  # Typing narration
                pg.hotkey("ctrl", "a")
        
        except Exception as e:
            print("Error:", e)

# Process the entry that was left when ledger creation was triggered
for entry in lost_entries:
    while paused:
        time.sleep(1)  # Wait while the script is paused

    if stopped:
        break  # Stop the script if stopped

    (da, ac, pa, na, wi, de, un) = entry

    if pd.isna(de):  # Check if withdrawal is empty and head is not "SBI Bank"
        try:
            pg.press("F5")
            pg.press("F2")
            # Convert to Pandas Timestamp
            date_obj = pd.Timestamp(da)
            # Format date as "1-4-2023"
            formatted_date = date_obj.strftime("%#d-%#m-%Y")
            # Typing date
            pg.typewrite(formatted_date)
            pg.press("enter")
            pg.typewrite(ac)
            pg.press("enter")
            pg.typewrite(pa)
            
            time.sleep(2)
           
            pg.press("enter")
            time.sleep(1)
            pg.typewrite(str(wi))  # Convert deposit to string before typing
            pg.press("enter")
            pg.press("enter")  # Pressing enter based on your Tally input steps
            pg.press("enter")
            pg.press("enter")
            pg.press("enter")
            pg.press("enter")
            pg.press("enter")
            pg.typewrite(na)  # Typing narration
            pg.hotkey("ctrl", "a")
        
        except Exception as e:
            print("Error:", e)

    elif pd.isna(wi):  # Check if deposit is empty and head is not "SBI Bank"
        try:
            pg.press("F6")
            pg.press("F2")
            # Convert to Pandas Timestamp
            date_obj = pd.Timestamp(da)
            # Format date as "1-4-2023"
            formatted_date = date_obj.strftime("%#d-%#m-%Y")
            # Typing date
            pg.typewrite(formatted_date)
            pg.press("enter")
            pg.typewrite(ac)
            pg.press("enter")
            pg.typewrite(pa)

            time.sleep(2)

            pg.press("enter")
            time.sleep(1)
            pg.typewrite(str(de))  # Convert withdrawal to string before typing
            pg.press("enter")
            pg.press("enter")  # Pressing enter based on your Tally input steps
            pg.press("enter")
            pg.press("enter")
            pg.press("enter")
            pg.press("enter")
            pg.press("enter")
            pg.press("enter")
            pg.typewrite(na)  # Typing narration
            pg.hotkey("ctrl", "a")
        
        except Exception as e:
            print("Error:", e)

print("Script completed or stopped.")
