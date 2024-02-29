import pandas as pd
import pyautogui as pg
import time

excel_path = r"C:\Users\Dell\.ipynb_checkpoints\Voucher Template.xls"
df = pd.read_excel(excel_path, sheet_name="TEST")
date = df["Date"].values
account = df["Account"].values
particular = df["Particular"].values
amount = df["Amount"].values
naration = df["Naration"].values
zipped = zip(date, account, particular, amount, naration)  # Fix the variable assignment
time.sleep(3)


# Start of the Loop:
for (a, b, c, d, e) in zipped:  # Indent the loop block
    pg.press("F2")
   
    
    # Convert to Pandas Timestamp
    date_obj = pd.Timestamp(a)
    
    # Format date as "1-4-2023"
    formatted_date = date_obj.strftime("%#d-%#m-%Y")
   
    # Typing date
    pg.typewrite(formatted_date)

   
    
    pg.press("enter")

    pg.typewrite(b)
   
    pg.press("enter")
   
    
    pg.typewrite(c)
    
    pg.press("enter")
    
    
    pg.typewrite(str(d))  # Convert c to string before typing
  
    pg.press("enter")
    pg.press("enter")
    pg.press("enter")
    pg.press("enter")
    pg.press("enter")
    pg.press("enter")
    pg.press("enter")

    pg.typewrite(e)  # Convert c to string before typing
    
      
    pg.hotkey("ctrl", "a")
      # Ensure proper indentation for readability and syntax
    

    # Tally-Entries-Automation
