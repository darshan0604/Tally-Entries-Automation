# 🧠 Tally Automation Using Python

This project automates repetitive data entry tasks in **Tally ERP** using **Python**, significantly reducing manual effort and errors.

It reads accounting data from an Excel sheet and simulates user input in Tally using GUI automation, with built-in error handling and dynamic ledger creation using **OpenCV image recognition**.

---

## 🚀 Features

- ✅ Automates data entry for **Contra**, **Payment**, and **Receipt** vouchers in Tally
- 📊 Reads data directly from a formatted Excel file
- 🧠 Intelligent error detection using OpenCV template matching
- 🔁 Automatically creates missing ledgers if Tally throws an error
- ⏯ Pause/Resume/Exit the script using custom **keyboard shortcuts**
- 💾 Tracks failed entries for later re-processing
- ⚙️ Minimal manual intervention once started

---

## 🛠 Tech Stack

- **Python**
- **pandas** – For reading and processing Excel data
- **pyautogui** – For simulating keyboard input
- **OpenCV** – For detecting errors by image comparison
- **keyboard** – For user control (Pause/Resume/Stop)
- **Tally ERP** – As the target accounting system

---

## 📂 File Structure
📁 Tally-Automation/
├── main.py # Core automation script
├── Screenshotnew.png # Reference image to detect Tally error popup
├── data.xlsx # Sample Excel input file (not included in repo)
└── README.md # Project documentation

---

## 📌 Keyboard Shortcuts

| Shortcut         | Function            |
|------------------|---------------------|
| Ctrl + `<`       | Pause / Resume      |
| Ctrl + `>`       | Start / Restart     |
| Ctrl + `S`       | Stop and Exit       |

---

## 🔍 How It Works

1. **Read Excel:**  
   Reads entries like `Date`, `Account`, `Particulars`, `Deposit`, `Withdrawal`, `Narration`, etc.

2. **Voucher Type Decision:**  
   Determines if entry is `Contra`, `Payment`, or `Receipt`.

3. **Automated Input:**  
   Simulates Tally navigation and data entry using `pyautogui`.

4. **Error Detection:**  
   Uses `OpenCV` to check if a Tally error popup appears.

5. **Auto Ledger Creation:**  
   If error detected, script exits current screens, creates a new ledger, and retries entry.

6. **Hotkey Control:**  
   Allows manual pausing/resuming or exiting the process during runtime.

---

## 📸 Sample Use Case

- Automating bank statement entries from Excel to Tally
- Reducing entry time from **several hours to a few minutes**
- Ensuring 100% consistency and structure across Tally records

---

## ❗ Important Notes

- ⚠️ Ensure Tally is **open, configured**, and visible on the screen before starting.
- 🖥 Use on the same screen resolution where `Screenshotnew.png` was captured for best accuracy.
- 📄 Excel input must follow a specific format (Date, Account, Head, Particulars, Deposit, Withdrawal, Under, Bank).

---

## 🤝 Contribution

Feel free to fork this project, submit pull requests, or suggest enhancements.  
If you use this automation in your workflow, I’d love to hear about it!

---

## 📧 Contact

**Author:** Darshan Soni  
📬 [LinkedIn](https://www.linkedin.com/) | 📧 [your-email@example.com]

---

## 🏷️ Tags

`#Python` `#Automation` `#TallyERP` `#pyautogui` `#OpenCV` `#ExcelAutomation` `#AccountingTech` `#KeyboardShortcuts` `#RPA`

