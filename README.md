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

