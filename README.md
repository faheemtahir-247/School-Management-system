# 🎓 School Management System (Google Apps Script + Google Sheets)

A lightweight, serverless, and highly efficient School Management System built entirely on **Google Workspace**. It uses **Google Apps Script** as the backend server and **Google Sheets** as a secure, real-time database. 

This system provides a clean HTML/JS portal for data entry, fee tracking, and smart invoice generation without the need for traditional hosting or SQL databases.

## 🚀 Live Demo
**[Click Here to view the Live System Portal](https://script.google.com/macros/s/AKfycbxIDitKukYf1pmHzj5QQ3d_DTsv86nP3QeTWS5XZPSYP_KpYJSyT_F-cKf3LbD5pHSWTA/exec)**

---

## ✨ Key Features

* **Serverless Architecture:** Completely hosted on Google Cloud via Apps Script.
* **Smart Database (Google Sheets):** Auto-sorts data chronologically and categorically by Class (Class 1 to 8).
* **Dynamic Fee Management:** * Automatically calculates current month dues and previous arrears (pending fees).
  * Hides future months automatically to keep the dashboard clean.
  * Real-time recalculation of Total Paid and Remaining Balances.
* **Smart Invoicing System:** One-click generation of printable fee slips/invoices detailing current fees and previous arrears.
* **Modern UI:** Responsive frontend built with **Bootstrap 5**, ensuring it works perfectly on desktops and tablets.
* **AJAX Integrated:** Seamless, page-reload-free experience using `google.script.run`.

---

## 🛠️ Tech Stack

* **Frontend:** HTML5, CSS3, JavaScript, Bootstrap 5
* **Backend:** Google Apps Script (ES6 JavaScript)
* **Database:** Google Sheets
* **Deployment:** Google Web Apps

---

## ⚙️ Installation & Setup Guide

Want to deploy this on your own Google Account? Follow these simple steps:

### Step 1: Prepare the Database (Google Sheet)
1. Go to [Google Sheets](https://sheets.new/) and create a new blank spreadsheet.
2. Rename the first worksheet to `Sheet1`.
3. Add the following exact headers in Row 1 (A to U):
   `Student ID` | `Name` | `Father Name` | `Contact` | `Address` | `Class` | `Monthly Fee` | `Jan` | `Feb` | `Mar` | `Apr` | `May` | `Jun` | `Jul` | `Aug` | `Sep` | `Oct` | `Nov` | `Dec` | `Total Paid` | `Remaining Fee`

### Step 2: Configure the Backend
1. In your Google Sheet, click on **Extensions > Apps Script**.
2. Delete any existing code in the editor.
3. Create a file named `Code.gs` and paste the contents of the `Code.js` file from this repository.
4. Run the `setupSheetHeaders()` function once from the editor to ensure your headers are correctly configured.

### Step 3: Configure the Frontend
1. In the Apps Script editor, click the **+** icon next to Files and select **HTML**.
2. Name the file exactly **`Index`** (with a capital 'I'). *Do not type the .html extension, Google adds it automatically.*
3. Paste the contents of `Index.html` from this repository into this file.
4. Save the project (Ctrl + S).

### Step 4: Deploy as Web App
1. Click on the **Deploy** button at the top right of the Apps Script editor and select **New deployment**.
2. Click the gear icon next to "Select type" and choose **Web app**.
3. Under "Execute as", select **Me**.
4. Under "Who has access", select **Anyone**.
5. Click **Deploy**, grant the necessary permissions, and copy your Live Web App URL!

---

## 📂 Repository Structure

```text
├── Code.js        # Backend logic (Upload as Code.gs in Apps Script)
├── Index.html     # Frontend UI & JavaScript logic
└── README.md      # Project documentation