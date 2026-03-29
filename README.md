# 🏛 Temple Visit Request System

![Script](https://img.shields.io/badge/Google-AppsScript-yellow) ![Status](https://img.shields.io/badge/Status-Active-green) 

**Author:** Elder Jenkins  
**Created:** March 2026  
**Purpose:** Helps missionaries request temple visits. Automatically sends emails to the president for approval, updates the sheet, and notifies missionaries.

---

## 📑 Table of Contents
1. [Project Structure](#project-structure)  
2. [Setup Notes](#setup-notes)  
3. [How It Works](#how-it-works)  
4. [Editing Tips](#editing-tips)  
5. [Adding New Questions](#adding-new-questions)  

---

## 📂 Project Structure

- **`Code.gs`** – Main code:
```javascript
function sendEmailOnFormSubmit(e) { ... }
function doGet(e) { ... }
function getRequests() { ... }
function setStatus(timestamp, status) { ... }
```

- ```sendEmailOnFormSubmit(e)``` – Sends an email to the president when a form is submitted.
- ```doGet(e)``` – Handles Approve/Deny actions from buttons. Updates sheet and emails missionary.
- ```getRequests()``` – Optional: used for a dashboard to see all requests.
- ```setStatus(timestamp, status)``` – Optional: updates status via the dashboard.
- Google Sheet: Form Responses 1 – Where responses are saved. Columns must match exactly:
- ```Timestamp | Email Address | Names of Missionaries | Status | [Other Questions] ```

---

## ⚙️ Setup Notes
1. **Column Names:** Make sure the headers in your sheet match exactly with the script.
2. **Timestamps: Each** request is identified by its timestamp. Keep them unique.
3. **Colors & Branding:** You can change these in the script:
   ```var PRIMARY_COLOR = "#1B5E20"; // deep green```
   ```var MISSION_NAME = "Philippines Manila Mission";
   ```
## How it works
1. Missionary submits the form.
2. Script sends a nice HTML email to the president with Approve/Deny buttons.
3. President clicks a button → script updates the Status in the sheet.
4. Script sends a confirmation email to the missionary.

---

## ✏️ Editing Tips
- Always **back up your Google Sheet** before making changes.
- Add new questions in the form → they will automatically show in emails.
- Keep the HTML styling consistent if changing colors or fonts.

## ➕ Adding New Questions
1. Add your new question to the Google Form.
2. Make sure the response shows up in the sheet.
3. ```sendEmailOnFormSubmit(e)``` automatically loops through all questions and adds them to the email.

## Contact / Support 
If something breaks or you need help:

- Elder Jenkins
- Email: ```jenkins.jaden@missionary.org```
- Leave a note in the README if you fix something so the next editor knows
