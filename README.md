# attrition-risk-flagging-system
# 🚨 AI-Powered Attrition Risk Flagging System

An automation tool built with **Google Sheets + Google Apps Script** that automatically detects employees at risk of attrition and sends real-time email alerts — with zero manual monitoring.

## 📌 Overview

HR teams often rely on manually scanning spreadsheets to spot employees who might be at risk of leaving. This project automates that process end-to-end:

- Reads employee data (tenure, performance score, department) directly from a Google Sheet
- Applies a simple rule engine to flag employees as **"At Risk"** or **"Stable"**
- Automatically sends an email alert to HR/managers the moment someone is newly flagged
- Runs in real time using an on-edit trigger — no need to manually run anything

## ⚙️ How It Works

1. **Data Input** — Employee data lives in an `EmployeeData` sheet with columns: `Name | Department | Tenure (months) | Performance Score | Last Review Date`
2. **Rule Engine** — For every row, the script checks:
   ```
   IF Tenure < 6 months AND Performance Score < 3 → "At Risk"
   ELSE → "Stable"
   ```
3. **Status Column** — The result is written automatically into a new "Status" column
4. **Smart Email Alerts** — When an employee is newly flagged as "At Risk," an automatic email is sent containing their name, department, tenure, score, and reason for the flag
5. **Duplicate Prevention** — If an employee was already flagged "At Risk" in a previous check, no repeat email is sent — only new or newly-changed at-risk cases trigger an alert
6. **Real-Time Automation** — An installable "On Edit" trigger runs the script automatically whenever the sheet is updated, with a built-in safety check to skip incomplete rows

## 🛠️ Tech Stack

- Google Sheets (data storage + interface)
- Google Apps Script (automation logic + email integration)
- Gmail (`MailApp` service for alerts)

## 🚀 Why This Project

This was built to demonstrate practical, low-code process automation and AI-native problem solving in an HR context — turning a manual, reactive process (spotting at-risk employees) into a proactive, automated workflow. It reflects how small automation scripts can meaningfully reduce manual HR overhead and speed up intervention for at-risk talent.

## 📄 Code

```javascript
function flagAtRiskEmployees(e) {
  if (e && e.range.getColumn() === 6) return; // avoid re-triggering when script writes to Status column

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("EmployeeData");
  var data = sheet.getDataRange().getValues();

  sheet.getRange(1, 6).setValue("Status");

  for (var i = 1; i < data.length; i++) {
    var name = data[i][0];
    var department = data[i][1];
    var tenure = data[i][2];
    var score = data[i][3];
    var previousStatus = data[i][5];
    var status = "";

    if (tenure === "" || score === "") {
      continue;
    }

    if (tenure < 6 && score < 3) {
      status = "At Risk";

      if (previousStatus !== "At Risk") {
        sendRiskEmail(name, department, tenure, score);
      }
    } else {
      status = "Stable";
    }

    sheet.getRange(i + 1, 6).setValue(status);
  }
}

function sendRiskEmail(name, department, tenure, score) {
  var recipient = "your-email@example.com";
  var subject = "Attrition Risk Alert: " + name;
  var body = "Employee: " + name + "\n" +
             "Department: " + department + "\n" +
             "Tenure: " + tenure + " months\n" +
             "Performance Score: " + score + "\n" +
             "Reason: Low tenure combined with low performance score.";

  MailApp.sendEmail(recipient, subject, body);
}
```

## 🔮 Possible Future Improvements

- Add an AI-generated plain-English summary of all flagged employees using an LLM API
- Build a simple dashboard tab showing at-risk counts by department
- Add scheduled (time-based) triggers in addition to on-edit triggers

## 📝 Note

This project uses realistic **sample/dummy data** for demonstration purposes — no real employee data is used.

---

*Built as a demonstration of process automation and AI-native thinking for HR/People Analytics use cases.*
