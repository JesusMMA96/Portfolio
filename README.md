# 💼 SAP Payments & Reporting Automation Suite

A modular Python toolkit for automating payments and generating reports within SAP GUI environments using Excel and PyQt integration.

---

## 📌 Project Scope

This project streamlines SAP payment processing and reporting by:
- Gathering user inputs via custom PyQt dialogs
- Generating SAP payment scripts
- Classifying and filtering Excel-based financial data
- Automating SAP GUI workflows for efficient execution

---

## ⚙️ Workflow Overview
1. **User Input Module**  
   Collects payment/report parameters from user via PyQt dialogs.
2. **SAP Auxiliary Functions**  
   Handles navigation, data extraction and load data through SAP GUI scripting.
3. **Daily Payments**  
   Applies daily transactions into SAP based on bank movement files.
4. **Payments Module**  
   Categorizes different types of payments, processes detail files, and automates their entry into SAP.
5. **Reports Module**  
   Generates:
   - Aging debt reports
   - Annual balance sheets
   - Large retailer summaries
6. **Utilities Module**  
   Advanced Excel manipulation using `xlwings` and the powerful Excel COM API.
7. **Config Loader**  
   Loads SAP environment info from a secure JSON configuration file *(placeholders used to protect sensitive data).*
8. **Main Module**  
   Application entry point. Loads UI and enables module selection.
9. **SAP Info JSON**  
   Contains SAP codes, cost centers, company data, and template paths *(sensitive information withheld).*

---

## 🧠 Innovation Highlights
- Seamless integration of **PyQt**, **Excel COM API**, and **SAP GUI Scripting**
- Modular architecture enabling easy maintenance and expansion
- Built-in classification, formatting, and automation for complex payment workflows
- Robust Excel manipulation toolkit for high-volume data processing

---

## 📊 Impact Metrics (Sample Results)

| Metric                        | Before Automation | After Automation |
|------------------------------|-------------------|------------------|
| Daily payment entry time     | ~2 hours          | ~20 minutes      |
| Report preparation duration  | ~4 hours          | ~45 minutes      |
| Error rate in transactions   | ~5%               | <0.5%            |

---

## 📦 Tech Stack
**Languages & Tools:**  
Python, SAP GUI Scripting, PyQt5, xlwings, JSON, Excel COM API

---

## 🚀 Future Enhancements
- Optimize performance when creating debt report
- Add logging and error recovery for SAP GUI automation
- Introduce role-based access control for sensitive data files
- Integrate Power BI or dashboards for report visualization

---

## 📝 Lessons Learned
- SAP GUI scripting requires careful timing and UI state validation.
- PyQt adds an intuitive front-end but requires meticulous signal management.
- Excel COM API offers incredible flexibility but demands disciplined memory and object management.

---

Feel free to reach out.
