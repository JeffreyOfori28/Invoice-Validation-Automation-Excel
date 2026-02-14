# 📊 Invoice Validation Automation (Excel + VBA)

# Overview

This project automates invoice validation and vendor-level performance reporting using Microsoft Excel, Pivot Tables, and VBA. 

The solution reduces manual review by applying margin-based approval rules and automatically refreshing vendor summaries when invoice data is updated.

# 🎯 Business Objective
Finance and operations teams often review supplier invoices manually to:

* Validate pricing and margin thresholds
* Identify exceptions
* Monitor vendor performance

This project simulates a real-world invoice control workflow and automates key validation steps.

# 🛠 Tools & Technologies

* Microsoft Excel
* Pivot Tables
* Logical formulas (IF, ROUND, GETPIVOTDATA)
* VBA (Macro automation)

# 🔎 Key Features
1️⃣ Automated Invoice Validation
* Margin-based approval logic flags pricing exceptions 
* Logical formulas automatically classify invoices as Approved or Not Approved

2️⃣ Dynamic Vendor Master Summary
Pivot table aggregates:
* Average approved quantity
* Average approved margin
Vendor summary updates automatically when invoice data changes

3️⃣ VBA Refresh Automation
* Custom macro clears old invoice data
* Imports updated data
* Executes:
ThisWorkbook.RefreshAll

Ensures real-time pivot and summary recalculation

# 🚀 How to Use

1. Open Invoice_Validation_Automation.xlsm
2. Enable macros
3. Click the Import & Update Invoice button
4. Vendor master updates automatically

# 📌 Skills Demonstrated

* Excel automation
* Financial data validation
* Exception handling logic
* Pivot table reporting
* VBA macro development
* Process improvement mindset

# 📷 Screenshots
See /screenshots folder for workflow visuals.
