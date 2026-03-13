# Report Automation Pipeline

## Overview

This project demonstrates an automated reporting pipeline built using Power BI and Power Automate.

The system refreshes BI datasets, triggers a scheduled workflow, generates a formatted report, and delivers it to stakeholders automatically.

The goal is to eliminate manual reporting effort and ensure consistent report delivery.

---

## Problem

Many reporting workflows require manual steps such as:

• refreshing dashboards
• exporting data
• formatting reports
• emailing stakeholders

These manual processes are time-consuming and prone to delays.

---

## Solution

The system automates the reporting workflow using a scheduled pipeline.

Once the dataset refresh is complete, a Power Automate flow retrieves the latest report data and delivers the formatted output via email.

---

## Architecture

Power BI Dataset Refresh 
↓
Power Automate Scheduled Flow 
↓
Fetch Report Data
↓
Write Output to Excel
↓
Generate Formatted Report
↓
Send Email to Stakeholders

---

## Workflow Steps

1. **Power BI Dataset Refresh**

   The reporting dataset refreshes automatically on a scheduled time.

2. **Power Automate Trigger**

   A single Power Automate workflow runs after the dataset refresh.

3. **Fetch Data from Power BI**

   The flow retrieves the required report dataset.

4. **Generate Output**

   Data is written to Excel and formatted into the final report structure.

5. **Email Delivery**

   The formatted report is sent to stakeholders via automated email.

---

## Tech Stack

Power BI
Power Automate
Excel Online
Email Automation

---

## Benefits

• Eliminates manual report generation
• Ensures same-day report delivery
• Reduces operational overhead
• Improves reporting reliability
