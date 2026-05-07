# Intern Attendance Register Automation

## Overview
An automated weekly attendance register system built for the CX Expert Intern Programme using Google Sheets and Google Apps Script. The system automatically generates, distributes and stores weekly attendance registers for 31 interns across three skill groups.

## Features
- Automatically creates a new Google Sheet every Monday at 7:00 AM SAST
- 31 interns listed and grouped by skill — IT, Marketing/Designing and Zendesk
- Clickable tick boxes for Attended and On Time for each day of the week
- Auto-counting totals row for daily attendance figures
- Emails the register link to the supervisor every Monday
- Saves all registers to Google Drive for record keeping
- Runs automatically from 04 May 2026 to 29 October 2026 (26 weeks)

## Technologies Used
- Google Apps Script (JavaScript)
- Google Sheets API
- Google Drive API
- Gmail API

## How It Works
1. Every Monday at 7:00 AM the script runs automatically
2. A new Google Sheet is created for that specific week
3. All 31 interns are populated with shift and skill group details
4. Clickable checkboxes are inserted for attendance tracking
5. The supervisor receives an email with a direct link to the register
6. The register is saved to a shared Google Drive folder

## Skills Demonstrated
- Process automation
- Google Workspace development
- Data tracking and reporting
- JavaScript / Apps Script programming
- Spreadsheet design and formatting
