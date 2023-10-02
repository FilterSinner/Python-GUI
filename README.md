# Python-GUI

# IT Asset Management System

## Overview

This project is an IT Asset Management System developed in Python with a graphical user interface (GUI) using the tkinter library. It allows users to efficiently manage IT assets, generate PDF reports, and send email notifications with asset details. Additionally, it includes a web service for acknowledging receipt of emails.

## Features

- **User Interface (main.py)**:
  - Easy-to-use GUI for entering employee and IT asset information.
  - Calendar widget for selecting verification dates.
  - Listbox for displaying the list of entered assets.
  
- **Data Entry**:
  - Enter employee information, including name, employee number, department, email, and verifier name.
  - Add IT assets to the list, specifying item type, serial number, asset code, and remarks.
  - Clear asset entry fields as needed.

- **PDF Generation**:
  - Generate PDF reports with asset details.
  - Customizable PDF reports with company logos.
  - Save generated PDFs to a user-specified location.

- **Email Sending**:
  - Send email notifications to employees with attached PDF reports.
  - Group assets by employee email address and create consolidated PDFs.
  - Customize email content using an HTML template.
  - Handle SMTP email sending with error handling.

- **Acknowledge Functionality (data_retrieval.py)**:
  - Web service for acknowledging receipt of emails.
  - Listens on port 80 and records acknowledgment status and IP addresses.
  - Acknowledgment data is stored in an Excel file.

## Usage

1. Launch the IT Asset Management System GUI by running `main.py`.
2. Enter employee information and asset details.
3. Generate PDF reports with the "Print" button.
4. Send email notifications with attached reports using the "Send Email" button.
5. Employees can acknowledge receipt of emails by clicking on the provided links.

## Acknowledgment

The project includes an HTML template for email content. Feel free to customize it further.

## Additional Notes

- Make sure to configure SMTP settings for sending emails (SMTP server, port, username, password).
- Replace the provided email templates and logo paths with your organization-specific content.




