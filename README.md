**GreenPath Diagnostic Billing & Report Software**
A comprehensive desktop application built with Python and CustomTkinter for managing diagnostic laboratory billing, patient records, and automated medical report generation.

🚀 Features
Billing System: Generate professional invoices with automated GST/Discount calculations.

Report Module: Dynamic medical report generation using Word (.docx) templates.

Database Management: Manage lists of Tests, Doctors, and Agents.

Patient History: Track all past bills and payments via Excel integration.

PDF Conversion: Automated conversion of invoices and reports to PDF for easy sharing and printing.

📂 Project Structure & Required Files
To run this project successfully, your folder structure should look like this. If any folder is missing, the application will attempt to create some (like bill or GeneratedReports), but template folders must be created manually before use.


green-path-billing/
├── main.py                     # The primary application script
├── logo.png                    # Your lab logo (used in sidebar and invoices)
├── settings_icon.png           # UI Icon for settings
├── moon_icon.png               # UI Icon for Dark Mode
├── sun_icon.png                # UI Icon for Light Mode
│
├── ReportTemplates/            # [REQUIRED] Store your .docx report templates here
│   └── CBC NEW 2025.docx       # Example template referenced in code
│
├── GeneratedReports/           # Automated: Stores generated medical reports
├── bill/                       # Automated: Stores generated invoices (.docx & .pdf)
│
├── test_amount.txt             # Data: Stores "Test Name - Price"
├── doctors.txt                 # Data: Stores list of referred doctors
├── agents.txt                  # Data: Stores list of agents
├── bill_counter.txt            # Data: Tracks the next Bill Number (e.g., GPDL0001)
├── bills.xlsx                  # Data: The main database for patient history
└── custom_reports.json         # Data: Stores configurations for dynamic reports

🛠️ Prerequisites & Installation
1. Install Python
Ensure you have Python 3.10 or higher installed.

2. Install Dependencies
Run the following command in your terminal to install all required libraries:

**pip install customtkinter Pillow pandas openpyxl python-docx comtypes docxtpl PyPDF2**

3. System Requirements
Windows OS: This app uses comtypes.client to communicate with Microsoft Word for PDF conversion. Microsoft Word must be installed on the system.

📖 How to Use
Initialize Data: Create test_amount.txt and add your tests in the format: Glucose - 150.

Run the App: Execute python** main.py**.

Billing: * Enter patient details.

Search and add tests (autofill prices).

Click Print Invoice to save the Excel record and open the PDF.

Reports:

Go to the Report tab in the sidebar.

Search for a patient by Bill Number.

Select a report type (e.g., CBC) to autofill patient data into the medical template.
