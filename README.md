Overview

This Python project automates the process of generating a daily sales report from a spreadsheet and emailing it to the management team. The agent can either send the report immediately when requested or schedule the mailing at a specified time.

Features

Extract Data from Spreadsheet:

The agent can load sales data from CSV or Excel files, automatically identifying the relevant columns.


Generate a Daily Sales Report in PDF Format:

The agent dynamically summarizes the sales data and generates a detailed PDF report, including summary statistics, charts, and a detailed table of all sales data.


Email the Report:

The agent can send the generated PDF report to the management team via email.


Scheduled or On-Demand Report Sending:

The agent can send the report immediately when asked or schedule the mailing at a specified time (e.g., after 1 hour or at 9 PM).


Project Structure:

sales_report_agent.py: Main script that handles data extraction, report generation, and email sending.


README.md: Documentation file (this file).

requirements.txt: List of dependencies required to run the script.

Requirements:
The following Python libraries are required:

pandas
matplotlib
fpdf
smtplib
apscheduler
email
threading
os
sys

You can install these dependencies by running:
pip install -r requirements.txt

Usage
1. Configure Email Settings
Before running the script, configure the sender's email settings in the send_email_report function inside the sales_report_agent.py file:
sender_email = "your_email@example.com"
password = "your_password"

2. Specify the Sales Data File
Specify the path to your sales data file in the main function:
SALES_FILE_PATH = "path_to_your_sales_data.csv"

3. Run the Script
Run the script to generate and send the sales report:
python sales_report_agent.py

4. Scheduled Report Sending
To schedule the report for a specific time, you can integrate the script with a task scheduler like cron (on Linux) or Task Scheduler (on Windows).

Code Explanation

Load Sales Data:

The script loads sales data from a CSV or Excel file using the pandas library. The data is stored in a DataFrame for easy manipulation.


Generate Sales Summary:

The script dynamically generates a summary of key statistics (e.g., total sales, total quantity) based on the available columns in the dataset.


Generate Sales Charts:

If the dataset includes relevant columns, the script generates charts (e.g., sales by product line, sales over time) using matplotlib.


Generate PDF Report:

The script creates a PDF report using the FPDF library. The report includes a summary, charts, and a detailed table of the sales data.


Send Email Report:

The generated PDF report is attached to an email and sent to the management team using the smtplib library.
