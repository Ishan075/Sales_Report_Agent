import pandas as pd
import matplotlib.pyplot as plt
from fpdf import FPDF
import smtplib
import os
import sys
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from apscheduler.schedulers.background import BackgroundScheduler
from dateutil import parser
from datetime import datetime, timedelta
import threading

# -------------------- Configuration -------------------- #

# Email Configuration
EMAIL_ADDRESS = 'phoneyuser33@gmail.com'      # Replace with your email address
EMAIL_PASSWORD = 'lqeq hdqp ziih ttra'        # Replace with your email password or app-specific password
SMTP_SERVER = 'smtp.gmail.com'              # Replace with your SMTP server (e.g., 'smtp.gmail.com' for Gmail)
SMTP_PORT = 587                               # Replace with your SMTP port (e.g., 587 for Gmail)

# Sales Data Configuration
SALES_FILE_PATH = "C:\\Users\\91740\Downloads\\sales_data_sample.csv"           

# Report Configuration
REPORT_FILE_PATH = 'daily_sales_report.pdf'

# Management Team Emails
MANAGEMENT_EMAILS = ['phoneyuser33@gmail.com', 'manager2@example.com']  # Replace with actual email addresses

# ------------------------------------------------------- #

# Initialize the scheduler
scheduler = BackgroundScheduler()
scheduler.start()

def load_sales_data(file_path):
    """
    Load sales data from an Excel or CSV file.

    Args:
        file_path (str): Path to the sales data file.

    Returns:
        pd.DataFrame: DataFrame containing sales data.
    """
    try:
        if file_path.endswith('.xlsx') or file_path.endswith('.xls'):
            df = pd.read_excel(file_path)
        elif file_path.endswith('.csv'):
            df = pd.read_csv(file_path)
        else:
            print("Unsupported file format. Please provide an Excel or CSV file.")
            return None
        return df
    except Exception as e:
        print(f"Error loading sales data: {e}")
        return None

def generate_sales_summary(df):
    """
    Generate summary statistics from sales data.

    Args:
        df (pd.DataFrame): Sales data.

    Returns:
        dict: Dictionary containing summary statistics.
    """
    try:
        total_sales = df['Total Amount'].sum()
        total_quantity = df['Quantity'].sum()
        top_product = df.groupby('Product')['Total Amount'].sum().idxmax()
        summary = {
            'Total Sales': total_sales,
            'Total Quantity': total_quantity,
            'Top Product': top_product
        }
        return summary
    except Exception as e:
        print(f"Error generating sales summary: {e}")
        return {}

def generate_sales_charts(df):
    """
    Generate sales charts and save as images.

    Args:
        df (pd.DataFrame): Sales data.

    Returns:
        list: List of image file paths.
    """
    try:
        images = []

        # Sales by Product
        sales_by_product = df.groupby('Product')['Total Amount'].sum().sort_values(ascending=False)
        plt.figure(figsize=(10,6))
        sales_by_product.plot(kind='bar')
        plt.title('Sales by Product')
        plt.xlabel('Product')
        plt.ylabel('Total Sales')
        product_chart = 'sales_by_product.png'
        plt.savefig(product_chart)
        images.append(product_chart)
        plt.close()

        # Sales Over Time
        sales_over_time = df.groupby('Date')['Total Amount'].sum()
        plt.figure(figsize=(10,6))
        sales_over_time.plot(kind='line', marker='o')
        plt.title('Sales Over Time')
        plt.xlabel('Date')
        plt.ylabel('Total Sales')
        time_chart = 'sales_over_time.png'
        plt.savefig(time_chart)
        images.append(time_chart)
        plt.close()

        return images
    except Exception as e:
        print(f"Error generating sales charts: {e}")
        return []

def generate_pdf_report(summary, charts, output_path):
    """
    Generate a PDF report containing sales summary and charts.

    Args:
        summary (dict): Sales summary statistics.
        charts (list): List of chart image file paths.
        output_path (str): Path to save the PDF report.
    """
    try:
        pdf = FPDF()
        pdf.add_page()

        # Title
        pdf.set_font("Arial", 'B', 16)
        pdf.cell(0, 10, "Daily Sales Report", ln=True, align='C')

        # Summary
        pdf.set_font("Arial", 'B', 12)
        pdf.cell(0, 10, "Summary", ln=True)
        pdf.set_font("Arial", '', 12)
        for key, value in summary.items():
            pdf.cell(0, 10, f"{key}: {value}", ln=True)

        # Charts
        for chart in charts:
            pdf.add_page()
            pdf.image(chart, x=10, y=10, w=190)
            os.remove(chart)  # Clean up the image file

        # Save PDF
        pdf.output(output_path)
        print(f"PDF report generated at {output_path}")
    except Exception as e:
        print(f"Error generating PDF report: {e}")

def send_email(subject, body, attachments, recipients):
    """
    Send an email with attachments.

    Args:
        subject (str): Email subject.
        body (str): Email body.
        attachments (list): List of file paths to attach.
        recipients (list): List of recipient email addresses.
    """
    try:
        # Set up the MIME
        msg = MIMEMultipart()
        msg['From'] = EMAIL_ADDRESS
        msg['To'] = ', '.join(recipients)
        msg['Subject'] = subject

        # Attach the body with the msg instance
        msg.attach(MIMEText(body, 'plain'))

        # Attach files
        for file_path in attachments:
            with open(file_path, "rb") as attachment:
                part = MIMEApplication(attachment.read(), Name=os.path.basename(file_path))
            part['Content-Disposition'] = f'attachment; filename="{os.path.basename(file_path)}"'
            msg.attach(part)

        # Create SMTP session
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()  # Enable security
        server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)

        # Send email
        server.send_message(msg)
        server.quit()
        print("Email sent successfully.")
    except Exception as e:
        print(f"Error sending email: {e}")

def create_and_send_report():
    """
    Complete workflow to create and send the sales report.
    """
    print("Generating sales report...")
    df = load_sales_data(SALES_FILE_PATH)
    if df is None:
        print("Failed to load sales data.")
        return

    summary = generate_sales_summary(df)
    charts = generate_sales_charts(df)
    generate_pdf_report(summary, charts, REPORT_FILE_PATH)

    subject = "Daily Sales Report"
    body = "Please find attached the daily sales report."
    send_email(subject, body, [REPORT_FILE_PATH], MANAGEMENT_EMAILS)

def schedule_report(send_time):
    """
    Schedule the report to be sent at a specified time.

    Args:
        send_time (datetime): Datetime object specifying when to send the report.
    """
    scheduler.add_job(create_and_send_report, 'date', run_date=send_time)
    print(f"Report scheduled to be sent at {send_time}.")

def parse_time_input(time_input):
    """
    Parse user input time into a datetime object.

    Args:
        time_input (str): User input time string.

    Returns:
        datetime: Parsed datetime object.
    """
    try:
        # Handle relative time (e.g., "in 1 hour")
        if 'in' in time_input.lower():
            delta = parser.parse(time_input, fuzzy=True) - datetime.now()
            send_time = datetime.now() + delta
        else:
            # Handle absolute time (e.g., "at 9 PM")
            send_time = parser.parse(time_input, fuzzy=True)
            if send_time < datetime.now():
                # If the time is already past for today, schedule for tomorrow
                send_time += timedelta(days=1)
        return send_time
    except Exception as e:
        print(f"Error parsing time input: {e}")
        return None

def conversational_agent():
    """
    Simple command-line conversational agent.
    """
    print("Welcome to the Sales Report Assistant!")
    print("You can enter commands like:")
    print("- 'send report now'")
    print("- 'schedule report in 1 hour'")
    print("- 'schedule report at 9 PM'")
    print("- 'exit' to quit")

    while True:
        user_input = input("\nYou: ").strip().lower()

        if user_input in ['exit', 'quit', 'q']:
            print("Goodbye!")
            scheduler.shutdown()
            sys.exit()

        elif 'send' in user_input and 'now' in user_input:
            print("Processing immediate report generation and sending...")
            # Run in a separate thread to avoid blocking
            threading.Thread(target=create_and_send_report).start()

        elif 'schedule' in user_input and 'report' in user_input:
            # Extract the time part from the user input
            try:
                if 'in' in user_input:
                    time_str = user_input.split('in',1)[1].strip()
                elif 'at' in user_input:
                    time_str = user_input.split('at',1)[1].strip()
                else:
                    print("Could not parse the time. Please try again.")
                    continue

                send_time = parse_time_input(time_str)
                if send_time:
                    schedule_report(send_time)
                else:
                    print("Failed to parse the specified time.")
            except Exception as e:
                print(f"Error processing schedule command: {e}")

        else:
            print("Sorry, I didn't understand that command. Please try again.")

# -------------------- Main Execution -------------------- #

if __name__ == "__main__":
    conversational_agent()
