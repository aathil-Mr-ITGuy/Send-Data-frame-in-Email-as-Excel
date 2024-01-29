import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
from datetime import date
import time
import pandas as pd

def sendSuccessEmail(dataframe):
    """
    Sends an email with a DataFrame attached as an Excel file.

    This function sends an email with the provided DataFrame attached as an Excel file.
    It constructs the email message, attaches the DataFrame as an Excel file, and sends
    the email using the configured SMTP server.

    Args:
        dataframe (pandas.DataFrame): The DataFrame to be attached to the email.

    Raises:
        Exception: An error occurred while sending the email.

    """
    try:
        # Set the email parameters
        from_address = "dummy.sender@example.com"
        to_address = "dummy.recipient@example.com"
        cc_addresses = ["dummy.cc1@example.com", "dummy.cc2@example.com", "dummy.cc3@example.com"]
        subject = "Dummy Subject: Sample Data"
        message = "Dear Recipient,\n\nPlease find the requested Sample Data attached with this mail.\n\nThanks,\nDummy Sender"

        # Create the email message
        msg = MIMEMultipart()
        msg['From'] = from_address
        msg['To'] = to_address
        msg['Cc'] = ", ".join(cc_addresses)
        msg['Date'] = formatdate(localtime=True)
        msg['Subject'] = subject

        msg.attach(MIMEText(message))

        # Convert the DataFrame to Excel
        file_name = f"Sample Data {date.today().strftime('%d-%m-%Y')}.xlsx"
        dataframe.to_excel(file_name, index=False)

        # Read the Excel file you want to attach
        with open(file_name, "rb") as excel_file:
            # Create a MIMEBase object to attach the file
            part = MIMEBase("application", "octet-stream")
            part.set_payload(excel_file.read())

        # Encode the payload using Base64
        encoders.encode_base64(part)

        # Add header with the file name
        part.add_header("Content-Disposition", f"attachment; filename= {file_name}")

        # Add the attachment to the email message
        msg.attach(part)

        # Connect to the SMTP server and send the email
        smtp_server = smtplib.SMTP('smtp.office365.com', 587)
        smtp_server.ehlo()
        smtp_server.starttls()
        smtp_server.ehlo()
        smtp_server.login(from_address, "EMAIL_PW")
        recipient_list = [to_address] + cc_addresses
        smtp_server.sendmail(from_address, recipient_list, msg.as_string())
        smtp_server.close()
    except Exception as e:
        time.sleep(5)
        print(f"Error in sending mail: {e}")
        raise

# Example usage
df = pd.read_csv("https://support.staffbase.com/hc/en-us/article_attachments/360009197031/username.csv")
sendSuccessEmail(df)
