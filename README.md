# Email Sender Script

This script sends an email with a DataFrame attached as an Excel file.

## Requirements

- Python 3.x
- pandas
- smtplib
- email

## Installation

No installation is required beyond having Python installed on your system. Ensure that the required packages (`pandas`, `smtplib`, `email`) are installed.

## Usage

1. Import the necessary modules:

    - `smtplib`: Import the Simple Mail Transfer Protocol (SMTP) library for sending emails.
    - `MIMEMultipart`: Import MIMEMultipart for constructing email messages with attachments.
    - `MIMEBase`: Import MIMEBase for creating base MIME objects for email attachments.
    - `MIMEText`: Import MIMEText for adding text content to email messages.
    - `formatdate`: Import formatdate for formatting dates in email headers.
    - `encoders`: Import encoders for encoding email attachments.
    - `date`: Import date for working with dates.
    - `time`: Import time for handling delays or waiting periods.
    - `pandas as pd`: Import pandas for working with DataFrames.
      
      ```python
      import smtplib
      from email.mime.multipart import MIMEMultipart
      from email.mime.base import MIMEBase
      from email.mime.text import MIMEText
      from email.utils import formatdate
      from email import encoders
      from datetime import date
      import time
      import pandas as pd

2. Define the function `sendSuccessEmail(dataframe)`:

    This function sends an email with the provided DataFrame attached as an Excel file.
    It constructs the email message, attaches the DataFrame as an Excel file, and sends
    the email using the configured SMTP server.

    Args:
        - `dataframe` (`pandas.DataFrame`): The DataFrame to be attached to the email.

    Raises:
        - `Exception`: An error occurred while sending the email.
   ```python
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
    # Function implementation...
3. Inside the function:
   ```python
       try:
        # Set the email parameters
        from_address = "dummy.sender@example.com"  # Set the sender's email address.
        to_address = "dummy.recipient@example.com"  # Set the recipient's email address.
        cc_addresses = ["dummy.cc1@example.com", "dummy.cc2@example.com", "dummy.cc3@example.com"]  # Set the CC email addresses.
        subject = "Dummy Subject: Sample Data"  # Set the subject of the email.
        message = "Dear Recipient,\n\nPlease find the requested Sample Data attached with this mail.\n\nThanks,\nDummy Sender"  # Set the email message.

        # Create the email message
        msg = MIMEMultipart()  # Create a new MIMEMultipart object for composing the email message.
        msg['From'] = from_address  # Set the sender's email address in the message headers.
        msg['To'] = to_address  # Set the recipient's email address in the message headers.
        msg['Cc'] = ", ".join(cc_addresses)  # Set the CC email addresses in the message headers.
        msg['Date'] = formatdate(localtime=True)  # Set the current date and time in the message headers.
        msg['Subject'] = subject  # Set the subject of the email in the message headers.

        msg.attach(MIMEText(message))  # Attach the email message as text to the MIMEMultipart object.

        # Convert the DataFrame to Excel
        file_name = f"Sample Data {date.today().strftime('%d-%m-%Y')}.xlsx"  # Generate a file name based on the current date.
        dataframe.to_excel(file_name, index=False)  # Convert the DataFrame to an Excel file.

        # Read the Excel file you want to attach
        with open(file_name, "rb") as excel_file:  # Open the Excel file in binary mode for reading.
            # Create a MIMEBase object to attach the file
            part = MIMEBase("application", "octet-stream")  # Create a new MIMEBase object with the specified content type and subtype.
            part.set_payload(excel_file.read())  # Set the payload of the MIMEBase object to the content of the Excel file.

        # Encode the payload using Base64
        encoders.encode_base64(part)  # Encode the payload of the MIMEBase object using Base64.

        # Add header with the file name
        part.add_header("Content-Disposition", f"attachment; filename= {file_name}")  # Add a header specifying the filename of the attachment.

        # Add the attachment to the email message
        msg.attach(part)  # Attach the MIMEBase object (attachment) to the MIMEMultipart object (email message).

        # Connect to the SMTP server and send the email
        smtp_server = smtplib.SMTP('smtp.office365.com', 587)  # Connect to the SMTP server using the specified host and port.
        smtp_server.ehlo()  # Send the EHLO command to the SMTP server to identify the client.
        smtp_server.starttls()  # Upgrade the connection to use Transport Layer Security (TLS) encryption.
        smtp_server.ehlo()  # Send the EHLO command again after starting TLS.
        smtp_server.login(from_address, "EMAIL_PW")  # Log in to the SMTP server using the sender's email address and password.
        recipient_list = [to_address] + cc_addresses  # Combine the recipient and CC email addresses into a single list.
        smtp_server.sendmail(from_address, recipient_list, msg.as_string())  # Send the email message as a string.
        smtp_server.close()  # Close the connection to the SMTP server.

    except Exception as e:  # Catch any exceptions that occur during the execution of the try block.
        time.sleep(5)  # Pause execution for 5 seconds.
        print(f"Error in sending mail: {e}")  # Print the error message.
        raise  # Raise the caught exception to propagate it to the caller.
4. Example usage:
   ```python
   df = pd.read_csv("https://support.staffbase.com/hc/en-us/article_attachments/360009197031/username.csv")  # Read a CSV file into a pandas DataFrame.
   sendSuccessEmail(df)  # Call the sendSuccessEmail function to send the email with the DataFrame attached.

### Function Purpose:
The `sendSuccessEmail` function is designed to send an email with a DataFrame attached as an Excel file. It's commonly used in data processing and reporting pipelines where automated email notifications are required.

### How It Works:

1. **Setting Email Parameters:**
   - The function starts by defining parameters such as the sender's email address (`from_address`), recipient's email address (`to_address`), CC email addresses (`cc_addresses`), subject of the email (`subject`), and the email message (`message`).

2. **Creating Email Message:**
   - It then creates an email message object using `MIMEMultipart` and sets various headers such as From, To, Cc, Date, and Subject using the `msg` object.

3. **Attaching Email Message:**
   - The message content (`message`) is attached to the email using `MIMEText` and added to the `msg` object.

4. **Attaching DataFrame as Excel File:**
   - The DataFrame (`dataframe`) is converted to an Excel file and attached to the email. The Excel file is created with a name based on the current date.

5. **Encoding and Adding Attachment:**
   - The Excel file is encoded using Base64 and added as an attachment to the email message.

6. **Sending the Email:**
   - The function connects to the SMTP server (`smtp.office365.com` in this case) and authenticates using the sender's email address and password. It then sends the email to the recipient and CC email addresses.

7. **Error Handling:**
   - The function includes error handling to catch any exceptions that occur during the email sending process. If an error occurs, it prints the error message and raises an exception to propagate it to the caller.

### Conclusion:
In summary, the `sendSuccessEmail` function facilitates the automated sending of emails with DataFrame attachments. It encapsulates the process of constructing email messages, converting DataFrames to Excel files, and handling the email sending process using the SMTP protocol.
## Configuration
Ensure that the email parameters (`from_address`, `to_address`, `cc_addresses`, `subject`, `message`, `EMAIL_PW`) are properly configured before running the script. Update these values according to your requirements.

## Contributing
Feel free to contribute by submitting issues or pull requests.


