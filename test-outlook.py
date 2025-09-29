from RPA.Outlook.Application import Application
import logging

harjoitus_name = "outlook-test"

attachment_file = r"D:\CAREERIA\it-opiskelu\2025-syksy\robot\rpa-p312\p1-h4_log_simple.txt"

# Set up logging file to log all ui events with timestamps
log_filename = harjoitus_name + "_log.txt"
logging.basicConfig(
    filename=log_filename,
    level=logging.INFO,
    format="%(asctime)s %(message)s",
    encoding="utf-8",
    filemode="w",
)  # Overwrite log file each run


def send_email():
    app = Application()
    app.open_application()
    # app.send_email(
    #     recipients='EMAILADDRESS_1, EMAILADDRESS_2',
    #     subject='email subject',
    #     body='email body message',
    #     attachments='../orders.csv')
    
    app.send_email(
        recipients='kennet.myllykoski@gmail.com',
        subject='RPA Framework test email with attachment',
        body='Hello! This is a test email sent using RPA Framework. \nSee the attachment.',
        attachments=attachment_file
        )
    
    print("Email sent!")

if __name__ == "__main__":
    send_email()
    logging.info("Email sent!")
    print("Email sent!")
    logging.info("****** DONE ******")
    print("****** DONE ******")
