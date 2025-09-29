from RPA.Outlook.Application import Application
import logging
import pandas as pd
import datetime
import os

# Päivä 5
# Lopputehtävä 
#
# RPA-sovellus, joka lukee saapuneista sähköposteista tietyltä lähettäjältä ja tietyllä otsikolla (subject) olevat sähköpostit.
#
# Sähköpostien tiedot viedään Exceliin: pvm, klo, aihe, viesti, luettelo liitetiedostoista
# Ohjelma laskee kuinka monta viestiä ko. lähettäjältä on tullut.
# Lopuksi se lähettää ao. viestin sekä itsellesi, että ko. henkilölle siten, että listan alussa on HOKS-aihetta käsittelevät viestit ja sen jälkeen muut viestit. 
# Voit rajata haettavien viestien määrän asetuksissa
# Esimerkki:
# Vastaanottaja: input-kyselyssä annettu sähköpostiosoite

# Aihe: RPA-lopputyö: OmaEtunimi OmaSukunimi

# Viesti: Opettaja Etu Suku on lähettänyt minulle Etu Suku xx kpl sähköpostiviestejä. Tässä lista viesteistä:

# EMAIL-viestejä 9 kappaletta:

# -Viesti 1 Päivämäärä: dd.mm.yyyy, Aihe: zzzz, Liitetiedostot: HOKS.xlsx, suunn.txt

# -Viesti 2 Päivämäärä: dd.mm.yyyy, Aihe: zzzz, Liitetiedostot: HOKS.xlsx

# email_sender = input("Anna sähköpostiosoite jolta viestit haetaan (esim. opettajan): ").strip()
# email_outlook_account = input("Anna oma sähköpostiosoitteesi (johon viesti tästä RPA ohjelmasta myös lähetetään): ").strip()

# email_sender = "aaa.bbb@careeria.fi"
email_sender = "ccc.dddd@careeria.fi"
email_outlook_account = "eeee.ffff@student.careeria.fi"
email_subject_keyword = "HOKS"
# email_report_recipients = f"{email_outlook_account}"  # when debugging send only to self
email_report_recipients = f"{email_sender}; {email_outlook_account}"

harjoitus_name = "p5-h-final"

log_simple_filename = harjoitus_name + "_log_simple.txt"

program_start_time = datetime.datetime.now()

current_dir = os.path.dirname(os.path.abspath(__file__))

# Set up logging file to log all RPA events with timestamps
log_filename = harjoitus_name + "_log.txt"
logging.basicConfig(
    filename=log_filename,
    level=logging.INFO,
    format="%(asctime)s %(message)s",
    encoding="utf-8",
    filemode="w",
)  # Overwrite log file each run


# Simple logger that uses diffrent logging file than general logging
# and is used only in the write_to_simple_log function
def get_simple_logger():
    logger = logging.getLogger("simple_logger")
    logger.setLevel(logging.INFO)
    # Prevent propagation to root logger
    logger.propagate = False
    if not logger.handlers:
        handler = logging.FileHandler(log_simple_filename, mode="w", encoding="utf-8")
        formatter = logging.Formatter("%(asctime)s %(message)s")
        handler.setFormatter(formatter)
        logger.addHandler(handler)
    
    return logger


def write_to_log(s):
    logging.info(s)
    write_to_simple_log(s)  # Log to simple log as well


def write_to_simple_log(s):
    logger = get_simple_logger()
    logger.info(s)


def get_firstname_lastname_from_email(email):
    try:
        local_part = email.split('@')[0]
        parts = local_part.split('.')
        if len(parts) >= 2:
            firstname = parts[0].capitalize()
            lastname = parts[-1].capitalize()  # Last part as last name if more than two parts
            return firstname, lastname
        else:
            return parts[0].capitalize(), ""
    except Exception as e:
        write_to_log(f"Error extracting names from email {email}: {e}")
        return None, None


def rpa_outlook_main():
    write_to_log(f"Aloitusaika: {program_start_time.strftime('%d.%m.%Y %H:%M:%S')}")

    try:
        outlook_app = Application()
        outlook_app.open_application()
        write_to_log("Outlook application opened.")

        # From RPA.Outlook.Application documentation:
        # https://rpaframework.org/libdoc/RPA_Outlook_Application.html
        # Get all emails from the specified sender
        #
        # account_name:	    needs to be given if there are shared accounts in use, defaults to None
        # folder_name:	    target folder where to get emails from, default Inbox
        # email_filter:	    how to filter email, default no filter, ie. all emails in folder
        #                   ex. "[Sender] = 'john.doe@example.com'"  https://learn.microsoft.com/en-us/office/vba/api/outlook.items.restrict
        #                   OBS!!! With the email_filter show here in documentation seems not to work.
        #                   The workaround is to get all emails and then filter them in pandas DataFrame.
        # save_attachments:
        #     if attachments should be saved, defaults to False
        # attachment_folder:
        #     target folder where attachments are saved, defaults to current directory
        # sort:	            if emails should be sorted, defaults to False
        # sort_key:	        needs to be given if emails are to be sorted
        # sort_descending:  set to False for ascending sort, defaults to True
        # return:	        list of emails (list of dictionaries)
        # emails = outlook_app.get_emails(folder="Inbox", filter=f'[Sender] ="{email_sender}" AND subject:"{email_subject_keyword}"', account=email_outlook_account)
        emails = outlook_app.get_emails(folder_name="Saapuneet", 
                                        # email_filter=f"[SenderEmailAddress] = '{email_sender}'", # got none with this?
                                        # email_filter=f"[Sender] = '{email_sender}'", # got 'Invalid email filter' with this?
                                        # save_attachments=True, 
                                        # attachment_folder="./attachments",
                                        account_name=email_outlook_account)
        
        if len(emails) == 0:
            write_to_log(f"No emails found from {email_sender}")
            print(f"No emails found from {email_sender}")
            return
        else:
            pass
            # This does not work as expected as the email_filter does not work
            # write_to_log(f"Total emails found from {email_sender}: {len(emails)}")
            # print(f"Total emails found from {email_sender}: {len(emails)}")
        
        email_outlook_account_firstname, email_outlook_account_lastname = get_firstname_lastname_from_email(email_outlook_account)
        email_sender_account_firstname, email_sender_account_lastname = get_firstname_lastname_from_email(email_sender)

        df = pd.DataFrame(emails)
        df = df.drop(columns=['object']) # Drop the 'object' column as unneeded 

        # drop rows where 'Sender' is not email_sender
        df = df[df['Sender'].str.lower() == email_sender.lower()]
        df = df.reset_index(drop=True)  # Reset index after filtering
        write_to_log(f"Total emails after filtering by sender {email_sender}: {len(df)}")
        print(f"Total emails after filtering by sender {email_sender}: {len(df)}")
        if len(df) == 0:
            write_to_log(f"No emails found from {email_sender} after filtering.")
            print(f"No emails found from {email_sender} after filtering.")
            return

        # Reformat the 'ReceivedTime' column orginally a string to separate date and time columns
        # ex. 'ReceivedTime': '2025-09-16T11:49:08.030000+00:00',
        # Convert 'ReceivedTime' to datetime
        df['ReceivedTime'] = pd.to_datetime(df['ReceivedTime'])
        df['ReceivedTime'] = df['ReceivedTime'].dt.tz_localize(None)  # Remove timezone info if present, needed for Excel export
        # Create 'Date' and 'Time' columns
        df['Date'] = df['ReceivedTime'].dt.strftime('%d.%m.%Y')
        df['Time'] = df['ReceivedTime'].dt.strftime('%H:%M:%S')


        # Sort the DataFrame so that emails with the specified subject keyword are first
        # Create a helper column that is True if the subject contains the keyword, False otherwise
        df['SubjectKeyword'] = df['Subject'].apply(lambda x: email_subject_keyword.lower() in x.lower())
        df = df.sort_values(by=['SubjectKeyword', 'ReceivedTimestamp'], ascending=[False, True])
        df = df.drop(columns=['SubjectKeyword'])  # Drop the helper column
        df = df.reset_index(drop=True)  # Reset index after sorting
        write_to_log("Emails sorted with subject keyword first, secondly on ReceivedTimestamp.")
        print("Emails sorted with subject keyword first, secondly on ReceivedTimestamp.")

        # Convert Attachments from a list to a string listing filenames
        df['Attachments'] = df['Attachments'].apply(
            lambda x: ', '.join([att['filename'] for att in x]) if x else 'None'
        )

        # Save the DataFrame to an Excel file
        # Only the following columns are needed: Date, Time, Subject, Body, Attachments. Other columns are dropped.
        df = df[['Date', 'Time', 'Subject', 'Body', 'Attachments']]
        excel_filename = f'{email_sender_account_firstname}_{email_sender_account_lastname}_emails.xlsx'
        try:
            df.to_excel(excel_filename, index=False)
        except Exception as e:
            write_to_log(f"Error saving to Excel file {excel_filename}: {e}")
            print(f"Error saving to Excel file {excel_filename}: {e}")
            if 'Permission denied:' in str(e):
                print("Excel tiedosto on auki, sulje se ja yritä uudelleen.")
            return
        write_to_log(f"Emails saved to Excel file: {excel_filename}")
        print(f"Emails saved to Excel file: {excel_filename}")

        # Send email that the run is done with summary information and log file as attachment
        email_subject = f"RPA-lopputyö: {email_outlook_account_firstname} {email_outlook_account_lastname}"
        email_body = f'''Opettaja {email_sender_account_firstname} {email_sender_account_lastname} on lähettänyt minulle {len(df)} kpl sähköpostiviestejä.\n\n'''
        email_body += f'''Tässä lista EMAIL-viesteistä {len(df)} kappaletta:\n\n'''

        for index, row in df.iterrows():
            email_body += f'''- Viesti {index + 1:3} Päivämäärä: {row['Date']}, Aihe: {row['Subject']}'''
            if row['Attachments']:
                email_body += f''', Liitetiedostot: {row['Attachments']}\n'''
            else:
                email_body += ''', Liitetiedostot: (ei liitteitä)\n'''
        
        program_end_time = datetime.datetime.now()
        program_duration = program_end_time - program_start_time

        email_body += f'''\nAjettu: {program_start_time.strftime('%d.%m.%Y %H:%M:%S')}\n'''
        email_body += f'''Kesto: {str(program_duration)}\n'''
        email_body += f'''LIITTEET: {excel_filename}\n'''
        email_body += "\n Lähetetty RPA-ohjelmalla.\n"

        attachment_file = os.path.join(current_dir, excel_filename)

        try:
        
            # From RPA.Outlook.Application documentation:
            # https://rpaframework.org/libdoc/RPA_Outlook_Application.html
            # Send email with Outlook parameters:
            # recipients:	    list of addresses, ex. 'EMAILADDRESS_1; EMAILADDRESS_2'
            #                   OBS!!! a string with ; separated email addresses NOT , separated as in documentation example
            # subject:	        email subject
            # body:	            email body
            # html_body:	    True if body contains HTML, defaults to False
            # attachments:	    list of filepaths to include in the email, defaults to []
            # save_as_draft:	email is saved as draft when True
            # cc_recipients:	list of addresses for CC field, default None
            # bcc_recipients:	list of addresses for BCC field, default None
            # reply_to:	        list of addresses for changing email's reply-to field, default None
            # check_names:	    all recipients are checked if the email address is recognized on True, default False
            # return:	        True if there were no errors

            return_value = outlook_app.send_email(
                recipients=email_report_recipients,
                subject=email_subject,
                body=email_body,
                attachments=attachment_file
                )
            
            if return_value:  # True if email was sent successfully
                write_to_log(f"Email sent successfully with return value {return_value}")
                print(f"Email sent successfully with return value: {return_value}")
            else:
                write_to_log(f"Email sending failed with no exceptions but return value {return_value}")
                print(f"Email sending failed with no exceptions but return value: {return_value}")
                print("The reason could be that Outlook does not recognize the recipient email addresses because they are not separated by semicolons (;).")
                print("Ex. recipients='EMAILADDRESS_1; EMAILADDRESS_2'")

            write_to_log(f"Lopetusaika: {program_end_time.strftime('%d.%m.%Y %H:%M:%S')}")
            write_to_log(f"Kesto: {str(program_duration)}")

        except Exception as e:
            write_to_log(f"Email sending failed: {e}")
            print(f"Email sending failed: {e}")
        
    finally:
        write_to_log("Sovellus suljettu.")


if __name__ == "__main__":
    rpa_outlook_main()   
    write_to_log("****** DONE ******")
    print("****** DONE ******")


