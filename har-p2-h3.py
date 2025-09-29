from RPA.Windows import Windows
from RPA.Outlook.Application import Application
import logging
import pandas as pd
import datetime
import os

# Päivä 2
# Harjoitus 3
#
# Luetaan henkilötietoja tiedostosta ja tallennetaan niitä käyttöliittymään.
# Tietolähteenä XLSX-tyyppinen Excel-tiedosto.
# Lokiin tallennetaan alkuaika, loppuaika, ajon kesto sekä käsiteltyjen rivien määrä. 
# Alku- ja loppuaikojen muotoilut suomalaisilla pvm-muotoiluilla.
#
# Onnistuneen ajon lopuksi lähetään sähköpostia kahdelle vastaanottajalle.
# Sähköpostissa RPA-ajon aihe/nimi, ajohetki, kesto ja käsitelty rivilukumäärä.

harjoitus_name = "p2-h3"

# Change these email addresses to valid ones before running the script
email_addresses = "john.doe@gmail.com; john.doe@student.careeria.fi"

input_file = r"D:\\CAREERIA\\it-opiskelu\\2025-syksy\\robot\\harjoitus_data\\henkilot.xlsx"
data_columns = "A:C"
data_rows_after_header = 10 

program_file = r"D:\\CAREERIA\\it-opiskelu\\2025-syksy\\robot\\win_progs\\TietojenSyottoWPFOhjelma\\TietojenSyottoWPF.exe"

# get path to running directory
current_dir = os.path.dirname(os.path.abspath(__file__))

log_simple_filename = harjoitus_name + "_log_simple.txt"

program_start_time = datetime.datetime.now()

# Set up logging file to log all ui events with timestamps
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
    write_to_simple_log(s)


def write_to_simple_log(s):
    logger = get_simple_logger()
    logger.info(s)


def rpa_data_entry_main():
    write_to_log(f"Aloitusaika: {program_start_time.strftime('%d.%m.%Y %H:%M:%S')}")

    win = Windows()

    # Reduce default wait times and timeouts to make the script run faster
    previous = win.set_wait_time(0.1)
    print(f"Previous default wait time after keyboard or mouse actions: {previous} -> (now 0.1s)")

    previous = win.set_global_timeout(5)
    print(f"Previous global timeout for element search: {previous} -> (now 5s)")

    # Launch the application
    # If the app is already running the following line with windows_run could be commented out
    win.windows_run(program_file)
    write_to_log("Sovellus avattu.")

    try:
        w = win.control_window('regex:"MainWindow*"')
        write_to_log(f"Window handle: {w}")

        # get the UI elements once and reuse them to make the script faster
        element_nimi = win.get_element('automationid:"txtNimi"')
        element_osoite = win.get_element('automationid:"txtOsoite"')
        element_puhelin = win.get_element('automationid:"txtPuhelin"')
        element_lisaa = win.get_element('automationid:"btnLisaa"')
        # element_listan_luku = win.get_element('automationid:"btnListanLuku"')
        # element_oliolista = win.get_element('automationid:"btnOlioLista"')

        # Read target temperatures from the Excel file and set them one by one
        df = pd.read_excel(input_file, 
                           sheet_name='Sheet1', 
                           header=0, 
                           usecols=data_columns, 
                           nrows=data_rows_after_header)
        
        headers = df.columns.tolist()
        write_to_log(f"Excel headers: {headers}")

        for index, row in df.iterrows():
            print(f"Row {index + 1}: {row['Nimi']}, {row['Osoite']}, {row['Puhelinnumero']}")
            if isinstance(row['Nimi'], str): # Check the first column is a string (not NaN)
                if row['Nimi'].strip() != '':  # Check the first column is not empty string
                    win.set_value(element_nimi, row['Nimi'])
                    win.set_value(element_osoite, row['Osoite'])
                    win.set_value(element_puhelin, row['Puhelinnumero'])

                    # Click the set button to set the temperature
                    win.click(element_lisaa)
                    write_to_log(f"Lisätty: Excel rivi {index + 1}: {row['Nimi']} {row['Osoite']} {row['Puhelinnumero']}")
                else:
                    write_to_log(f"Epäkelpo rivi - ohitetaan: Excel rivi {index + 1}: {row['Nimi']} {row['Osoite']} {row['Puhelinnumero']} Syy: Nimi solussa välilyöntejä.")
                    print(f"Epäkelpo rivi - ohitetaan: Excel rivi {index + 1}: {row['Nimi']} {row['Osoite']} {row['Puhelinnumero']} Syy: Nimi solussa välilyöntejä.")
            else:
                write_to_log(f"Epäkelpo rivi - ohitetaan: Excel rivi {index + 1}: {row['Nimi']} {row['Osoite']} {row['Puhelinnumero']} Syy: Nimi solu on tyhjä (nan).")
                print(f"Epäkelpo rivi - ohitetaan: Excel rivi {index + 1}: {row['Nimi']} {row['Osoite']} {row['Puhelinnumero']} Syy: Nimi solu on tyhjä (nan).")

    except Exception as e:
        write_to_log(f"Virhe: {e}")

    finally:
        win.close_current_window()
        write_to_log("Sovellus suljettu.")
        program_end_time = datetime.datetime.now()
        program_duration = program_end_time - program_start_time
        
        write_to_log(f"Lopetusaika: {program_end_time.strftime('%d.%m.%Y %H:%M:%S')}")
        write_to_log(f"Kesto: {str(program_duration)}")
        write_to_log(f"Käsiteltyjen rivien määrä: {len(df)}")

        # Send email that the run is done with summary information and log file as attachment
        
        email_subject = f"RPA henkilötietojen tallennus {harjoitus_name} ajettu"
        email_body = f'''RPA henkilötietojen tallennus {harjoitus_name} on ajettu.\n\n'''
        email_body += f'''Ajohetki: {program_start_time.strftime('%d.%m.%Y %H:%M:%S')}\n'''
        email_body += f'''Kesto: {str(program_duration)}\n'''
        email_body += f'''Käsiteltyjen rivien määrä: {len(df)}\n\n'''
        email_body += f'''LIITTEET: {log_simple_filename}\n'''

        attachment_file = os.path.join(current_dir, log_simple_filename)

        try:
            outlook_app = Application()
            outlook_app.open_application()

            # From RPA.Outlook.Application documentation:
            # https://rpaframework.org/libdoc/RPA_Outlook_Application.html
            # Send email with Outlook parameters:
            # recipients:	    list of addresses, ex. 'EMAILADDRESS_1; EMAILADDRESS_2'
            #                   a string with ; separated email addresses NOT , separated as in documentation example
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

            return_value =outlook_app.send_email(
                recipients=email_addresses,
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

        except Exception as e:
            write_to_log(f"Email sending failed: {e}")
            print(f"Email sending failed: {e}")

        write_to_log("****** DONE ******")
        print("****** DONE ******")

if __name__ == "__main__":
    rpa_data_entry_main()
