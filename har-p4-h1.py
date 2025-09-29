from RPA.Windows import Windows
# from RPA.Outlook.Application import Application
import logging
import pandas as pd
import datetime
from os import getenv
from dotenv import load_dotenv    # uv add dotenv
from mssql_python import connect  # uv add mssql-python
from time import sleep

# Päivä 4
# Harjoitus 1
#
# RPA-sovellus, joka lukee Northwind-tietokannasta kaikki asiakkaat (Customers-taulu) ja vie ne TietojenSyottoWPF.exe -ohjelmaan
#
# Nimi: CustomerID + ContactName 
# Osoite: Address + City + PostalCode + Country 
# Puhelinnumero: “Puh: “ + Phone + “ Fax: “ + Fax
#
# Lisätty virheenkäsittely try/except/finally lohkoilla tietokantayhteyden muodostukseen, datan SQL-hakuun
# sekä tietojen tallennukseen käyttöliittymässä. Virheistä tallennetaan lokiin kuvaus.

harjoitus_name = "p4-h1"

program_file = r"D:\\CAREERIA\\it-opiskelu\\2025-syksy\\robot\\win_progs\\TietojenSyottoWPFOhjelma\\TietojenSyottoWPF.exe"

program_start_time = datetime.datetime.now()


def connect_to_db():
    load_dotenv()
    try:
        conn = connect(getenv("SQL_CONNECTION_STRING"))  # get connection string from .env file
        # Create a file .env in the same folder as this script and add the connection string there
        # Example of .env file contents with local MSSQL Express database on Windows:
        # SQL_CONNECTION_STRING="Server=(LocalDb)\MSSQLLocalDB;Database=Northwind;Trusted_Connection=yes;Encrypt=no;TrustServerCertificate=yes;"
        # 
        # Script to create Northwind database in local MSSQL Express:
        # https://github.com/microsoft/sql-server-samples/blob/master/samples/databases/northwind-pubs/instnwnd.sql
        return conn
    except Exception as e:
        write_to_log(f"Virhe tietokantayhteydessä: {e}")
        return None

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
        handler = logging.FileHandler(harjoitus_name + "_log_simple.txt", mode="w", encoding="utf-8")
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
        element_oliolista = win.get_element('automationid:"btnOlioLista"')

        # Read customer data from the database. Limit to first 5 rows for testing
        try:
            sql_query = "SELECT TOP 5 CustomerID, ContactName, Address, City, PostalCode, Country, Phone, Fax FROM Customers"
            df = pd.read_sql(sql_query, conn)
            write_to_log(f"Asiakastiedot luettu tietokannasta Customers-taulusta. Rivien määrä: {len(df)}")
            print(f"Asiakastiedot luettu tietokannasta Customers-taulusta. Rivien määrä: {len(df)}")
        except Exception as e:
            write_to_log(f"Virhe SQL-kyselyssä tai datan lukemisessa: {e}")
            print(f"Virhe SQL-kyselyssä tai datan lukemisessa: {e}")
            return

        for index, row in df.iterrows():
            print("-" * 40)
            print(f"Processing row {index + 1} / {len(df)}: {row['CustomerID']}, {row['ContactName']}, {row['Address']}, {row['City']}, {row['PostalCode']}, {row['Country']}, {row['Phone']}, {row['Fax']}")
            nimi = f"{row['CustomerID']} {row['ContactName']}"
            osoite = f"{row['Address']}, {row['City']}, {row['PostalCode']}, {row['Country']}"
            puhelin = f"Puh: {row['Phone']} Fax: {row['Fax']}"
            print(f"Tallentuva nimi: {nimi}")
            print(f"Tallentuva osoite: {osoite}")
            print(f"Tallentuva puhelin: {puhelin}")
            
            try:
                win.set_value(element_nimi, nimi)
                win.set_value(element_osoite, osoite)
                win.set_value(element_puhelin, puhelin)
                # Click the Add button to save the information
                win.click(element_lisaa)
                win.click(element_oliolista)  # Click the "Oliolista" button to refresh the list
                write_to_log(f"Lisätty: Excel rivi {index + 1}: {nimi} {osoite} {puhelin}")
                
                # Click the button that reads the list and shows it in the text area
                win.click(element_oliolista)
            
            except Exception as e:
                write_to_log(f"Virhe tietojen tallennuksessa: {e}")

    except Exception as e:
        write_to_log(f"Virhe: {e}")

    finally:
        program_end_time = datetime.datetime.now()
        program_duration = program_end_time - program_start_time
        
        write_to_log(f"Lopetusaika: {program_end_time.strftime('%d.%m.%Y %H:%M:%S')}")
        write_to_log(f"Kesto: {str(program_duration)}")
        write_to_log(f"Käsiteltyjen rivien määrä: {len(df)}")

        sleep(5)  # Wait for 5 seconds before closing the window
        win.close_current_window()
        write_to_log("Sovellus suljettu.")

        write_to_log("****** DONE ******")
        print("****** DONE ******")

if __name__ == "__main__":
    conn = connect_to_db()
    if conn is not None:
        cursor = conn.cursor()
        rpa_data_entry_main()
