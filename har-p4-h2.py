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
# Harjoitus 2
#
# RPA-sovellus, joka lukee Northwind-tietokannan products tiedot ja vie ne TuoteTietojenSyottoWPF.exe -ohjelmaan
#
# products tauluun lisätään uusi sarake "RPAProcessed" (varchar(10)) johon tallennetaan 'X' jos rivi on käsitelty RPA-sovelluksella.
# Jos tuotetietojen käsittelyssä tapahtuu virhe, RPAProcessed sarakkeeseen tallennetaan 'Error' ja virhe tallennetaan lokiin.
# 'Error' merkittyjä tietueita ei käsitellä seuraavalla ajokerralla, vain ne rivit joiden RPAProcessed on NULL käsitellään.


harjoitus_name = "p4-h2"

program_file = r"D:\\CAREERIA\\it-opiskelu\\2025-syksy\\robot\\win_progs\\TuoteTietojenSyottoWPF\\TuoteTietojenSyottoWPF.exe"

program_start_time = datetime.datetime.now()


def connect_to_db():
    try:
        load_dotenv()
        conn = connect(getenv("SQL_CONNECTION_STRING"))  # get connection string from .env file
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


def check_rpaprocessed_column_exists():
    try:
        # cursor = conn.cursor()
        cursor.execute("SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'Products' AND COLUMN_NAME = 'RPAProcessed'")
        result = cursor.fetchone()
        if result:
            write_to_log("RPAProcessed sarake löytyy Products taulusta.")
            return True
        else:
            write_to_log("RPAProcessed saraketta ei löydy Products taulusta. Lisätään sarake.")
            cursor.execute("ALTER TABLE Products ADD RPAProcessed VARCHAR(10) NULL")
            conn.commit()
            write_to_log("RPAProcessed sarake lisätty Products tauluun.")
            return True
    except Exception as e:
        write_to_log(f"Virhe tarkistettaessa tai lisättäessä RPAProcessed saraketta: {e}")
        return False


def update_rpaprocessed(product_id, status):
    try:
        cursor.execute('''
                       UPDATE Products 
                       SET RPAProcessed = ? 
                       WHERE ProductID = ?''', (status, product_id))
        
        conn.commit()
        write_to_log(f"Updated ProductID {product_id} with RPAProcessed = {status}")
        return True
    except Exception as e:
        write_to_log(f"Virhe päivitettäessä RPAProcessed saraketta ProductID {product_id}: {e}")
        return False


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
        w = win.control_window('regex:"Tuotetietojen*"')
        write_to_log(f"Window handle: {w}")

        # get the UI elements once and reuse them to make the script faster
        element_tuote_id = win.get_element('automationid:"txtTuoteID"')
        element_nimi = win.get_element('automationid:"txtTuoteNimi"')
        element_pakkaus_maara = win.get_element('automationid:"txtPakkausMaara"')
        element_yksikko_hinta = win.get_element('automationid:"txtYksikkoHinta"')
        element_varasto_maara = win.get_element('automationid:"txtVarastoMaara"')
        element_varaston_arvo = win.get_element('automationid:"txtVarastonArvo"')
        element_tuote_on_voimassa = win.get_element('automationid:"cbVoimassa"')

        element_btnLisaaTuote = win.get_element('automationid:"btnLisaaTuote"')
        element_btnVieTaulukkoon = win.get_element('automationid:"btnVieTaulukkoon"')


        # Read products data from the database. Limit to first 5 rows for testing
        try:
            sql_query = '''
                SELECT TOP 5 ProductID, ProductName, QuantityPerUnit, UnitPrice, UnitsInStock, Discontinued 
                FROM Products
                WHERE RPAProcessed IS NULL '''
            
            df = pd.read_sql(sql_query, conn)
            write_to_log(f"Tuotetiedot luettu tietokannasta Products-taulusta. Rivien määrä: {len(df)}")
            print(f"Tuotetiedot luettu tietokannasta Products-taulusta. Rivien määrä: {len(df)}")
        except Exception as e:
            write_to_log(f"Virhe SQL-kyselyssä tai datan lukemisessa: {e}")
            print(f"Virhe SQL-kyselyssä tai datan lukemisessa: {e}")
            df = pd.DataFrame()  # Create an empty DataFrame to avoid further errors

        for index, row in df.iterrows():
            print("-" * 40)
            print(f"Processing row {index + 1} / {len(df)}: {row['ProductID']}, {row['ProductName']}, {row['QuantityPerUnit']}, {row['UnitPrice']}, {row['UnitsInStock']}, {row['Discontinued']}")
            
            try:
                win.set_value(element_tuote_id, row['ProductID'])
                win.set_value(element_nimi, row['ProductName'])
                win.set_value(element_pakkaus_maara, row['QuantityPerUnit'])
                win.set_value(element_yksikko_hinta, row['UnitPrice'])
                win.set_value(element_varasto_maara, row['UnitsInStock'])
                stock_value = row['UnitsInStock'] * row['UnitPrice']
                win.set_value(element_varaston_arvo, stock_value)

                # Tuote on voimassa (Discontinued = 0) tai ei ole voimassa (Discontinued = 1)
                if row['Discontinued'] == 0:
                    write_to_log("Tuote on voimassa, CheckBox tulee olla valittuna. Valitaan se.")
                    # win.click(element_tuote_on_voimassa)  # click also works here
                    win.set_focus(element_tuote_on_voimassa)
                    win.send_keys(element_tuote_on_voimassa, "{SPACE}")
                else:
                    write_to_log("Tuote ei ole voimassa ja CheckBox jätetään ennalleen tilaan ei valittuna.")
                # There is currently no easy way to check the state of CheckBox (selected or not)
                # as there is no keyword for example component.is_checked() in RPA Framework Windows library.
                # So we assume that the CheckBox is initially not selected and we select it only if Discontinued = 0
                # See https://github.com/robocorp/rpaframework/issues/1166
                # Maybe it would be possible to use the screenshot function and compare images to check the state of the CheckBox.
                # This is outside the scope of this exercise.
                    
                # Click the Add button to save the information
                win.click(element_btnLisaaTuote)
                win.click(element_btnVieTaulukkoon)
                write_to_log(f"Lisätty tuote {index + 1} / {len(df)}: {row['ProductID']} {row['ProductName']} {row['QuantityPerUnit']} {row['UnitPrice']} {row['UnitsInStock']} {row['Discontinued']}")

                if not update_rpaprocessed(row['ProductID'], 'X'):
                    write_to_log(f"Virhe päivitettäessä RPAProcessed saraketta ProductID {row['ProductID']}.")
                    return  # Stop processing further rows if update fails

            except Exception as e:
                write_to_log(f"Virhe tietojen tallennuksessa: {e}")
                if not update_rpaprocessed(row['ProductID'], 'Error'):
                    write_to_log(f"Virhe päivitettäessä RPAProcessed saraketta ProductID {row['ProductID']} virhetilanteessa.")
                    return  # Stop processing further rows if update fails

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
        
if __name__ == "__main__":
    conn = connect_to_db()
    cursor = conn.cursor()
    
    if check_rpaprocessed_column_exists():
        rpa_data_entry_main()
    else:
        write_to_log("RPAProcessed saraketta ei voida tarkistaa tai lisätä. Ohjelma lopetetaan.")
        print("RPAProcessed saraketta ei voida tarkistaa tai lisätä. Ohjelma lopetetaan.")
    
    conn.close()    
    write_to_log("****** DONE ******")
    print("****** DONE ******")


