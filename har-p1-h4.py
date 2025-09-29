from RPA.Windows import Windows
import logging
import csv
from time import sleep

# Harjoitus 2
# Luetaan henkilötietoja tiedostosta ja tallennetaan niitä käyttöliittymään

harjoitus_name = "p1-h4"

input_file = r"D:\\CAREERIA\\it-opiskelu\\2025-syksy\\robot\\harjoitus_data\\henkilot.csv"
program_file = r"D:\\CAREERIA\\it-opiskelu\\2025-syksy\\robot\\win_progs\\TietojenSyottoWPFOhjelma\\TietojenSyottoWPF.exe"

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
    logger = logging.getLogger("harjoitus_logger")
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


def harjoitus():
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

        # Read target temperatures from the input file and set them one by one.
        # Excel on Windows usually exports CSV files in ANSI/Windows-1252 (not UTF-8) encoding by default.
        # That's why we need to use encoding="latin-1" here to read scandinavian characters correctly.
        with open(input_file, encoding="latin-1", newline="") as csvfile:
            reader = csv.DictReader(csvfile, delimiter=";")
            for row in reader:
                if row:  # Check if the row is not empty          
                    win.set_value(element_nimi, row['Nimi'])
                    win.set_value(element_osoite, row['Osoite'])
                    win.set_value(element_puhelin, row['Puhelinnumero'])

                    # Click the set button to set the temperature
                    win.click(element_lisaa)
                    write_to_log(f"Lisätty: {row['Nimi']} {row['Osoite']} {row['Puhelinnumero']}")
        
        # Click the button that reads the list and shows it in the text area
        win.click(element_oliolista)
        write_to_log("Luettu lista.")

    except Exception as e:
        write_to_log(f"Virhe: {e}")

    finally:
        sleep(5)  # wait a bit before closing the app
        win.close_current_window()
        write_to_log("Sovellus suljettu.")


if __name__ == "__main__":
    harjoitus()
