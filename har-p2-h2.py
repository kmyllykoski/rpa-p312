from RPA.Windows import Windows
import logging
import csv
from time import sleep

# Päivä 2
# Harjoitus 2
#
# Luetaan lämpötiloja tiedostosta ja asetetaan uusia lämpotiloja.
# Aloitus, virhetilanteet ja lopetus kirjataan lokiin.
# Tarkistetaan että lämpötila on asetuksen jälkeen oikein. 
# Eli jos asetetaan 22, luetaan jälkeenpäin että lämpötila on 22.
# Ellei ole, kirjataan virhe lokiin.

harjoitus_name = "p2-h2"

input_file = r"D:\\CAREERIA\\it-opiskelu\\2025-syksy\\robot\\harjoitus_data\\tavoitelampoja.csv"
program_file = r"D:\\CAREERIA\\it-opiskelu\\2025-syksy\\robot\\win_progs\\WpfAlyTalo\\WpfAlyTalo.exe"


# Set up logging file to log all events with timestamps
log_filename = harjoitus_name + "_log.txt"
logging.basicConfig(
    filename=log_filename,
    level=logging.INFO,
    format="%(asctime)s %(message)s",
    encoding="utf-8",
    filemode="w",
)  # Overwrite log file each run


# Second simpler logger that uses diffrent logging file than general logging
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
        w = win.control_window('regex:"Älytalosovellus*"')
        write_to_log(f"Window handle: {w}")

        # get the UI elements once and reuse them to make the script faster
        element_current_temperature = win.get_element('automationid:"txtTalonLampotila"')
        element_new_temperature = win.get_element('automationid:"txtUusiLampotila"')
        element_set_button = win.get_element('automationid:"bAsetaLampotila"')

        # Read the current temperature value and log it
        lampotila = win.get_value(element_current_temperature)
        write_to_log(f"Talon lämpotila on nyt: {lampotila}")  # write_to_log will call also write_to_simple_log

        # Read target temperatures from the input file and set them one by one
        with open(input_file, newline="", encoding="utf-8") as csvfile:
            reader = csv.DictReader(csvfile, delimiter=";")
            for row in reader:
                if row:  # Check if the row is not empty
                    target_temperature = row["TavoiteLampo"]

                    if int(target_temperature) < 10:
                        write_to_log(f"Virheellinen lämpotila: {target_temperature} -> Ohitetaan.")
                        continue
                    if int(target_temperature) > 100:
                        write_to_log(f"Virheellinen lämpotila: {target_temperature} -> Ohitetaan.")
                        continue

                    # Set value to the new temperature field
                    win.set_value(element_new_temperature, target_temperature)
                    # Click the set button to set the temperature
                    win.click(element_set_button)
                    
                    lampotila = win.get_value(element_current_temperature)
                    write_to_log(f"Talon lämpotila on nyt: {lampotila}")

                    if lampotila[:-1] != target_temperature:        # lampotila is with degree sign e.g. "22°" so we ignore last character
                        write_to_log(f"Virhe: Asetettu lämpotila {target_temperature}, mutta talon lämpotila on {lampotila[:-1]}")

    except Exception as e:
        write_to_log(f"Virhe: {e}")

    finally:
        sleep(3)  # wait a bit before closing the app
        win.close_current_window()
        write_to_log("Sovellus suljettu.")


if __name__ == "__main__":
    harjoitus()
