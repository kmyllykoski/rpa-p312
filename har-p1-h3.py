from RPA.Windows import Windows
import logging
import csv
from time import sleep

# Päivä 1
# Harjoitus 3
#
# Luetaan lämpötiloja tiedostosta ja asetetaan uusia lämpotiloja.
# Klikataan lopuksi kaikkia muita nappeja käyttöliittymässä paitsi asetettu lämpötila -nappia.

harjoitus_name = "p1-h3"

input_file = r"D:\\CAREERIA\\it-opiskelu\\2025-syksy\\robot\\harjoitus_data\\tavoitelampoja.csv"
program_file = r"D:\\CAREERIA\\it-opiskelu\\2025-syksy\\robot\\win_progs\\WpfAlyTalo\\WpfAlyTalo.exe"


# Set up logging file to log all ui events with timestamps
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
                    write_to_log(f"Asetettu uusi lämpotila: {target_temperature}")

        # Read the current temperature value and log it
        lampotila = win.get_value(element_current_temperature)
        write_to_log(f"Talon lämpotila on nyt: {lampotila}")
        
        # get all buttons in the window 
        buttons_in_ui = []
        all_buttons = win.get_elements('path:2 and control:"ButtonControl"')
        write_to_log(f"Ui buttons found: {len(all_buttons)}")
        for button in all_buttons:
            write_to_log(f"Button found: {button}")
            buttons_in_ui.append(button)

        # Click all buttons found in the window except the set button for room temperature
        for button in buttons_in_ui:
            # if the automation id is not the same as the automation id of the set button, click it
            if str(getattr(button.item, "AutomationId")) != str(getattr(element_set_button.item, "AutomationId")):
                write_to_log(f"Clicking button: {button}")
                win.click(button)
            else:
                write_to_log(f"Not clicking button (is the set button): {button}")
        
    except Exception as e:
        write_to_log(f"Virhe: {e}")

    finally:
        sleep(5)  # wait a bit before closing the app
        win.close_current_window()
        write_to_log("Sovellus suljettu.")


if __name__ == "__main__":
    harjoitus()
