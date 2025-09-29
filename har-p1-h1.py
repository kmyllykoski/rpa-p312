from RPA.Windows import Windows
import logging
from time import sleep

# Päivä 1
# Harjoitus 1
# Luetaan lämpötiloja tiedostosta ja asetetaan uusia lämpotiloja

harjoitus_name = "p1-h1"

log_simple_filename = harjoitus_name + "_log_simple.txt"

# Set up logging file to log all RPA events with timestamps
log_filename = harjoitus_name + "_log.txt"
logging.basicConfig(
    filename=log_filename,
    level=logging.INFO,
    format="%(asctime)s %(message)s",
    encoding="utf-8",
    filemode="w",
)  # Overwrites log file each run


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


def harjoitus():
    win = Windows()

    # Reduce default wait times and timeouts to make the script run faster
    previous = win.set_wait_time(0.1)
    print(f"Previous default wait time after keyboard or mouse actions: {previous} -> (now 0.1s)")

    previous = win.set_global_timeout(5)
    print(f"Previous global timeout for element search: {previous} -> (now 5s)")

    # Launch the application
    # If the app is already running the following line with windows_run could be commented out
    win.windows_run("D:\\CAREERIA\\it-opiskelu\\2025-syksy\\robot\\win_progs\\WpfAlyTalo\\WpfAlyTalo.exe")
    write_to_log("Sovellus avattu.")

    try:
        w = win.control_window('regex:"Älytalosovellus*"')
        write_to_log(f"Window handle: {w}")

        element_current_temperature = win.get_element('path:17 and automationid:"txtTalonLampotila"')
        element_new_temperature = win.get_element('path:18 and automationid:"txtUusiLampotila"')
        element_set_button = win.get_element('path:19 and automationid:"bAsetaLampotila"')
        
        # Read the current temperature value and log it
        lampotila = win.get_value(element_current_temperature)
        write_to_log(f"Talon lämpotila on nyt: {lampotila}")    # write_to_log will call also write_to_simple_log

        # 1st way to change the temperature
        # - send keys to the text field
        win.send_keys(element_new_temperature, "22")

        # 2nd way to change the temperature
        # - get the input field element and set value directly
        win.set_value(element_new_temperature, "18")

        # Here the same thing but with first clearing the field
        win.set_value(element_new_temperature, "")
        # this would also work: win.send_keys('path:18 and automationid:"txtUusiLampotila"', "{Ctrl}a{Del}")
        win.set_value(element_new_temperature, "20")
        
        # Finally click the set button to set the temperature
        win.click(element_set_button)

        # Read the current temperature value and log it
        lampotila = win.get_value(element_current_temperature)
        write_to_log(f"Talon lämpotila on nyt: {lampotila}")

    except Exception as e:
        write_to_log(f"Virhe: {e}")

    finally:
        sleep(3)  # Wait a bit before closing the app
        win.close_current_window()
        write_to_log("Sovellus suljettu.")

if __name__ == "__main__":
    harjoitus()
    