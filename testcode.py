from RPA.Windows import Windows
import logging

log_filename = "harjoitus_log.txt"
logging.basicConfig(filename=log_filename, level=logging.INFO, format="%(asctime)s %(message)s")

def write_to_log(s):
    logging.info(s)

def harjoitus():
    win = Windows()

    previous = win.set_mouse_movement(True)
    print(f"Previous mouse simulation: {previous} (now enabled)")

    previous = win.set_wait_time(0.1)
    print(f"Previous default wait time after keyboard or mouse actions: {previous} -> (now 0.1s)")

    previous = win.set_global_timeout(5)
    print(f"Previous global timeout for element search: {previous} -> (now 5s)")

    win.windows_run("D:\\CAREERIA\\it-opiskelu\\2025-syksy\\robot\\win_progs\\WpfAlyTalo\\WpfAlyTalo.exe")

    try:
        win.control_window('regex:"Älytalosovellus*"')

        # 1st way to change the temperature
        # - mouse click to the text field
        win.click('path:18 and automationid:"txtUusiLampotila"')
        # - send keys to the text field
        win.send_keys('path:18 and automationid:"txtUusiLampotila"', "22")
        
        # 2nd way to change the temperature
        # - get the input field element and set value directly
        win.set_value('path:18 and automationid:"txtUusiLampotila"', "18")

        # Here the same thing but with first clearing the field
        win.send_keys('path:18 and automationid:"txtUusiLampotila"', "{Ctrl}a{Del}")
        win.set_value('path:18 and automationid:"txtUusiLampotila"', "20")
        
        # Finally click the button to set the temperature
        win.click('path:19 and automationid:"bAsetaLampotila"')

        # Read the current temperature value and log it
        lampotila = win.get_value('path:17 and automationid:"txtTalonLampotila"')
        write_to_log(f"Talon lämpotila on nyt: {lampotila}")

    finally:
        win.close_current_window()
        write_to_log("Sovellus suljettu.")

if __name__ == "__main__":
    harjoitus()