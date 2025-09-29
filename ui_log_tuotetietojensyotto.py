from RPA.Windows import Windows
import logging

# Avataan Windows ohjelma ja tulostetaan lokiin käyttöliittymän rakenne
# ja konsoliin haluttujen kontrollien tiedot

controls_to_get = ["EditControl", "ButtonControl", "CheckBoxControl"]
skip_controls_with_names = ["Minimize", "Maximize", "Close"]

ohjelma = "TuoteTietojenSyottoWPF"

program_file = r"D:\\CAREERIA\\it-opiskelu\\2025-syksy\\robot\\win_progs\\TuoteTietojenSyottoWPF\\TuoteTietojenSyottoWPF.exe"

# Set up logging file to log all ui events with timestamps
log_filename = "UI_" + ohjelma + "_log.txt"
logging.basicConfig(
    filename=log_filename,
    level=logging.INFO,
    format="%(asctime)s %(message)s",
    encoding="utf-8",
    filemode="w",
)  # Overwrite log file each run


def write_to_log(s):
    logging.info(s)


def get_ui_info():
    win = Windows()

    # Create a list to hold found AutomationIds of controls of interest
    found_automation_ids = []

    # Launch the application
    # If the app is already running the following line with windows_run could be commented out
    win.windows_run(program_file)
    write_to_log("Sovellus avattu.")

    try:
        w = win.control_window('regex:"Tuotetietojen*"')
        write_to_log(f"Window handle: {w}")

        # print_tree returns the UI structure and logs it in the main log file
        ui_components = win.print_tree(return_structure=True)

        for control in controls_to_get:
            print("-" * 40)
            print(f"\n\nListing all {control} controls in the window:\n")
            for group, elements in ui_components.items():
                print(f"Group {group}:")
                for el in elements:
                    if (
                        el.control_type == control
                        and el.name not in skip_controls_with_names
                    ):
                        print(f"  ControlType: {el.control_type}")
                        print(f"  Name: {el.name}")
                        print(f"  AutomationId: {el.automation_id}")
                        print(f"  ClassName: {el.class_name}")
                        print(f"  Locator: {el.locator}")
                        print(f"  Rect: ({el.left}, {el.top}, {el.right}, {el.bottom})")
                        print("-" * 40)
                        found_automation_ids.append(el.automation_id)

    except Exception as e:
        write_to_log(f"Virhe: {e}")

    finally:
        win.close_current_window()
        write_to_log("Sovellus suljettu.")
        print("")
        print("=" * 40)
        print("Found AutomationId's for controls of interest:")
        for aid in found_automation_ids:
            print(aid)


if __name__ == "__main__":
    get_ui_info()
