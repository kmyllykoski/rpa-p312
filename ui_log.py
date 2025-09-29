from RPA.Windows import Windows
import logging

# Avataan Windows ohjelma ja tulostetaan lokiin käyttöliittymän rakenne
# ja konsoliin haluttujen kontrollien tiedot

controls_to_get = ["EditControl", "ButtonControl"]
skip_controls_with_names = ["Minimize", "Maximize", "Close"]

ohjelma = "TietojenSyottoWPFOhjelma"

program_file = r"D:\\CAREERIA\\it-opiskelu\\2025-syksy\\robot\\win_progs\\TietojenSyottoWPFOhjelma\\TietojenSyottoWPF.exe"

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

    # Launch the application
    # If the app is already running the following line with windows_run could be commented out
    win.windows_run(program_file)
    write_to_log("Sovellus avattu.")

    try:
        w = win.control_window('regex:"MainWindow*"')
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

        # # get all edit controls in the window
        # print("Getting all edit controls in the window")
        # elements_in_ui = []
        # all = win.get_elements('path:2 and control:"EditControl"')
        # # all_buttons = win.get_elements('regex:"MainWindow*" and type:WindowControl')
        # write_to_log(f"Ui elements found: {len(all)}")
        # for e in all:
        #     write_to_log(f"Element found: {e}")
        #     elements_in_ui.append(e)
        #     print(f"Element found: {e}")
        #     print("-" * 40)

        # # get all button controls in the window
        # print("Getting all button controls in the window")
        # elements_in_ui = []
        # all = win.get_elements('path:1 and control:"ButtonControl"')
        # # all_buttons = win.get_elements('regex:"MainWindow*" and type:WindowControl')
        # write_to_log(f"Ui elements found: {len(all)}")
        # for e in all:
        #     write_to_log(f"Element found: {e}")
        #     elements_in_ui.append(e)
        #     print(f"Element found: {e}")
        #     print("-" * 40)

    except Exception as e:
        write_to_log(f"Virhe: {e}")

    finally:
        # win.close_current_window()
        write_to_log("Sovellus suljettu.")


if __name__ == "__main__":
    get_ui_info()
