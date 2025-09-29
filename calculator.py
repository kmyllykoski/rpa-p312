# write a script that uses rpaframerwork that opens calculator app and performs addition operation
# from robocorp.tasks import task
from RPA.Windows import Windows
from robot.api import logger

# windows = Windows()

# @task
def minimal_task():
    message = "Hello"
    message = message + " World!"
    write_to_console(message)
    # windows.windows_run("calc.exe")
    win = Windows()

    previous = win.set_mouse_movement(True)
    print(f"Previous mouse simulation: {previous} (now enabled)")

    win.windows_run("calc.exe")
    try:
        win.control_window("name:Calculator")
        # win.print_tree(max_depth=8, show_properties=True)
        win.click("automationid:clearButton")
        win.click("automationid:num2Button")
        win.click("automationid:plusButton")
        win.click("automationid:num3Button")
        win.click("automationid:equalButton")

        # You can also use the following selectors
        # win.click('name:"Clear"')
        # win.click('id:clearButton') 
        # win.click('name:"Two"')
        # win.click('name:"Plus"')
        # win.click('name:"Three"')
        # win.click('name:"Equals"')

        # Alternatively, you can use send_keys to perform the operation
        # win.send_keys(keys="96+4=")

        result = win.get_attribute("id:CalculatorResults", "Name")
        print(result)
        buttons = win.get_elements(
            'type:Group and name:"Number pad" > type:Button'
        )
        for button in buttons:
            print(button)
    finally:
        win.close_current_window()

    # app = windows.wait_for_element('regex:.*Calculator')
    # app('class:"Button" and name:"Clear"').click()
    # app('class:"Button" and name:"Two"').click()
    # app('class:"Button" and name:"Plus"').click()
    # app('class:"Button" and name:"Three"').click()
    # app('class:"Button" and name:"Equals"').click()

def write_to_console(s):
    logger.console(s)

if __name__ == "__main__":
    minimal_task()