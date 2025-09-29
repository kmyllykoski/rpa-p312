from RPA.Browser.Selenium import Selenium

def main():
    print("Hello from rpa-uv-rpaframework!")
    browser = Selenium()
    browser.open_available_browser("https://google.com")
    # waiting for 5 seconds to see the opened browser
    browser.wait_until_page_contains("Googlesomething", timeout=5)
    browser.close_browser()


if __name__ == "__main__":
    main()

