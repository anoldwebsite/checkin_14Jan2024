import time
import tkinter as tk
from selenium import webdriver
from selenium.common import TimeoutException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.wait import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
from excel_maker import make_sn_excel_file

options = Options()
options.add_argument("start-maximized")
options.add_experimental_option("detach", True)
service = Service(ChromeDriverManager().install())


# driver = webdriver.Chrome(options=options, service=service)

# driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager(version='114.0.5735.90').install()))
# driver = webdriver.Chrome(options=options, service=Service(ChromeDriverManager().install()))
def fill_form(url, serialNumList):
    # Start a new instance of Chrome
    # driver = webdriver.Chrome(options=options, service=Service(ChromeDriverManager(version='114.0.5735.90').install()))
    driver = webdriver.Chrome(options=options, service=service)
    # Load the URL
    driver.get(url)
    time.sleep(5)
    # Close the "Accept cookies" pop up box.
    try:
        cookies_acceptance = driver.find_element('xpath', '/html/body/div[3]/div[2]/div/div/div[2]/div/div/button[1]')
        cookies_acceptance.click()
        give_stars = driver.find_element('xpath',
                                         '/html/body/div/div/div/div[2]/neb-form-footer/div/div[1]/div/div/div/div/button[1]')
        give_stars.click()
    except:
        pass
    # Check the length of serialNumList and add new input fields as needed
    # 2 input boxes are already on the page. Subtract 2.
    required_input_boxes = len(serialNumList) - 2
    present_input_boxes = 2
    for i in range(required_input_boxes):
        add_serial_btn = driver.find_element('xpath',
                                             '/html/body/app-root/div/app-layout/app-check-warranty/div/div/div[2]/app-check-warranty-landing/div[2]/app-multiple-product/div/form/div[2]/span[1]')
        add_serial_btn.click()
        present_input_boxes += 1
    # Fill in the serial number input fields
    serial_num_inputs = []
    # Add new boxes to input serial numbers of hp laptops in.
    for i in range(present_input_boxes):
        serial_input_full_xpath = '/html/body/app-root/div/app-layout/app-check-warranty/div/div/div[2]/app-check-warranty-landing/div[2]/app-multiple-product/div/form/div[1]/'
        serial_input_full_xpath += 'div[' + str(i + 1) + ']'  # list starts from zero but our divs starts from 1.
        serial_input_full_xpath += '/div/div/div[1]/input'
        serial_ith = driver.find_element('xpath', serial_input_full_xpath)
        serial_num_inputs.append(serial_ith)
    # Fill the input boxes with serial numbers.
    for i in range(len(serialNumList)):
        serial_num_inputs[i].send_keys(serialNumList[i])

    # Submit the form
    submit_btn = driver.find_element('xpath',
                                     '/html/body/app-root/div/app-layout/app-check-warranty/div/div/div[2]/app-check-warranty-landing/div[2]/app-multiple-product/div/form/div[3]/button')
    # save current page url
    current_url = driver.current_url
    # print(current_url)
    # initiate page transition, e.g.:
    time.sleep(1)
    submit_btn.click()
    # Wait for the new page to load
    try:
        # wait for URL to change with 59 seconds timeout
        WebDriverWait(driver, 59).until(EC.url_changes(current_url))
    except TimeoutException:
        print(f"One of these serial numbers might have caused problems while scrapping data from HP website: {serialNumList}")
        # driver.quit()
        root = tk.Tk()
        root.withdraw()
        tk.messagebox.showerror(title="Error",
                                message="HP website did not return the required data! Please check manually if the HP website can return data using the input serial numbers of assets.")
        return []
    # print new URL
    # next_url = driver.current_url
    # print(next_url)

    # Call the get_soup() function to get the parsed HTML
    soup = get_soup(driver)

    # Call the scrape_laptops() function to extract the laptop information
    laptops = scrape_laptops(soup)

    # Close the browser window
    driver.quit()

    return laptops


# https://www.browserstack.com/guide/web-scraping-using-selenium-python
# https://tutorialscamp.com/beautifulsoup-how-to-get-nested-inner-divs/
def get_soup(d):
    # Scrape the results
    page_source = d.page_source
    s = BeautifulSoup(page_source, features='html.parser')
    return s

    # https://realpython.com/beautiful-soup-web-scraper-python/


def scrape_laptops(sauce):
    products = []
    divs = sauce.find_all("div", {"class": "product-info"})
    for laptop in divs:
        p = {
            'serial': laptop.find("div", {"class": "serial-no"}).find("span").text,
            'pid': laptop.find("div", {"class": "product-no"}).find("span").text,
            'model': laptop.find("h4").text
        }
        products.append(p)
    return products


def get_data_from_hp(serial_num_list):
    url = "https://support.hp.com/se-sv/check-warranty#multiple"
    # new_url = fill_form(url, serial_num_list)
    # soup = get_soup(driver)
    # serial_pid_title_list = scrape_laptops(soup)
    # for element in serial_pid_title_list:
    # print(element)
    serial_pid_title_list = fill_form(url, serial_num_list)
    return serial_pid_title_list