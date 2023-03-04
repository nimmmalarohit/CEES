from datetime import datetime
from time import sleep
import json
from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By
import os

output_directory = r"D:\git\CEES\output"
# output_directory = r"C:\Users\Rohit Nimmala\Documents\cees\output"
default_file_name = "Microsoft Forms.pdf"
student_name_xpath = "/html/body/div[2]/div/div/div/div/div[3]/div[1]/div[2]/div[1]/div/div[2]/div/div"
date_collected_css_selector = "#DatePicker_r4f38b41a4059401ca80255720cf09922"
next_response_xpath = "/html/body/div[2]/div/div/div/div/div[3]/div[1]/div[1]/div[2]/div/button[2]"
number_of_results_xpath = "/html/body/div[2]/div/div/div/div[2]/div[1]/div[2]/div/div/div[1]/div[2]/div[2]/div/div[1]/div[1]"
login_button_xpath = "/html/body/section/div/div[2]/div/form/div[3]/div/button"
view_results_button_xpath = "/html/body/div[2]/div/div/div/div[2]/div[1]/div[2]/div/div/div[1]/div[3]/button/div"


def wait_for_an_element(identifier, find_by=By.ID):
    try:
        element = WebDriverWait(driver, 900).until(EC.visibility_of_element_located((find_by, identifier)))
    except:
        driver.quit()
    else:
        return element


def click_an_element(identifier, find_by=By.ID):
    wait_for_an_element(identifier, find_by).click()


def rename_file(student_name, date_collected):
    new_file = fr'{output_directory}\{student_name}_{date_collected.replace(r"/","-")}.pdf'
    if not os.path.isfile(new_file):
        os.rename(fr'{output_directory}\{default_file_name}', new_file)
    else:
        print(f"File already exists, skipped renaming the file {new_file}")


start_date = datetime.now()
options = webdriver.ChromeOptions()
settings = {
       "recentDestinations": [{
            "id": "Save as PDF",
            "origin": "local",
            "account": "",
        }],
        "selectedDestinationId": "Save as PDF",
        "version": 2
    }
prefs = {'printing.print_preview_sticky_settings.appState': json.dumps(settings), 'savefile.default_directory': f'{output_directory}'}
options.add_experimental_option('prefs', prefs)
options.add_argument(r'C:\Users\Rohit Nimmala\AppData\Local\Google\Chrome\User Data')
options.add_argument(r'--kiosk-printing')

driver = webdriver.Chrome(r'C:\Users\Rohit\Downloads\chromedriver.exe', options=options)

forms_list = [
    "https://forms.office.com/pages/designpagev2.aspx?lang=en-US&origin=OfficeDotCom&route=Start&subpage=design&id=bC4i9cZf60iPA3PbGCA7YyvECyWnxklDhRUp86g5d0NUNjNPWlBPS1ZNWTI0RklCU0NQNVpLTFg5Uy4u&analysis=true"
]

for index, form_url in enumerate(forms_list):
    driver.get(form_url)
    student_name = ""
    date_collected = ""
    if index == 0:
        wait_for_an_element('username').send_keys('nimmalrt')
        wait_for_an_element('password').send_keys('<password>')
        click_an_element(login_button_xpath, By.XPATH)
    number_of_results = wait_for_an_element(number_of_results_xpath, By.XPATH).text
    click_an_element(view_results_button_xpath, By.XPATH)
    student_name = wait_for_an_element(student_name_xpath, By.XPATH).text
    date_collected = wait_for_an_element(date_collected_css_selector, By.CSS_SELECTOR).text
    driver.execute_script('window.print();')
    rename_file(student_name, date_collected)

    for i in range(int(number_of_results)-1):
        click_an_element(next_response_xpath, By.XPATH)
        student_name = wait_for_an_element(student_name_xpath, By.XPATH).text
        date_collected = wait_for_an_element(date_collected_css_selector, By.CSS_SELECTOR).text
        driver.execute_script('window.print();')
        rename_file(student_name, date_collected)

sleep(15)
driver.close()
