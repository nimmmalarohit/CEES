import json
from time import sleep

from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By
import os
from base64 import b64encode, b64decode


output_directory = r"C:\Users\Rohit Nimmala\Documents\cees\output"
default_file_name = "Microsoft Forms.pdf"
student_name_xpath = "/html/body/div[2]/div/div/div/div/div[3]/div[1]/div[2]/div[1]/div/div[2]/div/div"
date_collected_css_selector = "#DatePicker_r4f38b41a4059401ca80255720cf09922"
next_response_xpath = "/html/body/div[2]/div/div/div/div/div[3]/div[1]/div[1]/div[2]/div/button[2]"
number_of_results_xpath = "/html/body/div[2]/div/div/div/div[2]/div[1]/div[2]/div/div/div[1]/div[2]/div[2]/div/div[1]/div[1]"
login_button_xpath = "/html/body/section/div/div[2]/div/form/div[3]/div/button"
view_results_button_xpath = "/html/body/div[2]/div/div/div/div[2]/div[1]/div[2]/div/div/div[1]/div[3]/button"


def get_number_of_responses(driver):
    return driver.execute_script("""
    var responses = Number(document.evaluate('/html/body/div[2]/div/div/div/div[*]/div[1]/div[2]/div/div/div[1]/div[2]/div[2]/div/div[1]/div[1]', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.innerHTML);
    return responses
    """)


def get_form_name(driver):
    return driver.execute_script("""
    var path = "/html/body/div[2]/div/div/div/div[*]/div[1]/div[2]/div/div/div[1]/div[2]/div";
    return document.evaluate(path, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.innerHTML;
    """)


def click_view_responses(driver):
    return driver.execute_script("""
    var path = "/html/body/div[2]/div/div/div/div[*]/div[1]/div[2]/div/div/div[1]/div[3]/button";
    document.evaluate(path, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.click();
    """)

def encode_password(input):
    return b64encode(input)


def decode_password(encoded_password):
    return b64decode(encoded_password).decode()


def wait_for_an_element(identifier, find_by=By.ID, wait_time=900, is_blocker=True):
    try:
        element = WebDriverWait(driver, wait_time).until(EC.visibility_of_element_located((find_by, identifier)))
    except:
        print(f"Did not find the element with the identifier: {identifier}")
        if is_blocker:
            driver.quit()
        else:
            return None
    else:
        return element


def click_an_element(identifier, find_by=By.ID, wait_time=900, is_blocker=True):
    wait_for_an_element(identifier, find_by, wait_time, is_blocker).click()


def rename_file(student_name, date_collected):
    new_file = fr'{output_directory}\{student_name}_{date_collected.replace(r"/","-")}.pdf'
    if not os.path.isfile(new_file):
        os.rename(fr'{output_directory}\{default_file_name}', new_file)
    else:
        print(f"File already exists, skipped renaming the file {new_file}")


def download_file(site_name, response_number):
    student_name = wait_for_an_element(student_name_xpath, By.XPATH, wait_time=10, is_blocker=False)
    date_collected = wait_for_an_element(date_collected_css_selector, By.CSS_SELECTOR, wait_time=10, is_blocker=False)
    if student_name and date_collected:
        student_name = student_name.text
        date_collected = date_collected.text
        driver.execute_script('window.print();')
        rename_file(student_name, date_collected)
    else:
        print(f"Student name or date field is not present for the response in the form: {site_name}")
        new_file_name = site_name.replace(' ', '_') + '_' + str(response_number)
        print(f"downloading with the file name: {new_file_name}")
        driver.execute_script('window.print();')
        rename_file(new_file_name, "")



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
    "https://forms.office.com/pages/designpagev2.aspx?lang=en-US&origin=OfficeDotCom&route=Start&subpage=design&id=bC4i9cZf60iPA3PbGCA7YyvECyWnxklDhRUp86g5d0NUNjNPWlBPS1ZNWTI0RklCU0NQNVpLTFg5Uy4u&analysis=true",
    "https://forms.office.com/Pages/DesignPageV2.aspx?origin=NeoPortalPage&subpage=design&id=bC4i9cZf60iPA3PbGCA7YwFk15odqZZBk0nbS_TJHypUQkFKN09TMVNOSVJLVjQzME9WRFM3N0xLRS4u&analysis=true",
    "https://forms.office.com/Pages/DesignPageV2.aspx?origin=NeoPortalPage&subpage=design&id=bC4i9cZf60iPA3PbGCA7YwFk15odqZZBk0nbS_TJHypUNFU5Q1ZQVjRHUzJDWlMySDU3Rjg3VU85Vy4u&analysis=true"
]

for index, form_url in enumerate(forms_list):
    driver.get(form_url)
    student_name = ""
    date_collected = ""
    if index == 0:
        wait_for_an_element('username').send_keys('nimmalrt')
        wait_for_an_element('password').send_keys(decode_password(b'VGNzY3RzbTkh'))
        click_an_element(login_button_xpath, By.XPATH)
        print("Sleeping for 30 seconds.")
        sleep(20)
    sleep(10)
    number_of_results = get_number_of_responses(driver)
    site_name = get_form_name(driver)

    if number_of_results > 0:
        print(f"Processing for the form: {site_name}, number of responses: {number_of_results}")
        click_view_responses(driver)
        download_file(site_name, 0)

    if number_of_results > 1:
        for i in range(number_of_results-1):
            click_an_element(next_response_xpath, By.XPATH)
            download_file(site_name, i+1)

driver.close()
