import json
from time import sleep
from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By
import os
from base64 import b64encode, b64decode
from datetime import datetime, date as dt


output_directory = r"C:\Users\Rohit Nimmala\Documents\cees\output"
default_file_name = "Microsoft Forms.pdf"
student_name_xpath = "/html/body/div[2]/div/div/div/div/div[3]/div[1]/div[2]/div[1]/div[*]/div[2]/div/div"
student_name_xpath2 = "/html/body/div[2]/div/div/div/div/div[3]/div[1]/div[2]/div[1]/div[*]/div[2]/div/div[2]/div/div"
number_of_results_xpath = "/html/body/div[2]/div/div/div/div[2]/div[1]/div[2]/div/div/div[1]/div[2]/div[2]/div/div[1]/div[1]"
login_button_xpath = "/html/body/section/div/div[2]/div/form/div[3]/div/button"
view_results_button_xpath = "/html/body/div[2]/div/div/div/div[2]/div[1]/div[2]/div/div/div[1]/div[3]/button"
current_days_data = []


def html_generator(data):
    html_head = """
    <head>
    <title>Internship Data Collection</title>
    <style>
        table {
            border-collapse: collapse;
            width: 60%;
            }
            
        th, td {
            text-align: left;
            padding: 8px;
            border-bottom: 1px solid #ddd;
            }
        
        tr:nth-child(even) {
            background-color: #f2f2f2;
            }

        .yes {
            color: white;
            background-color: green;
            padding: 6px 12px;
            border-radius: 4px;
            }

        .no {
            color: white;
            background-color: red;
            padding: 6px 12px;
            border-radius: 4px;
            }
    </style>
    </head>
    """
    table_content = ""
    for student_name, site_name, input_date in data:
        table_content = table_content + f"""
        <tr>
            <td>{student_name}</td>
            <td>{site_name}</td>
            <td>{input_date}</td>
            <td class="yes">Yes</td>
        </tr>
        """

    html_string = f"""
    <!DOCTYPE html>
    <html>
    {html_head}
    <body>
        <h1>Data Collection Report</h1>
        <table>
            <thead>
                <tr>
                    <th>Name</th>
                    <th>Location</th>
                    <th>Date</th>
                    <th>Filled</th>
                </tr>
            </thead>
            <tbody>
            {table_content}				
            </tbody>
        </table>
    </body>
    </html>
    """
    return html_string


def store_current_day_data(student_name, site_name, input_date):
    current_date = dt.today()
    # current_date = dt(2023, 3, 6)
    input_date = datetime.strptime(input_date, '%m/%d/%Y').date()
    if input_date == current_date:
        current_days_data.append([student_name, site_name, str(input_date)])
        print(f"Today's record found for student_name: {student_name} , site_name: {site_name}, input_date: {str(input_date)}")


def get_next_response(driver):
    driver.execute_script("""
    function getElementByXpath(path) {
    return document.evaluate(path, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
    }
    getElementByXpath("/html/body/div[2]/div/div/div/div/div[3]/div[1]/div[1]/div[*]/div/button[2]").click();
    """)


def get_number_of_responses(driver):
    return driver.execute_script("""
    var path = '/html/body/div[2]/div/div/div/div[*]/div[1]/div[2]/div/div/div[1]/div[2]/div[2]/div/div[1]/div[1]';
    var responses = document.evaluate(path, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
    if (responses != null){
    output = Number(responses.innerHTML);
    }
    else{
    output = null;
    };
    return output
    """)


def get_form_name(driver):
    return driver.execute_script("""
    var path = "/html/body/div[2]/div/div/div/div[*]/div[1]/div[2]/div/div/div[1]/div[2]/div";
    var res = document.evaluate(path, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
    if (res != null){
        var output = res.innerHTML;
    }
    else{
        var output = null;
    };
    return output;
    """)


def click_view_responses(driver):
    return driver.execute_script("""
    var path = "/html/body/div[2]/div/div/div/div[*]/div[1]/div[2]/div/div/div[1]/div[3]/button";
    document.evaluate(path, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.click();
    """)


def get_input_date(driver):
    output = driver.execute_script("""
    var id = document.querySelector('[id^="DatePicker_"]');
    if (id != null){
    var return_value = document.getElementById(id.id).innerHTML
    }
    else{
    var return_value = null
    };
    return return_value;
    """)

    if output and 'Please input date' in output:
        return None
    elif output:
        return output
    else:
        return None


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
    response_number = response_number + 1
    print(f"trying to download the file: {site_name} and response {response_number}")
    student_name = wait_for_an_element(student_name_xpath, By.XPATH, wait_time=5, is_blocker=False)
    if student_name and 'Student Name' in student_name.text:
        student_name = wait_for_an_element(student_name_xpath2, By.XPATH, wait_time=5, is_blocker=False)
    date_collected = get_input_date(driver)
    if student_name and date_collected:
        student_name = student_name.text
        driver.execute_script('window.print();')
        rename_file(student_name, date_collected + '_' + site_name.replace(' ', '_'))
        print("Downloaded the file")
        store_current_day_data(student_name, site_name, date_collected)
    else:
        print(f"Student name or date field is not present for the response in the form: {site_name}")
        new_file_name = site_name.replace(' ', '_') + '_' + str(response_number)
        print(f"Alternatively downloading with the file name: {new_file_name}")
        driver.execute_script('window.print();')
        rename_file(new_file_name, "")
    sleep(2)



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
    "https://forms.office.com/Pages/DesignPageV2.aspx?origin=NeoPortalPage&subpage=design&collectionid=ci0npmuziej4ua0lscxuny&id=bC4i9cZf60iPA3PbGCA7Y2URZC97vDJCtT_H-EGhuJhUQUw1UDNIRUdOWTlFTExHUjMxNUIyNjJHSy4u&analysis=true",
    "https://forms.office.com/Pages/DesignPageV2.aspx?origin=NeoPortalPage&subpage=design&collectionid=ci0npmuziej4ua0lscxuny&id=bC4i9cZf60iPA3PbGCA7Y2URZC97vDJCtT_H-EGhuJhUQ0VYSkRLWTMwMFlHNFFONkdTQVpCNTBLSy4u&analysis=true",
    "https://forms.office.com/Pages/DesignPageV2.aspx?origin=NeoPortalPage&subpage=design&id=bC4i9cZf60iPA3PbGCA7Y2URZC97vDJCtT_H-EGhuJhUQkJYQlhFT0VKUjdDODI4TUdaUFRLRE9YUC4u&analysis=true",
    "https://forms.office.com/Pages/DesignPageV2.aspx?origin=NeoPortalPage&subpage=design&id=bC4i9cZf60iPA3PbGCA7Y2URZC97vDJCtT_H-EGhuJhUMThZNUxWUFdFUldSTTE4NzZKTUVKSVpCTC4u&analysis=true",
    "https://forms.office.com/Pages/DesignPageV2.aspx?origin=NeoPortalPage&subpage=design&id=bC4i9cZf60iPA3PbGCA7YwFk15odqZZBk0nbS_TJHypURU45SlRPOEJMWTgxOFNZVzFHTlpFUFRKUi4u&analysis=true",
    "https://forms.office.com/Pages/DesignPageV2.aspx?origin=NeoPortalPage&subpage=design&id=bC4i9cZf60iPA3PbGCA7YwFk15odqZZBk0nbS_TJHypURjI5UTNESkRTTzdFSUVHNzZHV1hWMEExUC4u&analysis=true",
    "https://forms.office.com/Pages/DesignPageV2.aspx?origin=NeoPortalPage&subpage=design&id=bC4i9cZf60iPA3PbGCA7YwFk15odqZZBk0nbS_TJHypUNFU5Q1ZQVjRHUzJDWlMySDU3Rjg3VU85Vy4u&analysis=true",
    "https://forms.office.com/Pages/DesignPageV2.aspx?origin=NeoPortalPage&subpage=design&id=bC4i9cZf60iPA3PbGCA7YwFk15odqZZBk0nbS_TJHypUQkFKN09TMVNOSVJLVjQzME9WRFM3N0xLRS4u&analysis=true",
    "https://forms.office.com/Pages/DesignPageV2.aspx?origin=NeoPortalPage&subpage=design&id=bC4i9cZf60iPA3PbGCA7YwFk15odqZZBk0nbS_TJHypUMEdVWlMySzg5VjAwR0tLWFFHUDBFRFc0NC4u&analysis=true",
    "https://forms.office.com/Pages/DesignPageV2.aspx?origin=NeoPortalPage&subpage=design&id=bC4i9cZf60iPA3PbGCA7YwFk15odqZZBk0nbS_TJHypUM0I5SkhNMUZRRkUwNzlKN0pBT0s0OVM4Ti4u&analysis=true",
    "https://forms.office.com/Pages/DesignPageV2.aspx?origin=NeoPortalPage&subpage=design&id=bC4i9cZf60iPA3PbGCA7YwFk15odqZZBk0nbS_TJHypUNEg4WTRRRUMyQU9TWUdLUTA1ME8wTlVETy4u&analysis=true",
    "https://forms.office.com/Pages/DesignPageV2.aspx?origin=NeoPortalPage&subpage=design&id=bC4i9cZf60iPA3PbGCA7YwFk15odqZZBk0nbS_TJHypUMVMzMEJYQTlXSzhJWTFDV0YwVzc2MDlWMy4u&analysis=true",
    "https://forms.office.com/Pages/DesignPageV2.aspx?origin=NeoPortalPage&subpage=design&id=bC4i9cZf60iPA3PbGCA7YwFk15odqZZBk0nbS_TJHypUMVlXS0dFNVFYMlNRQVAzWUpQTDBYVE8wSy4u&analysis=true",
    "https://forms.office.com/Pages/DesignPageV2.aspx?origin=NeoPortalPage&subpage=design&collectionid=ci0npmuziej4ua0lscxuny&id=bC4i9cZf60iPA3PbGCA7YyvECyWnxklDhRUp86g5d0NUQk5SMFlWUFU1UzFVVzhJVllXMlpSWThUNy4u&analysis=true",
    "https://forms.office.com/Pages/DesignPageV2.aspx?origin=NeoPortalPage&subpage=design&id=bC4i9cZf60iPA3PbGCA7YwFk15odqZZBk0nbS_TJHypUNzM3MjRUSVBTVzY0MVdORVFDRFJHUVNWWi4u&analysis=true",
    "https://forms.office.com/Pages/DesignPageV2.aspx?origin=NeoPortalPage&subpage=design&collectionid=ci0npmuziej4ua0lscxuny&id=bC4i9cZf60iPA3PbGCA7YyvECyWnxklDhRUp86g5d0NURE1JWkpHN0kzM0JFT0FBTVhBTjUwN1kwUy4u&analysis=true",
    "https://forms.office.com/Pages/DesignPageV2.aspx?origin=NeoPortalPage&subpage=design&collectionid=ci0npmuziej4ua0lscxuny&id=bC4i9cZf60iPA3PbGCA7YyvECyWnxklDhRUp86g5d0NUQkFaVTM5WUVUMVFMQUxQR1RIWjZZMzFTNS4u&analysis=true",
    "https://forms.office.com/Pages/DesignPageV2.aspx?origin=NeoPortalPage&subpage=design&collectionid=ci0npmuziej4ua0lscxuny&id=bC4i9cZf60iPA3PbGCA7YyvECyWnxklDhRUp86g5d0NUMTYwUDZCTkZHUEIzMUc4NVU0TEY1UU9INS4u&analysis=true",
    "https://forms.office.com/Pages/DesignPageV2.aspx?origin=NeoPortalPage&subpage=design&collectionid=ci0npmuziej4ua0lscxuny&id=bC4i9cZf60iPA3PbGCA7YyvECyWnxklDhRUp86g5d0NUN0g1NEtaVDMyM0YyWkNaSDNCTzNUWTVaRC4u&analysis=true",
    "https://forms.office.com/Pages/DesignPageV2.aspx?origin=NeoPortalPage&subpage=design&collectionid=ci0npmuziej4ua0lscxuny&id=bC4i9cZf60iPA3PbGCA7YyvECyWnxklDhRUp86g5d0NUREtRRkdCTjRIWkRaOUJaR0pQQVZaUFZLMi4u&analysis=true",
    "https://forms.office.com/Pages/DesignPageV2.aspx?origin=NeoPortalPage&subpage=design&collectionid=ci0npmuziej4ua0lscxuny&id=bC4i9cZf60iPA3PbGCA7YyvECyWnxklDhRUp86g5d0NUQUgwWkpRQzFPN1lRM1kxR1hJQkZVNVRRNy4u&analysis=true",
    "https://forms.office.com/Pages/DesignPageV2.aspx?origin=NeoPortalPage&subpage=design&id=bC4i9cZf60iPA3PbGCA7YwFk15odqZZBk0nbS_TJHypUNEE0N1FMVk9ZSDVWSURMUUg4UEJLSjRMRi4u&analysis=true",
    "https://forms.office.com/Pages/DesignPageV2.aspx?origin=NeoPortalPage&subpage=design&collectionid=ci0npmuziej4ua0lscxuny&id=bC4i9cZf60iPA3PbGCA7YyvECyWnxklDhRUp86g5d0NUM1VENlZDSEFRVzFOWDBPRk5YOEFIOVZURi4u&analysis=true",
    "https://forms.office.com/Pages/DesignPageV2.aspx?origin=NeoPortalPage&subpage=design&collectionid=ci0npmuziej4ua0lscxuny&id=bC4i9cZf60iPA3PbGCA7YyvECyWnxklDhRUp86g5d0NUNzBSMldPSkg1UTg4V1dFUDI5TldMQUZWMS4u&analysis=true",
    "https://forms.office.com/Pages/DesignPageV2.aspx?origin=NeoPortalPage&subpage=design&collectionid=ci0npmuziej4ua0lscxuny&id=bC4i9cZf60iPA3PbGCA7YyvECyWnxklDhRUp86g5d0NUM1pHSzVDNjBORlk1NEQ5U0laUTZWN0tZWC4u&analysis=true",
    "https://forms.office.com/Pages/DesignPageV2.aspx?origin=NeoPortalPage&subpage=design&collectionid=ci0npmuziej4ua0lscxuny&id=bC4i9cZf60iPA3PbGCA7YyvECyWnxklDhRUp86g5d0NUQVNDRzJENDhCTzVISjdVMUFZREhXRTcxRS4u&analysis=true",
    "https://forms.office.com/Pages/DesignPageV2.aspx?origin=NeoPortalPage&subpage=design&collectionid=ci0npmuziej4ua0lscxuny&id=bC4i9cZf60iPA3PbGCA7YyvECyWnxklDhRUp86g5d0NUNjNPWlBPS1ZNWTI0RklCU0NQNVpLTFg5Uy4u&analysis=true"
]

for index, form_url in enumerate(forms_list):
    driver.get(form_url)
    student_name = ""
    date_collected = ""
    if index == 0:
        wait_for_an_element('username').send_keys('nimmalrt')
        wait_for_an_element('password').send_keys(decode_password(b'VGNzY3RzbTkh'))
        click_an_element(login_button_xpath, By.XPATH)
        print("Sleeping for 20 seconds.")
        sleep(20)
    sleep(3)
    number_of_results = get_number_of_responses(driver)
    site_name = get_form_name(driver)

    if number_of_results:
        if number_of_results > 0:
            print(f"Processing for the form: {site_name}, number of responses: {number_of_results}")
            click_view_responses(driver)
            download_file(site_name, 0)
        else:
            print(f"No responses for the form: {site_name}, number of responses: {number_of_results}")

        if number_of_results > 1:
            for i in range(number_of_results-1):
                get_next_response(driver)
                download_file(site_name, i+1)
        print(f"Completed saving responses from the form: {site_name}")
    else:
        print(f"Unsuccessful in capturing number of responses for the site : {site_name} and responses values: {number_of_results}")

driver.close()
print("printing current days data:")
print(current_days_data)

email_html = html_generator(current_days_data)
with open(output_directory + '/' + 'data_collection_report_' + str(datetime.today().date()) + '.html', 'w') as email_file:
    email_file.write(email_html)
