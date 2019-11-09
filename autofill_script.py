from selenium import webdriver
from selenium.webdriver.support import ui
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
import datetime as dt
from multiprocessing.pool import ThreadPool
import time
import pandas as pd
import argparse
import timeit

# LINK = 'http://laodongkynghi.dolab.gov.vn/dang-ky-cap-thu-gioi-thieu/ung-vien/mau-dang-ky'
FORM_XPATH = "//form"

def page_is_loaded(driver):
    return driver.find_element_by_xpath(FORM_XPATH) != None

def fill_text_xpath(dn, driver, xpath, txt):
    try:
        e = driver.find_element_by_xpath(xpath)
        e.clear()
        e.send_keys(txt)
    except Exception as ex:
        log("[ERROR][TEXT][%s] %s" % (dn, str(ex)))

def choose_option_click(dn, driver, xpath):
    try:
        driver.find_element_by_xpath(xpath).click()
    except Exception as ex:
        log("[ERROR][CLICK][%s] %s" % (dn, str(ex)))

def save_screenshot(driver, path):
    # Ref: https://stackoverflow.com/a/52572919/
    original_size = driver.get_window_size()
    required_width = driver.execute_script('return document.body.parentNode.scrollWidth')
    required_height = driver.execute_script('return document.body.parentNode.scrollHeight')
    driver.set_window_size(required_width, required_height)
    # driver.save_screenshot(path)  # has scrollbar
    driver.find_element_by_xpath(FORM_XPATH).screenshot(path)  # avoids scrollbar
    driver.set_window_size(original_size['width'], original_size['height'])

def log(txt):
    timestamp_now = dt.datetime.today().strftime('%Y-%m-%d %H:%M:%S')
    print("[%s]%s" % (timestamp_now, txt))


xpath_config = {
    "FullName": "(//form/div[@class=\"form-group\"][1]//input)[1]",
    "DOB": "(//form/div[@class=\"form-group\"][1]//input)[2]",
    "Sex": "(//form/div[@class=\"form-group\"][1]//select)[1]/option[contains(.,'{}')]",
    "Email": "(//form/div[@class=\"form-group\"][2]//input)[1]",
    "Phone": "(//form/div[@class=\"form-group\"][2]//input)[2]",
    "Marriage": "(//form/div[@class=\"form-group\"][2]//select)[1]/option[contains(.,'{}')]",
    "Address": "(//form/div[@class=\"form-group\"][3]//input)[1]",
    "Province": "(//form/div[@class=\"form-group\"][4]//select)[1]/option[contains(.,'{}')]",
    "District": "(//form/div[@class=\"form-group\"][4]//select)[2]/option[contains(.,'{}')]",
    "Town": "(//form/div[@class=\"form-group\"][4]//input)[1]",
    "CurrentAddress": "(//form/div[@class=\"form-group\"][5]//input)[1]",
    "PassportNum": "(//form/div[@class=\"form-group\"][6]//input)[1]",
    "PassportPlace": "(//form/div[@class=\"form-group\"][6]//input)[2]",
    "PassportFromDate": "(//form/div[@class=\"form-group\"][6]//input)[3]",
    "PassportToDate": "(//form/div[@class=\"form-group\"][6]//input)[4]",
    "IsGraduate": "(//form/div[@class=\"form-group\"][7]//input)[{}]",
    "University": "(//form/div[@class=\"form-group\"][7]/following-sibling::div/div[1]//input)[1]",
    "Fulltime": "(//form/div[@class=\"form-group\"][7]/following-sibling::div/div[1]//select)[1]/option[{}]",
    "DiplomaCode": "(//form/div[@class=\"form-group\"][7]/following-sibling::div/div[2]//input)[1]",
    "FromYear": "(//form/div[@class=\"form-group\"][7]/following-sibling::div/div[2]//input)[2]",
    "ToYear": "(//form/div[@class=\"form-group\"][7]/following-sibling::div/div[2]//input)[3]",
    "NumOfYear": "(//form/div[@class=\"form-group\"][7]/following-sibling::div/div[2]//input)[4]",
    "GraduateYear": "(//form/div[@class=\"form-group\"][7]/following-sibling::div/div[2]//input)[5]",
    "ContactPerson": "(//form/div[@class=\"form-group\"][8]//input)[1]",
    "ContactPersonPhone": "(//form/div[@class=\"form-group\"][8]//input)[2]",
    "ContactPersonAddress": "(//form/div[@class=\"form-group\"][9]//input)[1]",
    "CheckBtn": "//*[@id=\"WNextForm\"]"
}

if __name__ == "__main__":

    parser = argparse.ArgumentParser(description='Choose batch.')
    parser.add_argument('--drivers', type=int, help='Number of drivers', default=2, required=True)
    parser.add_argument('--wait', type=int, help='Wait to load page (sec)', default=2, required=True)
    parser.add_argument('--browser', type=str, help='Browser type', default='chrome', required=True)
    parser.add_argument('--link', type=str, help='Link to regist form', required=True)
    args = parser.parse_args()

    NUM_DRIVERS = args.drivers
    BROWSER = args.browser.lower().strip()
    LINK_FORM = args.link.strip()

    WAIT_TIME = args.wait

    try:

        def fill_form_one_driver(driver, driver_name, data):
            log("[%s] Start fill form" % (driver_name))

            # Row 1
            fill_text_xpath(driver_name, driver, xpath_config['FullName'], data['FullName'])
            fill_text_xpath(driver_name, driver, xpath_config['DOB'], data['DOB'])
            # choose_option_click(driver_name, driver, "/html/body")

            choose_option_click(driver_name, driver, xpath_config['Sex'].format(data['Sex']))

            # Row 2
            fill_text_xpath(driver_name, driver, xpath_config['Email'], data['Email'])
            fill_text_xpath(driver_name, driver, xpath_config['Phone'], data['Phone'])
            choose_option_click(driver_name, driver, xpath_config['Marriage'].format(data['Marriage']))

            # Row 3
            fill_text_xpath(driver_name, driver, xpath_config['Address'], data['Address'])

            # Row 4
            choose_option_click(driver_name, driver, xpath_config['Province'].format(data['Province']))
            time.sleep(1.5)
            choose_option_click(driver_name, driver, xpath_config['District'].format(data['District']))
            fill_text_xpath(driver_name, driver, xpath_config['Town'], data['Town'])

            # Row 5
            fill_text_xpath(driver_name, driver, xpath_config['CurrentAddress'], data['CurrentAddress'])

            # Row 6
            fill_text_xpath(driver_name, driver, xpath_config['PassportNum'], data['PassportNum'])
            fill_text_xpath(driver_name, driver, xpath_config['PassportPlace'], data['PassportPlace'])
            fill_text_xpath(driver_name, driver, xpath_config['PassportFromDate'], data['PassportFromDate'])
            fill_text_xpath(driver_name, driver, xpath_config['PassportToDate'], data['PassportToDate'])

            # Row 7
            choose_option_click(driver_name, driver, xpath_config['IsGraduate'].format(data['IsGraduate']))
            # time.sleep(0.5)

            # Row 7.1
            fill_text_xpath(driver_name, driver, xpath_config['University'], data['University'])
            choose_option_click(driver_name, driver, xpath_config['Fulltime'].format(data['Fulltime']))

            # Row 7.2
            fill_text_xpath(driver_name, driver, xpath_config['DiplomaCode'], data['DiplomaCode'])
            fill_text_xpath(driver_name, driver, xpath_config['FromYear'], data['FromYear'])
            fill_text_xpath(driver_name, driver, xpath_config['ToYear'], data['ToYear'])
            fill_text_xpath(driver_name, driver, xpath_config['NumOfYear'], data['NumOfYear'])
            fill_text_xpath(driver_name, driver, xpath_config['GraduateYear'], data['GraduateYear'])

            # Row 8
            fill_text_xpath(driver_name, driver, xpath_config['ContactPerson'], data['ContactPerson'])
            fill_text_xpath(driver_name, driver, xpath_config['ContactPersonPhone'], data['ContactPersonPhone'])

            # Row 9
            fill_text_xpath(driver_name, driver, xpath_config['ContactPersonAddress'], data['ContactPersonAddress'])

            # choose_option_click(driver_name, driver, xpath_config['CheckBtn'])

            log("[%s] Finish fill form" % (driver_name))

            #timestamp_now = dt.datetime.today().strftime('%Y-%m-%d_%H-%M-%S')
            #screenshot_filename = r"./form_screenshot_%s_%s.png" % (driver_name, timestamp_now)
            #save_screenshot(driver, screenshot_filename)

            #timestamp_now = dt.datetime.today().strftime('%Y-%m-%d %H:%M:%S')
            #print("[%s][%s] Save form screenshot: %s" % (timestamp_now, driver_name, screenshot_filename))

        def run_driver(dn, driver, link_form, data):
            try:
                log("[%s] Start driver" % (dn))
                driver.maximize_window()
                driver.get(link_form)

                # wait = ui.WebDriverWait(driver, 5)
                # wait.until(page_is_loaded)
                try:
                    wait = WebDriverWait(driver, WAIT_TIME).until(ec.visibility_of_element_located((By.XPATH, '//form')))
                    log("[%s] Page is ready!" % (dn))
                except TimeoutException:
                    log("[%s] Loading took too much time!" % (dn))
                
                log("[%s] Load page completed" % (dn))

                # time.sleep(WAIT_TIME)
                fill_form_one_driver(driver, dn, data)

                # driver.close()
                log("[%s] Close driver" % (dn))
            except Exception as e:
                log("[ERROR][%s] %s" % (dn, str(e)))

        start_time = timeit.default_timer()

        pool = ThreadPool(NUM_DRIVERS)

        if BROWSER == 'chrome':
            from selenium.webdriver.chrome.options import Options
            caps = DesiredCapabilities().CHROME
            caps["pageLoadStrategy"] = "normal"
        elif BROWSER == 'firefox':
            from selenium.webdriver.firefox.options import Options
            caps = DesiredCapabilities().FIREFOX
            caps["pageLoadStrategy"] = "none"
        else:
            log("[ERROR] No support browser: %s" % (LINK_FORM))
        
        options = Options()
        options.add_argument("window-size=1920,1080")

        driver_names = ['driver_' + str(i) for i in range(1, NUM_DRIVERS + 1)]                 
        driver_tuples = pool.map(lambda x: (x, webdriver.Chrome(desired_capabilities=caps, chrome_options=options)), driver_names)

        src_xsl = pd.ExcelFile('./Data.xlsx')
        df = src_xsl.parse(0)
        df = df.iloc[:,:-1]
        df.columns = ['k', 'v']
        df['v'] = df['v'].fillna('').apply(lambda x: str(x).strip())
        data = df.set_index('k').to_dict()['v']
        
        res = pool.map(lambda x: run_driver(x[0], x[1], LINK_FORM, data), driver_tuples)

        pt = round((timeit.default_timer() - start_time), 2)
        print("Processing time: %s" % (str(pt)))
    
    except Exception as e:
        print(e)
