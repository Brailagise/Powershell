from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.firefox.options import Options as FirefoxOptions
import selenium.webdriver.support.expected_conditions as EC
from time import sleep
from datetime import datetime

def GetStorageRemaining():

    dateTimeObj = datetime.now().strftime("%d-%b-%H-%M-%p")

    firefox_capabilities = DesiredCapabilities().FIREFOX
    firefox_capabilities['handleAlerts'] = True
    firefox_capabilities['acceptSslCerts'] = True
    firefox_capabilities['acceptInsecureCerts'] = True
    options = FirefoxOptions()
    options.headless = True
    profile = webdriver.FirefoxProfile()
    profile.set_preference('network.http.use-cache', False)
    profile.set_preference("javascript.enabled", True)
    profile.accept_untrusted_certs = True
    driver = webdriver.Firefox(firefox_profile=profile, firefox_binary="C:/Program Files/Mozilla Firefox/firefox.exe", executable_path="C:/Users/HunterWhitlock/Downloads/geckodriver-v0.24.0-win64/geckodriver.exe", capabilities=firefox_capabilities, options=options)
    driver.get("https://*snipped*")

    sleep(10)

    elem = driver.find_element_by_name("userName")
    elem.clear()
    elem.send_keys("*snipped*")

    elem = driver.find_element_by_name("password")
    elem.clear()
    elem.send_keys("*snipped*")

    elem.send_keys(Keys.RETURN)

    sleep(30)

    stats = driver.find_element_by_id("availStorageSpaceSpanId").text

    if len(stats) <= 2:
        driver.get_screenshot_as_file("C:/test/" + dateTimeObj + ".png")

    driver.quit()

    return(stats)

print(GetStorageRemaining())