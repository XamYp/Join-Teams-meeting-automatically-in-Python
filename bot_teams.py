__author__ = "XamYp"
__copyright__ = "You can fork the project and improve it, it is open source."
__credits__ = "XamYp"
__license__ = "GPL"
__version__ = "1.0.0"
__maintainer__ = "XamYp"
__email__ = "maxence.peligry@gmail.com"
__status__ = "Test"
__note__ = """ It is a test project that meets my needs according to my specifications. 
              I made a simple code, do not take into account the improvement that I 
              could have made but you are free to improve it. """

from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from time import sleep
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait as wait
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.chrome.options import Options
import os
import configparser
import datetime as dt
import sys

path = os.path.dirname(os.path.abspath(__file__))
chromeDriver = os.path.join(path, 'chromedriver.exe')
configFile = os.path.join(path, 'config.ini')
configIni = configparser.RawConfigParser()
configIni.read(configFile, encoding='utf-8')
hourAndMinute = configIni.get('Time','HourAndMinuteToLeaveMeeting').partition(':')

dateNow = dt.datetime.now()
dateToLeave = dt.datetime(dateNow.year,dateNow.month,dateNow.day,int(hourAndMinute[0]),int(hourAndMinute[2]),00)

opt = Options()
opt.add_argument("--disable-infobars")
opt.add_argument("--disable-extensions")
opt.add_argument("--disable-popup-window")
opt.add_argument("disable-gpu")
opt.add_argument("--disable-popup-blocking")
opt.add_experimental_option("excludeSwitches", ["enable-automation"])
opt.add_experimental_option('useAutomationExtension', False)
opt.add_experimental_option("prefs", { \
    "profile.default_content_setting_values.media_stream_mic": 1, 
    "profile.default_content_setting_values.media_stream_camera": 1,
    "profile.default_content_setting_values.geolocation": 1, 
    "profile.default_content_setting_values.notifications": 1,
    "excludeSwitches": ["disable-popup-blocking"],
    "useAutomationExtension": 0
  })

wd = webdriver.Chrome(chrome_options=opt, executable_path=chromeDriver)
wd.implicitly_wait(10)
wd.delete_all_cookies()
wd.get("https://login.microsoftonline.com/common/oauth2/authorize?response_type=id_token&client_id=5e3ce6c0-2b1f-4285-8d4b-75ee78787346&redirect_uri=https%3A%2F%2Fteams.microsoft.com%2Fgo&state=5928e638-bb27-434a-9b46-93d66c2b1dd5&&client-request-id=fa8446a5-4e87-4ef5-88da-8ac51b305930&x-client-SKU=Js&x-client-Ver=1.0.9&nonce=eb3c2fb1-2b1f-4085-a008-a5f3bb4a8004&domain_hint=")
emailLogin = wd.find_element_by_xpath('//*[@id="i0116"]')
emailLogin.clear()
emailLogin.send_keys(configIni.get('Account','MicrosoftAccount'))
wd.find_element_by_xpath('//*[@id="idSIButton9"]').click()
sleep(2)
passwordLogin = wd.find_element_by_xpath('//*[@id="i0118"]')
passwordLogin.clear()
passwordLogin.send_keys(configIni.get('Account','MicrosoftPassword'))
wd.find_element_by_xpath('//*[@id="idSIButton9"]').click()
sleep(2)
wd.find_element_by_xpath('//*[@id="idSIButton9"]').click()
sleep(10)
wd.execute_script("window.open('about:blank', 'secondtab');")
wd.switch_to.window("secondtab")
wd.get(configIni.get('Url','TeamsMeetingUrl'))
wd.execute_script("window.open('about:blank', 'thirdtab');")
wd.switch_to.window("thirdtab")
wd.get(configIni.get('Url','TeamsMeetingUrl'))
sleep(5)
wd.find_element_by_xpath('//*[@id="buttonsbox"]/button[2]').click()
sleep(10)
wd.find_element_by_xpath('//*[@id="page-content-wrapper"]/div[1]/div/calling-pre-join-screen/div/div/div[2]/div[1]/div[2]/div/div/section/div[2]/toggle-button[1]/div/button/span[1]').click()
sleep(1)
wd.find_element_by_xpath('//*[@id="preJoinAudioButton"]/div/button/span[1]').click()
sleep(1)
wd.find_element_by_xpath('//*[@id="page-content-wrapper"]/div[1]/div/calling-pre-join-screen/div/div/div[2]/div[1]/div[2]/div/div/section/div[1]/div/div/button').click()
wd.switch_to.window(wd.window_handles[0])
wd.close()
wd.switch_to.window(wd.window_handles[0])
wd.close()
sleep((dateToLeave-dateNow).total_seconds())
wd.switch_to.window(wd.window_handles[0])
sleep(1)
wd.find_element_by_xpath('//*[@id="personDropdown"]').click()
sleep(2)
wd.find_element_by_xpath('//*[@id="logout-button"]').click()
sleep(2)
wd.quit()


