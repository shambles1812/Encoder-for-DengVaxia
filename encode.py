from selenium import webdriver
from selenium.webdriver.common.keys import Keys 

browser = webdriver.Chrome('C:\\Users\\Cj\\Desktop\\Temp\\chromedriver')
browser.get('https://facebook.com')
email = browser.find_element_by_id('email')
password = browser.find_element_by_id('pass')
login_button = browser.find_element_by_id('loginbutton')
email.send_keys('cj.villafuerte@yahoo.com')
password.send_keys('Levarya23')
login_button.click()
wait = input('Go to messenger')
browser.get('https://www.facebook.com/messages/t/')
disclaimer = " - This message was sent by a python bot"
message = input('Please add a message')
jalen = browser.get_element_by_id('js_d')
def full_message(message, disclaimer):
    return message + disclaimer
jalen.click()

