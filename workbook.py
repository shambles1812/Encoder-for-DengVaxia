from selenium import webdriver
from selenium.webdriver.common.keys import Keys 
from openpyxl import load_workbook, Workbook
from datetime import datetime
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Color, PatternFill, Font, Border
from openpyxl.styles import colors
from openpyxl.cell import Cell
import string
import time

browser = webdriver.Chrome('C:\\Users\\Cj\\Desktop\\Temp\\chromedriver')
browser.get('http://facebook.com')
user_field = browser.find_element_by_id('email')
#pass_field = browser.find_element_by_id('Password')
#login = browser.find_element_by_xpath("//input[@value='Log in']")
user_field.send_keys()
user_field.send_keys(Keys.TAB)
current_element = browser.switch_to_active_element()
current_element.send_keys()
current_element.send_keys(Keys.TAB)
current_element = browser.switch_to_active_element()
current_element.send_keys(Keys.ENTER)
input('Wait')
#pass_field.send_keys('Blackhatred1') 
