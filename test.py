from selenium import webdriver
from selenium.webdriver.common.keys import Keys 
from openpyxl import load_workbook, Workbook
from datetime import datetime
import string
import time

#load Workbook
wb = load_workbook('Database.xlsx')
sheet = wb.worksheets[0]

#alphabet = list(string.ascii_uppercase)
alphabet = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z']
#get the number of people
profiles_to_encode = sheet.max_row - 1 
profile_list = list(range(1,profiles_to_encode+1))
profile_user = []

i= 0
for profile in profile_list:
    profile_data = []
    for letter in alphabet:
        test = [str(letter+str(profile))]
        
        i = i + 1
        profile_data.extend(test)
        
        print(str(i))
        if i % len(alphabet) == 0:
            print(profile_data)
            profile_user.append(profile_data)
    #profile_user.append(profile_data)
  #c profile_data.clear()
print(profile_user)
