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
profile_users = []
i= 0
#create list of 
for profile in profile_list:
    profile_data = []
    for letter in alphabet:
        test = [str(letter+str(profile+1))]
        i = i + 1
        profile_data.extend(test)
        if i % len(alphabet) == 0:
            profile_users.append(profile_data)
    #profile_user.append(profile_data)
  #c profile_data.clear()
#load resources

#get first name
def first_name(user_number):
    if sheet[profile_users[user_number-1][0]].value:
        return sheet[profile_users[user_number-1][0]].value
def last_name(user_number):
    if sheet[profile_users[user_number-1][1]].value:
        return sheet[profile_users[user_number-1][1]].value
def middle_initial(user_number): 
    if sheet[profile_users[user_number-1][2]].value:
        return sheet[profile_users[user_number-1][2]].value 
def rel_to_household(user_number):
    if sheet[profile_users[user_number-1][3]].value:
        return sheet[profile_users[user_number-1][3]].value
def resp(user_number):
    if sheet[profile_users[user_number-1][4]].value:
        return sheet[profile_users[user_number-1][4]].value
def contact_no(user_number):
    if sheet[profile_users[user_number-1][5]].value:
        return sheet[profile_users[user_number-1][5]].value
def province(user_number):
    if sheet[profile_users[user_number-1][6]].value:
        return sheet[profile_users[user_number-1][6]].value
def city(user_number):
    if sheet[profile_users[user_number-1][7]].value:
        return sheet[profile_users[user_number-1][7]].value
def Barangay(user_number):
    if sheet[profile_users[user_number-1][8]].value:
        return sheet[profile_users[user_number-1][8]].value
def street(user_number):
    if sheet[profile_users[user_number-1][9]].value:
        return sheet[profile_users[user_number-1][9]].value
def house_no(user_number):
    if sheet[profile_users[user_number-1][10]].value:
        return sheet[profile_users[user_number-1][10]].value
def sex(user_number):
    if sheet[profile_users[user_number-1][11]].value:
        return sheet[profile_users[user_number-1][11]].value
def age(user_number):
    if sheet[profile_users[user_number-1][12]].value:
        return sheet[profile_users[user_number-1][12]].value
def religion(user_number):
    if sheet[profile_users[user_number-1][13]].value:
        return sheet[profile_users[user_number-1][13]].value
def date_of_birth(user_number):
    if sheet[profile_users[user_number-1][14]].value:
        return sheet[profile_users[user_number-1][14]].value
def birthplace(user_number):
    if sheet[profile_users[user_number-1][15]].value:
        return sheet[profile_users[user_number-1][15]].value
def yrs(user_number):
    if sheet[profile_users[user_number-1][16]].value:
        return sheet[profile_users[user_number-1][16]].value
def id_no(user_number):
    if sheet[profile_users[user_number-1][17]].value:
        return sheet[profile_users[user_number-1][17]].value

def get_info_on_page_1(user_number):
    print(first_name(user_number))
    print(last_name(user_number))
    print(middle_initial(user_number))
    print(rel_to_household(user_number))
    print(resp(user_number))
    print(contact_no(user_number))
    print(province(user_number))
    print(city(user_number))
    print(Barangay(user_number))
    print(street(user_number))
    print(house_no(user_number))
    print(sex(user_number))
    print(age(user_number))
    print(religion(user_number))
    print(date_of_birth(user_number))
    print(birthplace(user_number))
    print(yrs(user_number))
    print(id_no(user_number))
    

#login
browser = webdriver.Chrome('C:\\Users\\Cj\\Desktop\\Temp\\chromedriver')
browser.get('http://122.54.82.254:65428/')
user_field = browser.find_element_by_id('UserName')
pass_field = browser.find_element_by_id('Password')
login = browser.find_element_by_xpath("//input[@value='Log in']")
user_field.send_keys('NATE155')
pass_field.send_keys('123456')
login.click()

#navigate to vaccination entry page
create_new_profile = browser.find_element_by_xpath("//button[@data-toggle='modal']")
time.sleep(2)
create_new_profile.click()
add_vacinee_profile = browser.find_element_by_xpath("//a[@href='/VaccinneeProfile/Add']")
time.sleep(2)
add_vacinee_profile.click()
time.sleep(2)

# paths 
id_field = browser.find_element_by_id('VaccinationCardId')
date_interviewed_field = browser.find_element_by_id('DateInterviewed')
relation_to_vaccined = browser.find_element_by_id('RelasyonSaBakunado')
last_name_field = browser.find_element_by_id('LastName')
first_name_field = browser.find_element_by_id('FirstName')
middle_inintial_field = browser.find_element_by_id('MiddleName')
relation_to_household = browser.find_element_by_id('RelationToHouseholdHead')
respondent = browser.find_element_by_id('Respondent')
contact_no_field = browser.find_element_by_id('ContactNumber')
female_selection = browser.find_element_by_xpath("//input[@value='Female']")
male_selection = browser.find_element_by_xpath("//input[@value='Male']")
age_field = browser.find_element_by_id('Age')
birthplace_field = browser.find_element_by_id('BirthPlace')
years_at_curr = browser.find_element_by_id('YearsInAddress')
province_option = browser.find_element_by_id('s2id_PSGC_ProvinceId')

city_option = browser.find_element_by_id('s2id_PSGC_CityMunicipalityId')
barangay_option = browser.find_element_by_id('s2id_PSGC_BarangayId')
street_field = browser.find_element_by_id('Street')
house_no_field = browser.find_element_by_id('LotNo')
birthdate_field = browser.find_element_by_id('BirthDate')
roman_catholic_selection = browser.find_element_by_xpath("//input[@value='Roman Catholic']")
christian_selection = browser.find_element_by_xpath("//input[@value='Christian']")
inc_selection = browser.find_element_by_xpath("//input[@value='INC']")
islam_selection = browser.find_element_by_xpath("//input[@value='Islam']")
jehovah_selection = browser.find_element_by_xpath("//input[@value='Jehovah']")
other_selection = browser.find_element_by_xpath("//input[@value='Others']")
otherReligion_field = browser.find_element_by_xpath("//input[@class='form-control text-box single-line']")
forgotBirthdate_selection = browser.find_element_by_id('called')

#test input
id_field.send_keys(id_no(1))