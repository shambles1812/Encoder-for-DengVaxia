from selenium import webdriver
from selenium.webdriver.common.keys import Keys 
from openpyxl import load_workbook, Workbook
from datetime import datetime
import string
import time

#load Workbook
wb = load_workbook('Database.xlsx')
sheet = wb.worksheets[0]

alphabet = list(string.ascii_uppercase)
#get the number of people
profiles_to_encode = sheet.max_column - 1 
profile_list = list(range(1,profiles_to_encode))

#load resources
first_name = sheet['A2'].value
last_name = sheet['B2'].value
middle_initial = sheet['C2'].value
rel_to_household = sheet['D2'].value
resp = sheet['E2'].value
contact_no = sheet['F2'].value
province = sheet['G2'].value
city = sheet['H2'].value
Barangay = sheet['I2'].value
street = sheet['J2'].value
house_no = sheet['K2'].value
sex = sheet['L2'].value
age = sheet['M2'].value
religion = sheet['N2'].value
date_of_birth = sheet['O2'].value
birthplace = sheet['P2'].value
yrs = sheet['Q2'].value
id_no = sheet['R2'].value

for letter in alphabet:
    print(letter)

print(date_of_birth.strftime("%m/%d/%Y"))
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

#input to paths
id_field.send_keys(id_no)
last_name_field.send_keys(last_name)
first_name_field.send_keys(first_name)
middle_inintial_field.send_keys(middle_initial)
relation_to_household.send_keys(rel_to_household)
respondent.send_keys(resp)
contact_no_field.send_keys(contact_no)

#choose from dropdown
province_option.click()
province_field = browser.find_element_by_xpath('//input[@class="select2-input"]')
province_field.send_keys(province)
selector_button = browser.find_element_by_xpath('//li[@class="select2-results-dept-0 select2-result select2-result-selectable select2-highlighted"]')
selector_button.click()

city_option.click()
city_field = browser.find_element_by_id('s2id_autogen2_search')

city_field.send_keys(city)
selector_button = browser.find_element_by_xpath('//li[@class="select2-results-dept-0 select2-result select2-result-selectable select2-highlighted"]')
selector_button.click()

barangay_option.click()
barangay_field = browser.find_element_by_id('s2id_autogen3_search')
barangay_field.send_keys(Barangay)
selector_button = browser.find_element_by_xpath('//li[@class="select2-results-dept-0 select2-result select2-result-selectable select2-highlighted"]')
selector_button.click()

street_field.send_keys(street)
house_no_field.send_keys(house_no)
#sex selection
if sex == 'MALE':
    male_selection.click()
else:
    female_selection.click()
age_field.send_keys(age)
#religion 
if religion == "Roman Catholic":
    roman_catholic_selection.click()
elif religion == "Christian":
    christian_selection.click()
elif religion == "INC" :
    inc_selection.click()
elif religion == "Islam" :
    islam_selection.click()
elif religion == "Jehovah" :
    jehovah_selection.click()
else: 
        other_selection.click()
        otherReligion_field.send_keys(religion)

birthdate_field.send_keys(date_of_birth.strftime("%m/%d/%Y"))
birthplace_field.send_keys(birthplace)
years_at_curr.send_keys(yrs)
input('waiting for enter')

