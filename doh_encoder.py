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

#load Workbook
wb = load_workbook('Database3.xlsx')
alphabet = list(string.ascii_uppercase)
#alphabet = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z']

#cell colors
greenFill = PatternFill(start_color='2ecc71',
                   end_color='2ecc71',
                   fill_type='solid')
redFill = PatternFill(start_color='e74c3c',
                   end_color='e74c3c',
                   fill_type='solid')



#load resources for a specific worksheet and user
class get_worksheet_data:
    def __init__(self,worksheet_no,user_number):
        #choose worksheet
        self.worksheet_no = worksheet_no
        self.user_number = user_number
        sheet = wb.worksheets[worksheet_no-1]
        self.sheet = sheet
        wb.active = sheet
        #get the number of people from page
        profiles_to_encode = sheet.max_row - 1 
        self.profiles_to_encode = profiles_to_encode
        #get the number of data from page
        data_to_encode = sheet.max_column 
        self.data_to_encode = data_to_encode
        #profile 1 profile 2 profile 3 etc..
        profile_list = list(range(1,profiles_to_encode+1))
        self.profile_list = profile_list
        data_list = list(range(1,data_to_encode+1))
        self.data_list = data_list
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
        self.profile_users = profile_users

    def get_all_data(self):
        for data in self.data_list:
            print(str(self.user_number-1) + str(data-1))
            if self.sheet[self.profile_users[self.user_number-1][data-1]].value:
                print(self.sheet[self.profile_users[self.user_number-1][data-1]].value) 
    def get_specific_data(self,index_no):
            if self.sheet[self.profile_users[self.user_number-1][index_no-1]].value:
                return self.sheet[self.profile_users[self.user_number-1][index_no-1]].value
    def set_to_ok(self,index_no):
            self.sheet[self.profile_users[self.user_number-1][index_no-1]].fill = greenFill
           # wb.save('Stylized.xlsx')
    def set_to_not_ok(self,index_no):
            self.sheet[self.profile_users[self.user_number-1][index_no-1]].fill = redFill
           # wb.save('Stylized.xlsx')



        
pageData1 = get_worksheet_data(1,1)
print (get_worksheet_data(1,1).get_specific_data(1))

#main loop 
# login user -> navigate to encoding page -> encode[ input on fields -> change cell color -> click save] 
for profile in pageData1.profile_list:
    
    #login
    browser = webdriver.Chrome('C:\\Users\\Cj\\Desktop\\Temp\\chromedriver')
    browser.get('http://122.54.82.254:65428/')
    user_field = browser.find_element_by_id('UserName')
    pass_field = browser.find_element_by_id('Password')
    login = browser.find_element_by_xpath("//input[@value='Log in']")
    user_input = input('Type User')
    pass_input = input('Type Pass')
    user_field.send_keys(user_input)
    pass_field.send_keys(pass_input)
    login.click()

    #navigate to vaccination entry page
    create_new_profile = browser.find_element_by_xpath("//button[@data-toggle='modal']")
    time.sleep(2)
    create_new_profile.click()
    add_vacinee_profile = browser.find_element_by_xpath("//a[@href='/VaccinneeProfile/Add']")
    time.sleep(2)
    add_vacinee_profile.click()
    time.sleep(2)

    # paths on General Information
    barangay_option = browser.find_element_by_xpath('//*[@id="content-main"]/div/div/div[2]/form/div/div/div[2]/table[3]/tbody/tr[1]/td[3]')
    id_field = browser.find_element_by_id('VaccinationCardId')
    date_interviewed_field = browser.find_element_by_id('DateInterviewed')
    relation_to_vaccined = browser.find_element_by_id('RelasyonSaBakunado')
    last_name_field = browser.find_element_by_id('LastName')
    first_name_field = browser.find_element_by_id('FirstName')
    middle_inintial_field = browser.find_element_by_id('MiddleName')
    extension_name_field = browser.find_element_by_id('ExtensionName')
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
    street_field = browser.find_element_by_id('Street')
    house_no_field = browser.find_element_by_id('LotNo')
    birthdate_field = browser.find_element_by_id('BirthDate')
    roman_catholic_selection = browser.find_element_by_xpath("//input[@value='Roman Catholic']")
    christian_selection = browser.find_element_by_xpath("//input[@value='Christian']")
    inc_selection = browser.find_element_by_xpath("//input[@value='INC']")
    islam_selection = browser.find_element_by_xpath("//input[@value='Islam']")
    jehovah_selection = browser.find_element_by_xpath("//input[@value='Jehovah']")
    other_selection = browser.find_element_by_xpath("//input[@value='Others']")
   # otherReligion_field = browser.find_element_by_xpath("//input[@class='form-control text-box single-line']")
    #paths to
# General Information functions
    def id_input(profile_no):
        if get_worksheet_data(1,profile_no).get_specific_data(1):
            id_field.send_keys(get_worksheet_data(1,profile_no).get_specific_data(1))
            get_worksheet_data(1,profile_no).set_to_ok(1)
        else:
            get_worksheet_data(1,profile_no).set_to_not_ok(1)
    def date_interviewed_input(profile_no):
        if get_worksheet_data(1,profile_no).get_specific_data(2):
            date_interviewed_field.send_keys(get_worksheet_data(1,profile_no).get_specific_data(2).strftime("%m/%d/%Y"))
            get_worksheet_data(1,profile_no).set_to_ok(2)
        else:
            get_worksheet_data(1,profile_no).set_to_not_ok(2)
    def relation_to_vaccined_input(profile_no):
        if get_worksheet_data(1,profile_no).get_specific_data(3):
            relation_to_vaccined.send_keys(get_worksheet_data(1,profile_no).get_specific_data(3))
            get_worksheet_data(1,profile_no).set_to_ok(3)
        else:
            get_worksheet_data(1,profile_no).set_to_not_ok(3)
    def last_name_field_input(profile_no):
        if get_worksheet_data(1,profile_no).get_specific_data(4):
            last_name_field.send_keys(get_worksheet_data(1,profile_no).get_specific_data(4))
            get_worksheet_data(1,profile_no).set_to_ok(4)
        else:
            get_worksheet_data(1,profile_no).set_to_not_ok(4)
    def first_name_field_input(profile_no):
        if get_worksheet_data(1,profile_no).get_specific_data(5):
            first_name_field.send_keys(get_worksheet_data(1,profile_no).get_specific_data(5))
            get_worksheet_data(1,profile_no).set_to_ok(5)
        else:
            get_worksheet_data(1,profile_no).set_to_not_ok(5)
    def middle_inintial_field_input(profile_no):
        if get_worksheet_data(1,profile_no).get_specific_data(6):
            middle_inintial_field.send_keys(get_worksheet_data(1,profile_no).get_specific_data(6))
            get_worksheet_data(1,profile_no).set_to_ok(6)
        else:
            get_worksheet_data(1,profile_no).set_to_not_ok(6)
    def extension_name_field_input(profile_no):
        if get_worksheet_data(1,profile_no).get_specific_data(7):
            extension_name_field.send_keys(get_worksheet_data(1,profile_no).get_specific_data(7))
            get_worksheet_data(1,profile_no).set_to_ok(7)
        else:
            get_worksheet_data(1,profile_no).set_to_not_ok(7)
    def relation_to_household_input(profile_no):
        if get_worksheet_data(1,profile_no).get_specific_data(8):
            relation_to_household.send_keys(get_worksheet_data(1,profile_no).get_specific_data(8))
            get_worksheet_data(1,profile_no).set_to_ok(8)
        else:
            get_worksheet_data(1,profile_no).set_to_not_ok(8)
    def respondent_input(profile_no):
        if get_worksheet_data(1,profile_no).get_specific_data(9):
            respondent.send_keys(get_worksheet_data(1,profile_no).get_specific_data(9))
            get_worksheet_data(1,profile_no).set_to_ok(9)
        else:
            get_worksheet_data(1,profile_no).set_to_not_ok(9)
    def contact_no_field_input(profile_no):
        if get_worksheet_data(1,profile_no).get_specific_data(10):
            contact_no_field.send_keys(("0")+str(get_worksheet_data(1,profile).get_specific_data(10)))
            get_worksheet_data(1,profile_no).set_to_ok(10)
        else:
            get_worksheet_data(1,profile_no).set_to_not_ok(10)
    def province_input(profile_no):
        if get_worksheet_data(1,profile_no).get_specific_data(11):
            province_option.click()
            time.sleep(1)
            province_field.send_keys(get_worksheet_data(1,profile_no).get_specific_data(11))
            selector_button = browser.find_element_by_xpath('//li[@class="select2-results-dept-0 select2-result select2-result-selectable select2-highlighted"]')
            selector_button.click()
            get_worksheet_data(1,profile_no).set_to_ok(11)
        else:
            get_worksheet_data(1,profile_no).set_to_not_ok(11)
    def city_input(profile_no):
        if get_worksheet_data(1,profile_no).get_specific_data(12):
            city_option.click()
            city_field = browser.find_element_by_id('s2id_autogen2_search')
            city_field.send_keys(get_worksheet_data(1,profile_no).get_specific_data(12))
            selector_button = browser.find_element_by_xpath('//li[@class="select2-results-dept-0 select2-result select2-result-selectable select2-highlighted"]')
            selector_button.click()
            get_worksheet_data(1,profile_no).set_to_ok(12)
        else:
            get_worksheet_data(1,profile_no).set_to_not_ok(12)
    def barangay_input(profile_no):
        if get_worksheet_data(1,profile_no).get_specific_data(13):
            barangay_option.click()
            barangay_field = browser.find_element_by_id('s2id_autogen3_search')
            barangay_field.send_keys(get_worksheet_data(1,profile_no).get_specific_data(13))
            selector_button = browser.find_element_by_xpath('//li[@class="select2-results-dept-0 select2-result select2-result-selectable select2-highlighted"]')
            selector_button.click()
            get_worksheet_data(1,profile_no).set_to_ok(13)
        else:
            get_worksheet_data(1,profile_no).set_to_not_ok(13)
    def street_field_input(profile_no):
        if get_worksheet_data(1,profile_no).get_specific_data(14):
            street_field.send_keys(get_worksheet_data(1,profile).get_specific_data(14))
            get_worksheet_data(1,profile_no).set_to_ok(14)
        else:
            get_worksheet_data(1,profile_no).set_to_not_ok(14)
    def house_no_field_input(profile_no):
        if get_worksheet_data(1,profile_no).get_specific_data(15):
            house_no_field.send_keys(get_worksheet_data(1,profile).get_specific_data(15))
            get_worksheet_data(1,profile_no).set_to_ok(15)
        else:
            get_worksheet_data(1,profile_no).set_to_not_ok(15)
    def sex_input(profile_no):
        if get_worksheet_data(1,profile_no).get_specific_data(16):
            if get_worksheet_data(1,profile).get_specific_data(16) == 'MALE':
                male_selection.click()
                get_worksheet_data(1,profile_no).set_to_ok(16)
            else:
                female_selection.click()
                get_worksheet_data(1,profile_no).set_to_ok(16)

        else:
            get_worksheet_data(1,profile_no).set_to_not_ok(16)
    def age_input(profile_no):
        if get_worksheet_data(1,profile_no).get_specific_data(17):
            age_field.send_keys(get_worksheet_data(1,profile).get_specific_data(17))
            get_worksheet_data(1,profile_no).set_to_ok(17)
        else:
            get_worksheet_data(1,profile_no).set_to_not_ok(17)
    def religion_input(profile_no):
        if get_worksheet_data(1,profile_no).get_specific_data(18):
            if get_worksheet_data(1,profile_no).get_specific_data(18) == "ROMAN CATHOLIC":
                roman_catholic_selection.click()
                get_worksheet_data(1,profile_no).set_to_ok(18)
            elif get_worksheet_data(1,profile_no).get_specific_data(18) == "CHRISTIAN":
                christian_selection.click()
                get_worksheet_data(1,profile_no).set_to_ok(18)
            elif get_worksheet_data(1,profile_no).get_specific_data(18) == "INC" :
                inc_selection.click()
                get_worksheet_data(1,profile_no).set_to_ok(18)
            elif get_worksheet_data(1,profile_no).get_specific_data(18) == "ISLAM" :
                islam_selection.click()
                get_worksheet_data(1,profile_no).set_to_ok(18)
            elif get_worksheet_data(1,profile_no).get_specific_data(18) == "JEHOVAH" :
                jehovah_selection.click()
                get_worksheet_data(1,profile_no).set_to_ok(18)
            else: 
                other_selection.click()
                otherReligion_field.send_keys(get_worksheet_data(1,1).get_specific_data(18))
                get_worksheet_data(1,profile_no).set_to_ok(18)
        else:
            get_worksheet_data(1,profile_no).set_to_not_ok(18)
    def birthdate_input(profile_no):
        if get_worksheet_data(1,1).get_specific_data(19):
            birthdate_field.send_keys((get_worksheet_data(1,profile_no).get_specific_data(19)).strftime("%m/%d/%Y"))
            get_worksheet_data(1,profile_no).set_to_ok(19)
        else: 
            forgotBirthdate_selection.click()
            get_worksheet_data(1,profile_no).set_to_ok(19)
    def birthplace_input(profile_no):
        if get_worksheet_data(1,profile_no).get_specific_data(20):
            birthplace_field.send_keys(get_worksheet_data(1,profile_no).get_specific_data(20))        
            get_worksheet_data(1,profile_no).set_to_ok(20)
        else: 
            get_worksheet_data(1,profile_no).set_to_not_ok(20)
    def years_at_curr_input(profile_no):
        if get_worksheet_data(1,profile_no).get_specific_data(21):
            years_at_curr.send_keys(get_worksheet_data(1,profile_no).get_specific_data(21))
            get_worksheet_data(1,profile_no).set_to_ok(21)
        else:
            get_worksheet_data(1,profile_no).set_to_not_ok(21)

    def input_to__general_information(profile):

        id_input(profile)
        date_interviewed_input(profile)
        relation_to_vaccined_input(profile)
        last_name_field_input(profile)
        first_name_field_input(profile)
        middle_inintial_field_input(profile)
        extension_name_field_input(profile)
        relation_to_household_input(profile)
        respondent_input(profile)
        contact_no_field_input(profile)
        province_input(profile)
        city_input(profile)
        barangay_input(profile)
        street_field_input(profile)
        house_no_field_input(profile)
        sex_input(profile)
        age_input(profile)
        religion_input(profile)
        birthdate_input(profile)
        birthplace_input(profile)
        years_at_curr_input(profile)
        print(submit_button)
        submit_button.click()
        wb.save('Stylized3.xlsx')
        input('wait')
    #new method for encoding data on General Info Page
    
    numlist =[ 3,4,5,6,7,8,9]
    numlist2 = [11,12,13]
    numlist3 = [14,15]
    numlist4 = [20,21]
    def input_test_data(profile):
        #select first element
        id_input(profile)
        current_element = browser.switch_to_active_element()
        current_element.send_keys(Keys.TAB)
        current_element = browser.switch_to_active_element()
        #datetime element
        if get_worksheet_data(1,profile).get_specific_data(2):
            current_element.send_keys(get_worksheet_data(1,profile).get_specific_data(2).strftime("%m/%d/%Y"))
            get_worksheet_data(1,profile).set_to_ok(2)
            current_element.send_keys(Keys.TAB)
        else: 
            get_worksheet_data(1,profile).get_specific_data(2).set_to_not_ok
            current_element.send_keys(Keys.TAB)
        for number in numlist:
            current_element = browser.switch_to_active_element()
            if get_worksheet_data(1,profile).get_specific_data(number):
                current_element.send_keys(get_worksheet_data(1,profile).get_specific_data(number))
                get_worksheet_data(1,profile).set_to_ok(number)
                current_element.send_keys(Keys.TAB)
            else: 
                get_worksheet_data(1,profile).set_to_not_ok(number)

                current_element.send_keys(Keys.TAB)

        if get_worksheet_data(1,profile).get_specific_data(10):
                current_element = browser.switch_to_active_element()    
                current_element.send_keys("0"+ str(get_worksheet_data(1,profile).get_specific_data(10)))
                get_worksheet_data(1,profile).set_to_ok(10)
                current_element.send_keys(Keys.TAB)
        else: 
            current_element = browser.switch_to_active_element()
            get_worksheet_data(1,profile).set_to_not_ok(10)
            current_element.send_keys(Keys.TAB)
        for number in numlist2:
            current_element = browser.switch_to_active_element()
            if get_worksheet_data(1,profile).get_specific_data(number):
                current_element.send_keys(Keys.ENTER)
                current_element = browser.switch_to_active_element()
                current_element.send_keys(get_worksheet_data(1,profile).get_specific_data(number))
                current_element.send_keys(Keys.ENTER)
                current_element = browser.switch_to_active_element()
                get_worksheet_data(1,profile).set_to_ok(number)
                current_element.send_keys(Keys.TAB)
            else:
                get_worksheet_data(1,profile).set_to_not_ok(number)
                current_element.send_keys(Keys.TAB)
        for number in numlist3:
            current_element = browser.switch_to_active_element()
            if get_worksheet_data(1,profile).get_specific_data(number):
                current_element.send_keys(get_worksheet_data(1,profile).get_specific_data(number))
                get_worksheet_data(1,profile).set_to_ok(number)
                current_element.send_keys(Keys.TAB)
            else: 
                get_worksheet_data(1,profile).set_to_not_ok(number)
                current_element.send_keys(Keys.TAB)         
        current_element = browser.switch_to_active_element()
        if get_worksheet_data(1,profile).get_specific_data(16) == "MALE":
          #  current_element.send_keys(Keys.ARROW_RIGHT)
           # current_element = browser.switch_to_active_element()
            current_element.send_keys(Keys.ARROW_LEFT)
            current_element = browser.switch_to_active_element()
            get_worksheet_data(1,profile).set_to_ok(16)
            current_element.send_keys(Keys.TAB)
        elif get_worksheet_data(1,profile).get_specific_data(16) == "FEMALE":
            current_element.send_keys(Keys.ARROW_RIGHT)
            current_element = browser.switch_to_active_element()
            current_element.send_keys(Keys.ARROW_RIGHT)
            current_element = browser.switch_to_active_element()
            get_worksheet_data(1,profile).set_to_ok(16)
            current_element.send_keys(Keys.TAB)
        else:
            get_worksheet_data(1,profile).set_to_not_ok(16)
            current_element.send_keys(Keys.TAB)
        if get_worksheet_data(1,profile).get_specific_data(17):
            current_element = browser.switch_to_active_element()
            current_element.send_keys(get_worksheet_data(1,profile).get_specific_data(17))
            get_worksheet_data(1,profile).set_to_ok(17)
            current_element.send_keys(Keys.TAB)
        else: 
            get_worksheet_data(1,profile).set_to_not_ok(number)
            current_element.send_keys(Keys.TAB)   
        #set religion
        if get_worksheet_data(1,profile).get_specific_data(18):
            if get_worksheet_data(1,profile).get_specific_data(18) == "ROMAN CATHOLIC":
                current_element = browser.switch_to_active_element()
                current_element.send_keys(Keys.ARROW_RIGHT)
                current_element = browser.switch_to_active_element()
                current_element.send_keys(Keys.ARROW_LEFT)
                current_element = browser.switch_to_active_element()
                current_element.send_keys(Keys.TAB)
                  
                get_worksheet_data(1,profile).set_to_ok(18)
            elif get_worksheet_data(1,profile).get_specific_data(18) == "CHRISTIAN":
                current_element = browser.switch_to_active_element()
                current_element.send_keys(Keys.ARROW_RIGHT)
                current_element = browser.switch_to_active_element()
                current_element.send_keys(Keys.TAB)
                  
                get_worksheet_data(1,profile).set_to_ok(18)
            elif get_worksheet_data(1,profile).get_specific_data(18) == "INC" :
                current_element = browser.switch_to_active_element()
                current_element.send_keys(Keys.ARROW_RIGHT)
                current_element = browser.switch_to_active_element()
                current_element.send_keys(Keys.ARROW_RIGHT)
                current_element = browser.switch_to_active_element()
                current_element.send_keys(Keys.TAB)
            
                get_worksheet_data(1,profile).set_to_ok(18)
            elif get_worksheet_data(1,profile).get_specific_data(18) == "ISLAM" :
                current_element = browser.switch_to_active_element()
                current_element.send_keys(Keys.ARROW_RIGHT)
                current_element = browser.switch_to_active_element()
                current_element.send_keys(Keys.ARROW_RIGHT)
                current_element = browser.switch_to_active_element()
                current_element.send_keys(Keys.ARROW_RIGHT)
                current_element = browser.switch_to_active_element()
                current_element.send_keys(Keys.TAB)
                
                get_worksheet_data(1,profile).set_to_ok(18)
            elif get_worksheet_data(1,profile).get_specific_data(18) == "JEHOVAH" :
                current_element = browser.switch_to_active_element()
                current_element.send_keys(Keys.ARROW_RIGHT)
                current_element = browser.switch_to_active_element()
                current_element.send_keys(Keys.ARROW_RIGHT)
                current_element = browser.switch_to_active_element()
                current_element.send_keys(Keys.ARROW_RIGHT)
                current_element = browser.switch_to_active_element()
                current_element.send_keys(Keys.ARROW_RIGHT)
                current_element = browser.switch_to_active_element()
                current_element.send_keys(Keys.TAB) 
                
                get_worksheet_data(1,profile).set_to_ok(18)
            else: 
                
                for number in [1,2,3,4,5]:
                    current_element = browser.switch_to_active_element()
                    current_element.send_keys(Keys.ARROW_RIGHT) 
                current_element.send_keys(Keys.TAB)
                current_element = browser.switch_to_active_element()
                current_element.send_keys(get_worksheet_data(1,1).get_specific_data(18))
                get_worksheet_data(1,profile).set_to_ok(18)
                current_element.send_keys(Keys.TAB)
        else:
            get_worksheet_data(1,profile).set_to_not_ok(18)
        
        if get_worksheet_data(1,profile).get_specific_data(19):
            current_element = browser.switch_to_active_element()
            current_element.send_keys(get_worksheet_data(1,profile).get_specific_data(19).strftime("%m/%d/%Y"))
            get_worksheet_data(1,profile).set_to_ok(19)
            current_element.send_keys(Keys.TAB)
        else: 
            get_worksheet_data(1,profile).get_specific_data(19).set_to_not_ok
           #if forgotBirthdate_selection:
                #forgotBirthdate_selection.click()
                #current_element = browser.switch_to_active_element()
                #current_element.send_keys(Keys.TAB)
           # else:
            get_worksheet_data(1,profile).get_specific_data(19).set_to_not_ok
            current_element = browser.switch_to_active_element()
            current_element.send_keys(Keys.TAB)
            
        for number in numlist4:
            current_element = browser.switch_to_active_element()
            if get_worksheet_data(1,profile).get_specific_data(number):
                current_element.send_keys(get_worksheet_data(1,profile).get_specific_data(number))
                get_worksheet_data(1,profile).set_to_ok(number)
                current_element.send_keys(Keys.TAB)
            else: 
                get_worksheet_data(1,profile).set_to_not_ok(number)
                current_element.send_keys(Keys.TAB)  

    #input_to__general_information(profile)
   
    input_test_data(profile)
    input('wait')