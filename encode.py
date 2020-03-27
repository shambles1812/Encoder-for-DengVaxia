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

def id_no(user_number):
    if sheet[profile_users[user_number-1][0]].value:
        return sheet[profile_users[user_number-1][0]].value
def date_interviewed(user_number):
    if sheet[profile_users[user_number-1][0]].value:
        return sheet[profile_users[user_number-1][0]].value
def first_name(user_number):
    if sheet[profile_users[user_number-1][17]].value:
        return sheet[profile_users[user_number-1][17]].value
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


#profile 1 profile 2 profile 3 etc..
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
