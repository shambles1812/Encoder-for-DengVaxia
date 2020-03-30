from tkinter import *
from selenium import webdriver

#browser = webdriver.Chrome('C:\\Users\\Cj\\Desktop\\Temp\\chromedriver')
def HelloWorld():
    print("Hellowworld")
    
class Login:
    def __init__(self):
        browser = webdriver.Chrome('C:\\Users\\Cj\\Desktop\\Temp\\chromedriver')
        browser.get('http://122.54.82.254:65428/')
        user_field = browser.find_element_by_id('UserName')
        pass_field = browser.find_element_by_id('Password')
        login = browser.find_element_by_xpath("//input[@value='Log in']")
        self.user_field = user_field
        self.pass_field = pass_field
        self.login = login
    def With_this(self,UserName,Password):
        self.user_field.send_keys(UserName)
        self.pass_field.send_keys(Password)
        self.login.click()

if __name__ == "__main__":
    master = Tk() 
    message = "Enter DOH login credentials"
    messageVar = Message(master, text = message)
    messageVar.grid(row=0)
    Label(master, text='First Name').grid(row=1) 
    Label(master, text='Last Name').grid(row=2) 
    e1 = Entry(master) 
    e2 = Entry(master) 
    e1.grid(row=1, column=2) 
    e2.grid(row=2, column=2) 
    button = Button(master, text='Login', width=25, command=lambda : Login().With_this(e1.get(),e2.get()))
    button.grid(row=3)
    mainloop() 
