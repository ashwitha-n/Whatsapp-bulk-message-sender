import xlrd
import time
import tkinter.filedialog as tk
import tkinter 
from tkinter import Button, Entry, Message, StringVar, Tk
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options

from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager

import os
import sys



wait=None
browser=None
Link="https://web.whatsapp.com/"
message="hello"
numbers=[]
docChoice="no"
choice="no"
doc_filename=""
image_filename=""



def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(__file__)
    return os.path.join(base_path, relative_path)


def whatsapp_login( ):
    global wait, browser, Link
    chrome_options = Options()
    chrome_options.add_argument('--user-data-dir=./User_Data')
    
    #browser = webdriver.Chrome(resource_path('C:/Users/Ashwitha/Desktop/MyProj/chromedriver.exe'),options=chrome_options)
    #browser=webdriver.Chrome(executable_path='chromedriver.exe',
    #                         options=chrome_options)
    #browser = webdriver.Chrome(ChromeDriverManager().install())
    browser = webdriver.Chrome(executable_path='chromedriver.exe', options=chrome_options)
    wait = WebDriverWait(browser, 600)
    browser.get(Link)
    browser.maximize_window()
    time.sleep(10)
    print("QR scanned")

def sender():
    global choice, docChoice, numbers
    time.sleep(10)
    if len(numbers) > 0:
        for i in numbers:
            
            link = "https://web.whatsapp.com/send?phone={}&text&source&data&app_absent".format(i)
            #driver  = webdriver.Chrome()
            browser.get(link)
            print("Sending message to", i)
            time.sleep(10)
            send_unsaved_contact_message()
            if(choice == "yes"):
                try:
                    send_attachment()
                except:
                    print('Attachment not sent.')
            if(docChoice == "yes"):
                try:
                    send_files()
                except:
                    print('Files not sent')
            time.sleep(7)
def send_unsaved_contact_message():
    global message
    print("send_unsaved_contact_message called")
    try:
        time.sleep(7)
        input_box = browser.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div[2]/div/div[2]')
        for ch in message:
            if ch == "\n":
                ActionChains(browser).key_down(Keys.SHIFT).key_down(Keys.ENTER).key_up(Keys.ENTER).key_up(Keys.SHIFT).key_up(Keys.BACKSPACE).perform()
            else:
                input_box.send_keys(ch)
        input_box.send_keys(Keys.ENTER)
        print("Message sent successfuly")
    except NoSuchElementException:
        print("Failed to send message")
        return


def send_attachment():
    global image_filename
    # Attachment Drop Down Menu
    print("send_attachment called")
    clipButton = browser.find_element_by_xpath('//*[@id="main"]/header/div[3]/div/div[2]/div/span')
    clipButton.click()
    time.sleep(1)

    # To send Videos and Images.
    mediaButton = browser.find_element_by_xpath('//*[@id="main"]/header/div[3]/div/div[2]/span/div/div/ul/li[1]/button/input')

    time.sleep(3)
    mediaButton.send_keys(image_filename)
    
    image_path=image_filename
    print(image_path)

    time.sleep(3)
    whatsapp_send_button=browser.find_element_by_xpath('//*[@id="app"]/div/div/div[2]/div[2]/span/div/span/div/div/div[2]/span/div/div/span')
    whatsapp_send_button.click()


def send_files():
    global doc_filename
    # Attachment Drop Down Menu
    clipButton = browser.find_element_by_xpath('//*[@id="main"]/header/div[3]/div/div[2]/div/span')
    clipButton.click()
    time.sleep(1)

    # To send a Document(PDF, Word file, PPT)
    docButton = browser.find_element_by_xpath('//*[@id="main"]/header/div[3]/div/div[2]/span/div/div/ul/li[3]/button/input')
    docButton.send_keys(doc_filename)
    time.sleep(1)

    time.sleep(3)
    whatsapp_send_button = browser.find_element_by_xpath('//*[@id="app"]/div/div/div[2]/div[2]/span/div/span/div/div/div[2]/span/div/div/span')
    whatsapp_send_button.click()

def buttonClick(tk):
    global numbers
    filename =tk.askopenfilename()
    excel_file=filename
    numbe=xlrd.open_workbook(excel_file)
    sheet = numbe.sheet_by_index(0) 
    sheet.cell_value(0, 0) 
        
    for i in range(sheet.nrows): 
        nu=910000000000
        numbers.append(int(sheet.cell_value(i, 0))+nu) 
    #print (numbers)
def inputmsg():
    global message
    message=name.get()
    
def inputimage():
    global image_filename,choice
    choice="yes"
    
    image_filename2=tk.askopenfilename()
    
    for i in image_filename2:
        if(i=='/'):
            image_filename=image_filename+'\\'
        else:
            image_filename=image_filename+i

def inputattachement():
    global doc_filename,docChoice
    docChoice="yes"
    doc_filename2=tk.askopenfilename()
    for i in doc_filename2:
        if(i=='/'):
            doc_filename= doc_filename+'\\'
        else:
            doc_filename=doc_filename+i


def submit():
    whatsapp_login()
    sender()
    print("done")

## ui part
tkWindow = Tk()  
tkWindow.geometry('400x150')  
tkWindow.title('whatsup')
button=Button(tkWindow,text='upload excel file for phone numbers',command=lambda: buttonClick(tk))
name = tk.StringVar()
nameEntered = Entry(tkWindow, width = 15, textvariable = name)
nameEntered.grid(column = 0, row = 2)
 
button2=Button(tkWindow,text='submit text',command=lambda: inputmsg())
button3=Button(tkWindow,text='upload image', command=lambda: inputimage())
button4=Button(tkWindow,text='upload attachement', command=lambda: inputattachement())
button.grid(column=0,row=1)
button2.grid(column=0,row=3)
button3.grid(column=0,row=4)
button4.grid(column=0,row=5)
button5=Button(tkWindow,text='submit',command=lambda: submit())
button5.grid(column=0,row=6)
tkWindow.mainloop()

print(message)
print(doc_filename)
print(image_filename)

