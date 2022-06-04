from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import sys
from tkinter import *
from tkinter import messagebox
import tkinter.simpledialog
import os
from selenium.common.exceptions import *
import winsound
import win32com.client as wincl

tk=Tk()
tk.title("Tibia Online Checker")
tk.iconbitmap('main.ico')
tk.geometry("300x300")
tk.configure(bg="#6989AA")
tkname = StringVar()
tkcooldown = IntVar()

def login(browser):
    browser.get("https://www.tibia.com/community/?subtopic=characters")
    print('Go to community')
    time.sleep(6)

def check_status(browser,name):
    account = browser.find_element_by_xpath('//*[@id="characters"]/div[5]/div/div/form/table/tbody/tr[2]/td/table/tbody/tr/td[2]/input')
    account.send_keys(name)
    submit = browser.find_element_by_xpath('//*[@id="characters"]/div[5]/div/div/form/table/tbody/tr[2]/td/table/tbody/tr/td[3]/input')
    submit.click()
    time.sleep(3)

def load_data():
    #LOAD SETTINGS FROM SAVE.TXT
    if os.path.isfile('save.txt'):
        with open('save.txt','r') as f:
            all_lines = f.read().split(',')
            #all_lines = f.readlines()
            tkname.set(all_lines[0])
            tkcooldown.set(all_lines[1])

def save_data():
    name = tkname.get()
    cooldown = tkcooldown.get()
    #Error message
    if(tkname==''):
        print('Error')
        messagebox.showerror('error','somewhere is missing requied value...')

    elif(tkcooldown==''):
        print('Error')
        messagebox.showerror('error','somewhere is missing requied value...')
    else:
        result = messagebox.askquestion("Submit","Are you sure you want to save the data?")
        if(result=='yes'):
            print('You have succesfully save your settings!')
        #SAVE LIST TO TXT
            pList = [name,cooldown]
            with open("save.txt", 'a') as f:
                f.truncate(0)
                f.write(','.join(map(str,pList)))

def main():
    speak = wincl.Dispatch("SAPI.SpVoice")
    options = webdriver.ChromeOptions()
    options.add_argument('--lang=en')
    options.add_argument('--headless')
    options.add_argument('--disable-gpu')
    browser = webdriver.Chrome(executable_path=r"chromedriver.exe",chrome_options=options)
    name = tkname.get()
    character = name
    cooldown = 60*tkcooldown.get()
    tk.destroy()
    login(browser)
    check_status(browser,name)
    while True:
        try:
            browser.refresh()
            Online = Online_active.get()
            Offline = Offline_active.get()
            Eror = Error_active.get()
            green = browser.find_element_by_css_selector('#characters > div.Border_2 > div > div > table:nth-child(16) > tbody > tr:nth-child(3) > td:nth-child(3) > b')
            if green:
                print(name,' - Status Online ')
                if Online == 1 & Eror == 1:
                    winsound.PlaySound("error.wav", winsound.SND_ASYNC)
                    time.sleep(1)
                    speak.Speak("Character is Online")
                elif Online == 1:
                    speak.Speak("Character is Online")
                    time.sleep(1)
                elif Eror == 1:
                    winsound.PlaySound("error.wav", winsound.SND_ASYNC)
                    time.sleep(1)
                else:
                    pass
        except:
            print(name,' - Status Offline')
            if Offline == 1 & Eror == 1:
                    winsound.PlaySound("error.wav", winsound.SND_ASYNC)
                    time.sleep(1)
                    speak.Speak("Character is Offline")
            elif Offline == 1:
                speak.Speak("Character is Offline")
                time.sleep(1)
            elif Eror == 1:
                winsound.PlaySound("error.wav", winsound.SND_ASYNC)
                time.sleep(1)
            else:
                pass
        time.sleep(cooldown)
#-----------TKINTER LABELS--------------------------#
name_label = Label(tk,text="Account name:",bg="#6989AA")
name_label.grid(column=0, row=0)
entry_name = Entry(tk,textvariable=tkname)
entry_name.grid(column=1,row=0)
    #COOLDOWN
cooldown_label = Label(tk,text="Wait time:",bg="#6989AA")
cooldown_label.grid(column=0, row=1)
entry_cooldown = Entry(tk,textvariable=tkcooldown)
entry_cooldown.grid(column=1,row=1)
cooldown_label = Label(tk,text="min",bg="#6989AA")
cooldown_label.grid(column=2, row=1)
    #SAVE BUTTON
save_button= Button(tk,text="Save", bg="#0099FF", command=save_data)
save_button.grid(column=2, row=5)
    #START BUTTON
start_button= Button(tk,text="Start", bg="#89FCB6", command=main)
start_button.grid(column=1, row=5)
    #LOAD BUTTON
load_button= Button(tk,text="Load", bg="#F9C900", command=load_data)
load_button.grid(column=0, row=5)
#---CHECKBOXES---#
name_label = Label(tk,text="Voice on status:",bg="#6989AA")
name_label.grid(column=0, row=2)
Online_active = IntVar()
Checkbutton(tk, text="Online", bg="#6989AA", variable=Online_active).grid(column=1,row=2, sticky=W)
Offline_active = IntVar(value=1)
Checkbutton(tk, text="Offline",bg="#6989AA", variable=Offline_active).grid(column=1,row=3, sticky=W)
Error_active = IntVar()
Checkbutton(tk, text="Error Sound",bg="#6989AA", variable=Error_active).grid(column=1,row=4, sticky=W)

tk.mainloop()