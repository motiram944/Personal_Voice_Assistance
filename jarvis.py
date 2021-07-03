import pyttsx3 #pip install pyttsx3
import speech_recognition as sr #pip install speechRecognition
import datetime
import wikipedia #pip install wikipedia
import webbrowser
import os
import smtplib
import subprocess 
import wolframalpha 
import pyttsx3 
import tkinter 
import json 
import random 
import operator 
import speech_recognition as sr 
import datetime 
import wikipedia 
import webbrowser 
import os 
import file_search
import psutil
import pywinauto
import pyautogui
import time
import math
import time
import winshell 
import pyjokes 
import feedparser 
import smtplib 
import ctypes 
import time 
import requests 
import shutil 
import osascript
import keyboard
import pyautogui
import json
from twilio.rest import Client 
from clint.textui import progress 
from selenium import webdriver
from bs4 import BeautifulSoup 
import win32com.client as wincl 
from urllib.request import urlopen 
from googletrans import Translator
from math import pi
from pywinauto.application import Application

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
# print(voices[1].id)
engine.setProperty('voice', voices[0].id)

def speak(audio):
    engine.say(audio)
    engine.runAndWait()

def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour>=0 and hour<12:
        speak("Good Morning!")

    elif hour>=12 and hour<18:
        speak("Good Afternoon!")   

    else:
        speak("Good Evening!")  

    speak("I am Luci Sir. Please tell me how may I help you")

def batte():
    battery = psutil.sensors_battery()
    plugged = battery.power_plugged
    global percent
    percent=str(battery.percent)
    p=int (percent)
    if plugged == False:
        plugged="not plugged in"
        if p<26:
            speak("Sir battery percentage is"+percent+"percent please connect charger")
    else:
        plugged="Plugged in"

    
    

def takeCommand():
    #It takes microphone input from the user and returns string output

    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 0.5
        audio = r.listen(source)

    try:
        print("Recognizing...")    
        query = r.recognize_google(audio, language='en-in')
        print(f"User said: {query}\n")

    except Exception as e:
        # print(e)    
        print("Say that again please...")  
        return "None"
    return query

def sendEmail(to, content):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    server.login('youremail@gmail.com', 'your-password')
    server.sendmail('youremail@gmail.com', to, content)
    server.close()

def ans(query):
    myjson=open('info.json','r')
    data=myjson.read()

    b1=json.loads(data)
    d=str(b1[query])
    speak(d)


if __name__ == '__main__': 
    clear = lambda: os.system('cls') 
    
    # This Function will clean any 
    # command before execution of this python file 
    clear() 
    wishMe() 
    batte()
    
    
    while True: 
        
        query = takeCommand().lower() 
        
        # All the commands said by user will be 
        # stored here in 'query' and will be 
        # converted to lower case for easily 
        # recognition of command 
        if 'wikipedia' in query: 
            speak('Searching Wikipedia...') 
            query = query.replace("wikipedia", "") 
            results = wikipedia.summary(query, sentences = 3) 
            speak("According to Wikipedia") 
            print(results) 
            speak(results) 

        elif 'open youtube' in query: 
            speak("Here you go to Youtube\n") 
            webbrowser.register('chrome',None,webbrowser.BackgroundBrowser("C://Program Files (x86)//Google//Chrome//Application//chrome.exe"))
            webbrowser.get('chrome').open("youtube.com") 

        elif 'open google' in query: 
            speak("Here you go to Google\n") 
            webbrowser.register('chrome',None,webbrowser.BackgroundBrowser("C://Program Files (x86)//Google//Chrome//Application//chrome.exe"))
            webbrowser.get('chrome').open("google.com") 

        elif 'open stackoverflow' in query: 
            speak("Here you go to Stack Over flow.Happy coding") 
            webbrowser.register('chrome',None,webbrowser.BackgroundBrowser("C://Program Files (x86)//Google//Chrome//Application//chrome.exe"))
            webbrowser.get('chrome').open("stackoverflow.com") 

    

        elif 'play music' in query or "play song" in query: 
            speak("Here you go with music") 
            # music_dir = "G:\\Song" 
            music_dir = 'E:\m'
            songs = os.listdir(music_dir) 
            print(songs)     
            os.startfile(os.path.join(music_dir, songs[1])) 
        
        elif 'the time' in query: 
            strTime = datetime.datetime.now().strftime("%H:%M:%S")   
            speak(f"Sir, the time is {strTime}") 

        elif 'open chrome' in query: 
            codePath = r"C://Program Files (x86)//Google//Chrome//Application//chrome.exe"
            os.startfile(codePath) 

        elif 'email to gaurav' in query: 
            try: 
                speak("What should I say?") 
                content = takeCommand() 
                to = "Receiver email address"   
                sendEmail(to, content) 
                speak("Email has been sent !") 
            except Exception as e: 
                print(e) 
                speak("I am not able to send this email") 

        elif 'send a mail' in query: 
            try: 
                speak("What should I say?") 
                content = takeCommand() 
                speak("whome should i send") 
                to = input()     
                sendEmail(to, content) 
                speak("Email has been sent !") 
            except Exception as e: 
                print(e) 
                speak("I am not able to send this email") 

        elif 'how are you' in query: 
            speak("I am fine, Thank you") 
            speak("How are you, Sir") 

        elif 'fine' in query or "good" in query: 
            speak("It's good to know that your fine") 

        elif "change my name to" in query: 
            query = query.replace("change my name to", "") 
            assname = query 

        elif "change name" in query: 
            speak("What would you like to call me, Sir ") 
            assname = takeCommand() 
            speak("Thanks for naming me") 

        elif "what's your name" in query or "What is your name" in query: 
            speak("My friends call me") 
            speak(assname) 
            print("My friends call me", assname) 

        elif 'exit' in query: 
            speak("Thanks for giving me your time") 
            exit() 

        elif "who made you" in query or "who created you" in query: 
            speak("I have been created by Motiram sir.") 
            
        elif 'joke' in query: 
            speak(pyjokes.get_joke()) 

        elif 'search' in query: 
            query = query.replace("search", "")
            query = query +".com"        
               
            
            webbrowser.register('chrome',None,
            webbrowser.BackgroundBrowser("C://Program Files (x86)//Google//Chrome//Application//chrome.exe"))
            webbrowser.get('chrome').open(query)
            
        elif "calculate" in query.lower(): 
            
            app_id = "X87J88-PQPE398GKR"
            client = wolframalpha.Client(app_id) 
            indx = query.lower().split().index('calculate') 
            query = query.split()[indx + 1:] 
            res = client.query(' '.join(query)) 
            answer = next(res.results).text 
            print("The answer is " + answer) 
            speak("The answer is " + answer) 
            

        elif 'play' in query: 
            query = query.replace("play", "")         
            
            
            webbrowser.register('chrome',None,
            webbrowser.BackgroundBrowser("C://Program Files (x86)//Google//Chrome//Application//chrome.exe"))
            driver = webdriver.Chrome(executable_path = r'C:\Users\Lenovo\Downloads\chromedriver_win32 (1)\chromedriver.exe')
            driver.get('https://gaana.com')

            sb = driver.find_element_by_xpath('//*[@id="sb"]')
            sb.send_keys(query)

            b1 = driver.find_element_by_xpath('//*[@id="mainarea"]/div[1]/div[2]/div[1]/div[1]/a')
            b1.click()
            time.sleep(5)
            b2 = driver.find_element_by_xpath('//*[@id="new-release-album"]/li[1]/div/div[1]/a/img')
            b2.click()

        elif 'youtube' in query:
            query = query.replace("youtube", "")
            webbrowser.register('chrome',None,
            webbrowser.BackgroundBrowser("C://Program Files (x86)//Google//Chrome//Application//chrome.exe"))
            driver = webdriver.Chrome(executable_path = r'C:\Users\Lenovo\Downloads\chromedriver_win32 (1)\chromedriver.exe')
            driver.get('https://youtube.com')

            searchbox = driver.find_element_by_xpath('//*[@id="container"]')
            searchbox.click()
            pyautogui.write(query)
            time.sleep(2)
            searchButton = driver.find_element_by_xpath('//*[@id="search-icon-legacy"]')
            searchButton.click()

        elif 'shopping' in query:
            query = query.replace("shopping", "")   
            webbrowser.register('chrome',None,
            webbrowser.BackgroundBrowser("C://Program Files (x86)//Google//Chrome//Application//chrome.exe"))
            driver = webdriver.Chrome(executable_path = r'C:\Users\Lenovo\Downloads\chromedriver_win32 (1)\chromedriver.exe')
            driver.get('https://flipkart.com')

            box = driver.find_element_by_xpath('//*[@id="container"]/div/div[1]/div[1]/div[2]/div[2]/form/div/div/input')
            box.send_keys(query)

            gu = driver.find_element_by_xpath('/html/body/div[2]/div/div/button')
            gu.click()
            Bu =driver.find_element_by_xpath('//*[@id="container"]/div/div[1]/div[1]/div[2]/div[2]/form/div/button')
            Bu.click()

        elif 'program for' in query:
            query = query.replace("program", "")  
            query = "python program for" + query + "geeksforgeeks"
            webbrowser.register('chrome',None,
            webbrowser.BackgroundBrowser("C://Program Files (x86)//Google//Chrome//Application//chrome.exe"))
            driver = webdriver.Chrome(executable_path = r'C:\Users\Lenovo\Downloads\chromedriver_win32 (1)\chromedriver.exe')
            driver.get('https://google.com')

            box = driver.find_element_by_xpath('//*[@id="tsf"]/div[2]/div[1]/div[1]/div/div[2]/input')
            box.send_keys(query)

            gu = driver.find_element_by_xpath('//*[@id="tsf"]/div[2]/div[1]/div[3]/center/input[1]')
            gu.click()
            b3 = driver.find_element_by_xpath('//*[@id="rso"]/div[1]/div/div[1]/a/h3')
            b3.click()
            b4 = driver.find_element_by_xpath('//*[@id="highlighter_664614"]/table/tbody/tr/td/div')
            b4.click()

            
        elif 'copy program of' in query:
            query = query.replace("copy program of", "")  
            speak("In which language sir")
            note1 = takeCommand()
            query = note1 +"program for" + query + "geeksforgeeks"
            webbrowser.register('chrome',None,
            webbrowser.BackgroundBrowser("C://Program Files (x86)//Google//Chrome//Application//chrome.exe"))
            driver = webdriver.Chrome(executable_path = r'C:\Users\Lenovo\Downloads\chromedriver_win32 (1)\chromedriver.exe')
            driver.get('https://google.com')

            box = driver.find_element_by_xpath('/html/body/div[1]/div[3]/form/div[2]/div[1]/div[1]/div/div[2]/input')
            box.send_keys(query)

            gu = driver.find_element_by_xpath('/html/body/div[1]/div[3]/form/div[2]/div[1]/div[3]/center/input[1]')
            gu.click()
            b5 = driver.find_element_by_xpath('//*[@id="rso"]/div[1]/div/div[1]/a/h3/span')
            b5.click()
            b7 =driver.find_element_by_xpath('//*[@id="highlighter_496100"]/table/tbody/tr/td/div')
            b7.click()
            b6 = driver.find_element_by_xpath('//*[@id="copy-code-button"]')
            b6.click()
            

        elif 'open facebook' in query:
            webbrowser.register('chrome',None,
            webbrowser.BackgroundBrowser("C://Program Files (x86)//Google//Chrome//Application//chrome.exe"))
            driver = webdriver.Chrome(executable_path = r'C:\Users\Lenovo\Downloads\chromedriver_win32 (1)\chromedriver.exe')
            driver.get('https://facebook.com')

            log = driver.find_element_by_xpath('//*[@id="email"]')
            pyautogui.write("8975303848")
            pass1 = driver.find_element_by_xpath('//*[@id="pass"]')
            pass1.click()
            pyautogui.write("9922042916")

            b8 =driver.find_element_by_xpath('//*[@id="loginbutton"]')
            b8.click()
            speak("You want to check friends request sir")
            note2 = takeCommand()
            if note2=="yes":
                fr = driver.find_element_by_xpath('//*[@id="js_o"]/div')
                fr.click()
            elif note2=="no":
                speak("Ok sir")

        elif "who i am" in query: 
            speak("If you talk then definately your human.") 

        elif "why you came to world" in query: 
            speak("Thanks to motiram. further It's a secret") 

        elif 'power point presentation' in query: 
            speak("opening Power Point presentation") 
            power = r"C:\\Users\\GAURAV\\Desktop\\Minor Project\\Presentation\\Voice Assistant.pptx"
            os.startfile(power) 

        elif 'is love' in query: 
            speak("It is 7th sense that destroy all other senses") 

        elif "who are you" in query: 
            speak("I am your virtual assistant created by Gaurav") 

        elif 'reason for you' in query: 
            speak("I was created as a Minor project by Mister Gaurav ") 

        elif 'change background' in query: 
            ctypes.windll.user32.SystemParametersInfoW(20, 
                                                    0, 
                                                    "Location of wallpaper", 
                                                    0) 
            speak("Background changed succesfully") 

        elif 'open bluestack' in query: 
            appli = r"C:\\ProgramData\\BlueStacks\\Client\\Bluestacks.exe"
            os.startfile(appli) 

        elif 'news' in query: 
            
            try: 
                jsonObj = urlopen("http://newsapi.org/v1/articles?source=the-times-of-india&sortBy=top&apiKey=d50344cc650f47e4b89a245be1657b5a") 
                data = json.load(jsonObj) 
                i = 1
                
                speak('here are some top news from the times of india') 
                print('''=============== TIMES OF INDIA ============'''+ '\n') 
                
                for item in data['articles']: 
                    
                    print(str(i) + '. ' + item['title'] + '\n') 
                    print(item['description'] + '\n') 
                    speak(str(i) + '. ' + item['title'] + '\n') 
                    i += 1
            except Exception as e: 
                
                print(str(e)) 

        
        elif 'lock window' in query: 
                speak("locking the device") 
                ctypes.windll.user32.LockWorkStation() 

        elif 'shutdown system' in query: 
                speak("Hold On a Sec ! Your system is on its way to shut down") 
                subprocess.call('shutdown / p /f') 
                
        elif 'empty recycle bin' in query: 
            winshell.recycle_bin().empty(confirm = False, show_progress = False, sound = True) 
            speak("Recycle Bin Recycled") 

        elif "don't listen" in query or "stop listening" in query: 
            speak("for how much time you want to stop jarvis from listening commands") 
            a = int(takeCommand()) 
            time.sleep(a) 
            print(a) 

        elif "where is" in query: 
            query = query.replace("where is", "") 
            location = query 
            speak("User asked to Locate") 
            speak(location) 
            webbrowser.BackgroundBrowser("C://Program Files (x86)//Google//Chrome//Application//chrome.exe")
            webbrowser.get('chrome').open("https://www.google.nl / maps / place/" + location + "") 

        elif "camera" in query or "take a photo" in query: 
            ec.capture(0, "Jarvis Camera ", "img.jpg") 

        elif "restart" in query: 
            subprocess.call(["shutdown", "/r"]) 
            
        elif "hibernate" in query or "sleep" in query: 
            speak("Hibernating") 
            subprocess.call("shutdown / h") 

        elif "log off" in query or "sign out" in query: 
            speak("Make sure all the application are closed before sign-out") 
            time.sleep(5) 
            subprocess.call(["shutdown", "/l"]) 

        elif "write a note" in query: 
            speak("What should i write, sir") 
            note = takeCommand() 
            file = open('jarvis.txt', 'w') 
            speak("Sir, Should i include date and time") 
            snfm = takeCommand() 
            if 'yes' in snfm or 'sure' in snfm: 
                strTime = datetime.datetime.now().strftime("% H:% M:% S") 
                file.write(strTime) 
                file.write(" :- ") 
                file.write(note) 
            else: 
                file.write(note) 
        
        elif "show note" in query: 
            speak("Showing Notes") 
            file = open("jarvis.txt", "r") 
            print(file.read()) 
            speak(file.read(6)) 

        elif "update assistant" in query: 
            speak("After downloading file please replace this file with the downloaded one") 
            url = '# url after uploading file'
            r = requests.get(url, stream = True) 
            
            with open("Voice.py", "wb") as Pypdf: 
                
                total_length = int(r.headers.get('content-length')) 
                
                for ch in progress.bar(r.iter_content(chunk_size = 2391975), 
                                    expected_size =(total_length / 1024) + 1): 
                    if ch: 
                        Pypdf.write(ch) 
                    
        # NPPR9-FWDCX-D2C8J-H872K-2YT43 
        elif "jarvis" in query: 
            
            wishMe() 
            speak("Jarvis 1 point o in your service Mister") 
            speak(assname) 

        elif "weather" in query: 
            
            # Google Open weather website 
            # to get API of Open weather 
            api_key = "4249ae8e9074b64b7b7cf2163a9006df"
            base_url = "http://api.openweathermap.org/data/2.5/weather?"
            speak(" City name ") 
            print("City name : ") 
            city_name = takeCommand() 
            complete_url = base_url + "q =" + city_name + "&appid=" +api_key 
            response = requests.get(complete_url) 
            if response.status_code == 200:
                data = response.json()
                main = data['main']
                temperature = main['temp']
                humidity = main['humididty']
                pressure = main['pressure']
                report = data['weather']
                print(f"{city_name:-^30}")
                print(f"Temperature: {temperature}")
                print(f"Humidity: {humidity}")
                print(f"Pressure: {pressure}")
                print(f"Weather Report: {report[0]['description']}")
            else:
                print("Error in the HTTP request")
            
        elif "send message " in query: 
                # You need to create an account on Twilio to use this service 
                account_sid = 'Account Sid key'
                auth_token = 'Auth token'
                client = Client(account_sid, auth_token) 

                message = client.messages\
                                .create( 
                                    body = takeCommand(), 
                                    from_='Sender No', 
                                    to ='Receiver No'
                                ) 

                print(message.sid) 

        elif "battery percentage" in query:
            batte()
            speak("sir"+percent+"percent battery is avilable")

        elif 'find' in query:
            query = query.replace("find", "")
            file_search.set_root('E:')
            files = file_search.searchFile(query)
            speak("Here are some file location")
            for file in files:
                print(file)

        elif "volume up" in query: 
            pyautogui.press('volumeup')

        elif "translate" in query:
            query = query.replace("Translate", "")
            translator=Translator()
            speak("In which language sir")
            note3 = takeCommand()
            if note3=="Hindi":
                tra_sen=translator.translate(query,src='en',dest='hi')
                print(tra_sen.text)
            
            elif note3=="Marathi":
                tra_sen=translator.translate(query,src='en',dest='mr')
                print(tra_sen.text)
            
            elif note3=="English":
                tra_sen=translator.translate(query,src='mr',dest='en')
                print(tra_sen.text)
            

            
        
        elif "take screenshot" in query:
            time.sleep(5)
            img=pyautogui.screenshot()
            img.save("E:\screenshot\screenshot.png")

        elif "type" in query: 
            speak("What should i write, sir") 
            note = takeCommand() 
            pyautogui.write(note)

        elif "volume down" in query: 
            pyautogui.press('volumedown')

        elif "mute" in query: 
            pyautogui.press('volumemute')

        elif "wikipedia" in query: 
            webbrowser.open("wikipedia.com") 

        elif "scroll down" in query:
            pyautogui.scroll(-1500)

        elif "scroll up" in query:
            pyautogui.scroll(1500)

        elif "scroll right" in query:
            pyautogui.hscroll(-1000)

        elif "scroll left" in query:
            pyautogui.hscroll(-1000)

        elif "right click" in query:
            pyautogui.click(button='right')

        elif "switch window" in query:
            pyautogui.keyDown('alt')
            pyautogui.press('tab')
            pyautogui.keyUp('alt')

        elif "page up" in query:
            pyautogui.press('pageup')

        elif "page down" in query:
            pyautogui.press('pagedown')

        elif "numlock" in query:
            pyautogui.press('numlock')

        elif "backspace" in query:
            pyautogui.press('backspace')

        elif "inter" in query:
            pyautogui.press('enter')

        elif "escape" in query:
            pyautogui.press('esc')

        elif "browsersearch" in query:
            pyautogui.press('browsersearch')

        elif "refresh" in query:
            pyautogui.press('browserrefresh')

        elif "capslock" in query:
            pyautogui.press('capslock')

        elif "copy" in query:
            pyautogui.hotkey('ctrl','c')

        elif "paste" in query:
            pyautogui.hotkey('ctrl','v')

        elif "pause" in query or "play" in query:
            pyautogui.press('playpause')

        elif "previous" in query:
            pyautogui.press('prevtrack')

        elif "next" in query:
            pyautogui.press('nexttrack')

        elif "home" in query:
            pyautogui.press('home')

        elif "end" in query:
            pyautogui.press('end')

        elif "Good Morning" in query: 
            speak("A warm" +query) 
            speak("How are you Sir") 
            speak(assname) 

        # most asked question from google Assistant 
        elif "will you be my gf" in query or "will you be my bf" in query: 
            speak("I'm not sure about, may be you should give me some time") 

        elif "how are you" in query: 
            speak("I'm fine, glad you me that") 

        elif "i love you" in query: 
            speak("It's hard to understand") 
        
        elif "first name" in query:

            ans(query)
            

        elif "paint" in query:
            def start_paint():
                app = Application().start("mspaint.exe")
                
            def move_mouse(x, y):
                pyautogui.moveTo(x,y,duration=0.5)
                
            def drag_circle(radius, integers):

                coordinates = [(round(math.cos(2*pi/integers*x)*radius), round(math.sin(2*pi/integers*x)*radius)) for x in range(0, integers+1)]
                
                for x,y in coordinates:
                    print (x,y)
                    pyautogui.dragRel((x,y), 0, duration=0.1)
                
            
            def drag_half_circle(radius, integers): 
                
                coordinates = [(round(math.cos(2*pi/integers*x)*radius), round(math.sin(2*pi/integers*x)*radius)) for x in range(0, integers+1)]
            
                x = (len(coordinates))

                quarter_list = ((len(coordinates)))/4
                
                del coordinates[0:round(quarter_list)]
                del coordinates[2*round(quarter_list):x]
            
                for x,y in coordinates:
                    print (x,y)
                    pyautogui.dragRel((x,y), 0, duration=0.1)
                    
                    
            #start paint
            start_paint()
            time.sleep(2)
            move_mouse(x=500, y=200)
            time.sleep(2)

            #make eyes and face
            drag_circle(radius=20, integers=50)
            move_mouse(x=550, y=250)
            drag_circle(radius=5, integers=50)
            move_mouse(x=450, y=250)
            drag_circle(radius=5, integers=50)

            #make mouth
            move_mouse(x=600, y=400)
            drag_half_circle(radius=10, integers=50)

            #make pupils 
            move_mouse(x=430, y=270)
            drag_circle(radius=2, integers=50)

            move_mouse(x=530, y=270)
            drag_circle(radius=2, integers=50)
            

        elif "what is" in query or "who is" in query: 
            
            # Use the same API key 
            # that we have generated earlier 
            client = wolframalpha.Client("API_ID") 
            res = client.query(query) 
            
            try: 
                print (next(res.results).text) 
                speak (next(res.results).text) 
            except StopIteration: 
                print ("No results") 

        # elif "" in query: 
            # Command go here 
            # For adding more commands 

 
