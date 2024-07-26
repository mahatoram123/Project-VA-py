import pyttsx3
import speech_recognition as sr
import datetime
import os
import cv2
import random
import requests
from requests import get
import wikipedia
import webbrowser
import pywhatkit as kit
import smtplib
import sys
import time
import pyjokes
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import pyautogui
import instaloader
import PyPDF2
import re

"""
IN PLACE OF PYTTSX3 WE CAN ALSO USE WIN32COM.CLIENT

# python program to convert
# text to speech

# import the required module from text to speech conversion
import win32.com.client

# calling the Dispatch method of the module which
# intersect with Microsoft speech SDK to speak
# the given input from the keyboard

speaker = win32com.client.Dispatch("SAPI.spVoice")

while 1:
    print("Enter the word you want to speak it out by computer")
    s = input(s)
    speaker.Speak(s)

# To stop the program press
3 CTRL + z
"""


engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
#print(voices[0].id)
engine.setProperty('voices',voices[0].id)

#text tp speech
def speak(audio):
    engine.say(audio)
    print(audio)
    engine.runAndWait()


# To convert voice into text
def takecommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listning...")
        r.pause_threshold = 1
        audio = r.listen(source,timeout=5,phrase_time_limit=8)

    try:
        print("Reconizing...")
        query = r.recognize_google(audio,language='en-IN')
        print(f"user said: {query}")

    except Exception as e:
        speak("sorry i didn't hear anything, Say that again please...")
        return "none"
    return query

#to wish
def wish():
    hour = int(datetime.datetime.now().hour)
    tt = time.strftime("%I:%M %p")

    if hour>= 0 and hour<= 12:
        speak(f"good morning, its {tt}")
    elif hour>12 and hour<18:
        speak(f"good afternoon, its {tt}")
    else:
        speak(f"Hi, good evening, its {tt}")
    speak(" i am your voice assistant, please tell me how can i help you")


#For news uupdate
def news():
    main_url = 'http://newsapi.org/v2/top-headlines?source=techcrunch&apiKey=83e1a7ee898e495ba1ad9e16ff9973b0'

    main_page = requests.get(main_url).json()
    #print(main_page)
    articles = main_page["articles"]
    #print(articles)
    head = []
    day=["first","second","tird","fourth","fifth","sixth","seventh","eighth","ninth","tenth"]
    for ar in articles:
        #print(f"today's {day[i]} news is: ",head[i])
        speak(f"today's {day[i]} news is: ", head[i])


# Function to perform mathematical operations
def calculate(query):
    # Regular expression to extract numbers and operation
    matches = re.findall(r'(\d+) ([plus|minus|multiply|divide]+) (\d+)', query)
    if len(matches) == 0:
        return "Sorry, I didn't understand the math query."
    
    result = ""
    for match in matches:
        num1 = int(match[0])
        operator = match[1].lower()
        num2 = int(match[2])

        if operator == "plus":
            result += f"{num1} plus {num2} is {num1 + num2}. "
        elif operator == "minus":
            result += f"{num1} minus {num2} is {num1 - num2}. "
        elif operator == "multiply":
            result += f"{num1} multiplied by {num2} is {num1 * num2}. "
        elif operator == "divide":
            if num2 == 0:
                result += "Division by zero error. "
            else:
                result += f"{num1} divided by {num2} is {num1 / num2}. "

    return result

def pdf_reader():
    book = open('py3.pdf','rb')
    pdfReader = PyPDF2.PdfFileReader(book)
    pages = pdfReader.numPages
    speak(f"Total numbers of pages in this book {pages} ")
    speak("please ener the page number i have to read")
    pg + pdfReader.getPage(pg)
    text = pages.extractText()
    speak(text)
    # speaking speed should be controlled by user 

if __name__ == "__main__":
    wish()
    while True:
    #if 1:
        
        query = takecommand().lower()

        #Logic building for tasks

        if "open notepad"in query:
            npath = "C:\\Windows\\system32\\notepad.exe"
            os.startfile(npath)

        elif "open camera" in query:
            cap = cv2.VideoCapture(0)
            while True:
                ret, img = cap.read()
                cv2.imshow('webcam', img)
                k = cv2.waitKey(50)
                if k==27:
                    break;
            cap.release()
            cv2.destroyAllWindows()

        # To close camera
        elif "close camera" in query:
            speak("Okay, closing camera")
            cap.release()
            cv2.destroyAllWindows()
        


        elif "open command prompt" in query:
            os.system("start cmd")

        

        elif "play music" in query:
            music_dir= "E:\\musics"
            songs = os.listdir(music_dir)
            #rd = random.choice(songs)
            for song in songs:
                if song.endswith('.mp3'):
                    os.startfile(os.path.join(music_dir, song))

        
        
        elif "ip address" in query:
            ip = get('https://api.ipify.org').text
            speak(f"your IP address is {ip}")

        elif "open wikipedia" in query:
            speak("searching wikipedia...")
            query = query.replace("wikipedia","")
            results = wikipedia.summary(query,sentences=2)
            speak("according to wikipedia")
            speak(results)
            #print(results)

        elif "open youtube" in query:
            webbrowser.open("www.youtube.com")

        elif "open facebook" in query:
            webbrowser.open("www.facebook.com")

        elif "open stack overflow" in query:
            webbrowser.open("www.stackoverflow.com")

        elif "open google" in query:
            speak("what should i search on google")
            cm = takecommand().lower()
            webbrowser.open(f"{cm}")

        #elif "send whatsapp massage" in query:
        #    kit.sendwhatmsg("+9779817889629","this is testing massage of py",4,13)
        #    time.sleep(120)
        #    speak("massage has been sent")
        
        #elif "play songs on youtube" in query:
        #    kit.playonyt("see you again")

        #elif "send email to ram" in query:
        #    try:
        #        speak("what should i say")
        #        content = takecommand().lower()
        #        to= "mahatoram023@gmail.com"
        #        sendEmail(to,content)
        #        speak("Email has been sent to ram")

        #    except Exception as e:
        #        print(e)
        #        speak(" sorry, i am not able to sent this mail to ram")

        elif "you can sleep" in query:
            speak("thanks for using me, have a great day")
            sys.exit()


 #To close any application
        elif "close notepad" in query:
            speak("Okay, closing Notepad")
            os.system("taskkill /f /im notepad.exe")

        elif "close google" in query or "close chrome" in query:
            speak("Okay, closing Google Chrome")
            os.system("taskkill /f /im chrome.exe")

        elif "close youtube" in query:
            speak("Okay, closing YouTube")
            os.system("taskkill /f /im chrome.exe /t")
#To set an alarm
        elif"set alarm" in query:
            nn = int(datetime.datetime.now().hour)
            if nn==22:
                music_dir = 'E:\\music'
                songs = os.listdir(music_dir)
                os.startfile(os.path.join(music_dir,song[0]))
# To find jokes
        elif "tell me joke" in query:
            joke = pyjokes.get_joke()
            speak(joke)

        elif"shutdown the system" in query:
            os.system("shutdown /s/ t/ 5")

        elif"restart the system" in query:
            os.system("shutdown /r /t 5")

        elif "sleep the system" in query:
            os.system("rundll32.exe powrprof.dll,SetSuspendState 0,1,0")



####################################################################################################################################################
####################################################################################################################################################

        
        elif 'switch the window' in query:
            pyautogui.keyDown("alt")
            pyautogui.press("tab")
            time.sleep(1)
            pyautogui.keyUp("alt")

        elif "tell me news" in query:
            speak("please wait, featching the latest news")
            news()


    # To find my location using Ip Address
        elif "where i am" in query or "where we are" in query:
            speak("Wait, let me check")
            try:
                ipAdd = get('https://api.ipify.org').text
                url = 'https://geo.ipify.org/api/v1?apiKey=YOUR_API_KEY&ipAddress=' + ipAdd
                geo_requests = requests.get(url)
                geo_data = geo_requests.json()

                city = geo_data['location']['city']
                country = geo_data['location']['country']
                speak(f"I think we are in {city} city, {country} country.")
            except Exception as e:
                speak("Sorry, I am not able to find where we are due to network issues.")


    # To check a instagram profile
        elif "instagram profile" in query or "profile on instagram" in query:
            speak("sir please enter the user name correctly.")
            name = input("Enter the username here:")
            webbrowser.open(f"www.instagram.com/{name}")
            speak(f"here is the profile of the user {name}")
            time.sleep(5)
            speak("would you like to download profile picture of this account.")
            condition = takecommand().lower()
            if "yes" in condition:
                mod = instaloader.Instaloader()
                mod.download_profile(name, profile_pic_only=True)
                speak("I am done, the profile picture is saved in our main folder. Now I am ready.")
            else:
             pass
                

    # To Take screenshot

        elif "take screenshot" in query or "take a screenshot" in query:
            speak("please tell me the name for this screemshot file")
            name = takecommand().lower()
            speak("please hold the sceen for few seconds, i am taking screenshot")
            time.sleep(3)
            img = pyautogui.screenshot()
            img.save(f"{name}.png")
            speak("i am done, the screenshot is saved in our main folder. now i am ready for other work")


    # To read PDF file
        elif "read pdf" in query:
            pdf_reader()

    # To do some athematical operation
       



    # To hide folders
        elif "hide all files" in query or "hide this folder" in query or "visible for everyone" in query:
            speak("please tell me you want to hide this folder or make it visible for everyone")
            condition = takecommand().lower()
            if "hide" in condition:
                os.system("attrib +h /s /d")
                speak("all the files in this folder are now hidden.")

        elif "visible" in query:
            os.system("attrib -h /s /d")
            speak("all the files in this folder are now visible to everyone.")

        elif "leave it" in query or "leave for now" in query:
            speak("okey")
        # speak("sir, do you have any other work")
