import pyttsx3
import speech_recognition as sr
import datetime
import os
import subprocess
import random
import requests
from requests import get         #pip install requests
from bs4 import BeautifulSoup    #pip install bs4
from pywikihow import search_wikihow
import wikipediaapi as  wikipedia            #pip install wikipedia-api wikipedia
import webbrowser
import smtplib
import sys
import time
import pyjokes                  #pip install pyjokes
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import pyautogui
import instaloader              #pip install instaloader
import PyPDF2                   #pip install PyPDF2
import re
from PyQt5 import QtWidgets, QtCore, QtGui
from PyQt5.QtCore import QTimer, QTime, QDate, Qt
from PyQt5.QtGui import QMovie
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.uic import loadUiType
from alexui import Ui_Alex

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


# To convert text to speech
engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
#print(voices[0].id)
engine.setProperty('voice',voices[1].id)
engine.setProperty('rate',180)

#Def function for speak audio
def speak(audio):
    engine.say(audio)
    print(audio)
    engine.runAndWait()

#to wish
def wish():
    hour = int(datetime.datetime.now().hour)
    tt = time.strftime("%I:%M %p")

    if hour>= 0 and hour<= 12:
        speak(f"Hi, good morning, its {tt}")
    elif hour>12 and hour<18:
        speak(f"Hi, good afternoon, its {tt}")
    else:
        speak(f"Hi, good evening, its {tt}")
    speak(" i am Alex, please tell me how can i assist you.")





# Function to perform mathematical operations
def calculate(query):
    # Regular expression to extract numbers and operation
    matches = re.findall(r'(\d+) (plus|minus|multiply|divide) (\d+)', query)
    
    if len(matches) == 0:
        return "Sorry, I didn't understand the math query."
    
    result = ""
    for match in matches:
        num1 = int(match[0])
        operator = match[1].lower()
        num2 = int(match[2])

        if operator == "+":
            result += f"{num1} + {num2} is {num1 + num2}. "
        elif operator == "-":
            result += f"{num1} minus {num2} is {num1 - num2}. "
        elif operator == "*":
            result += f"{num1} multiplied by {num2} is {num1 * num2}. "
        elif operator == "/":
            if num2 == 0:
                result += "Division by zero error. "
            else:
                result += f"{num1} divided by {num2} is {num1 / num2}. "
        else:
            result += "Unknown operation. "
    
    return result



def pdf_reader():
    book = open('py3.pdf','rb')
    pdfReader = PyPDF2.PdfFileReader(book)
    pages = pdfReader.numPages
    speak(f"Total numbers of pages in this book {pages} ")
    speak("please ener the page number i have to read")
    pg = pdfReader.getPage(pg)
    text = pages.extractText()
    speak(text)
    # speaking speed should be controlled by user 


class MainThread(QThread):
    def __init__(self):
        super(MainThread,self).__init__()
           
    def run(self):
        self.Wake_up()


    # To convert voice into text
    def takecommand(self):
        r = sr.Recognizer()
        with sr.Microphone() as source:
            print("Listning...")
            r.pause_threshold = 1
            audio = r.listen(source,timeout=6,phrase_time_limit=8)

        try:
            print("Reconizing...")
            query = r.recognize_google(audio,language='en-IN')
            print(f"user said: {query}")

        except Exception as e:
        # speak("sorry i didn't hear anything, Say that again please...")
            return "none"
        query = query.lower()
        return query



    def TaskExecution(self):
        wish()
        while True:
        #if 1:
            
            self.query = self.takecommand()

#Logic building for tasks

            if "open notepad"in self.query:
                speak("okey, Opening notepad")
                npath = "C:\\Windows\\system32\\notepad.exe"
                os.startfile(npath)
                

            elif "close notepad" in self.query:
                speak("Okay, closing Notepad")
                os.system("taskkill /f /im notepad.exe")
                
            
            elif "open camera" in self.query:
                speak("okey, Opening camera")
                subprocess.Popen('start microsoft.windows.camera:', shell=True)
                

    # To close camera
            elif "close camera" in self.query:
                subprocess.call('taskkill /IM WindowsCamera.exe /F', shell=True)
                speak("Okay, closing camera")
            

            elif "open command prompt" in self.query:
                speak("okey, Opening cmd")
                os.system("start cmd")


            elif "play music" in self.query or "play a song" in self.query:
                speak("please wait, the query is i process")
                music_dir= "E:\\musics"
                songs = os.listdir(music_dir)
                #rd = random.choice(songs)
                for song in songs:
                    if song.endswith('.mp3'):
                        os.startfile(os.path.join(music_dir, song))

            
            elif "open youtube" in self.query:
                speak("okey, Opening youtube from browser")
                webbrowser.open("www.youtube.com")
                
                
            elif "close youtube" in self.query:
                speak("Okay, closing YouTube")
                os.system("taskkill /f /im chrome.exe /t")
                

            elif "open facebook" in self.query:
                speak("okey, Opening Facebook")
                webbrowser.open("www.facebook.com")
                

            elif "open stack overflow" in self.query:
                speak("okey, Opening stack overflow")
                webbrowser.open("www.stackoverflow.com")
                

            elif "open google" in self.query:
                speak("okey, Opening google")
                speak("what should i search on google")
                cm = self.takecommand().lower()
                webbrowser.open(f"{cm}")

            elif "close google" in self.query or "close chrome" in self.query:
                speak("Okay, closing Google")
                os.system("taskkill /f /im chrome.exe")
                
                
    #To Switch the desktop Window
            elif 'switch the window' in self.query:
                speak("okey, please wait")
                pyautogui.keyDown("alt")
                pyautogui.press("tab")
                time.sleep(1)
                pyautogui.keyUp("alt") 
                       
                
    #To set an alarm
            elif"set alarm" in self.query:
                nn = int(datetime.datetime.now().hour)
                if nn==22:
                    music_dir = 'E:\\music'
                    songs = os.listdir(music_dir)
                    os.startfile(os.path.join(music_dir,song[0])) 
                           
                    
    #To get ip an address
            elif "ip address" in self.query:
                speak("please wait, i am getting ip address")
                speak("Switching window")
                ip = get('https://api.ipify.org').text
                speak(f"your IP address is {ip}")
                
                
    #To search from wikipedia.
            elif "open wikipedia" in self.query:
                speak("searching wikipedia...")
                self.query = self.query.replace("wikipedia","")
                results = wikipedia.summary(self.query,sentences=2)
                speak("according to wikipedia")
                speak(results)
                print(results)
                
        
    #To do how to query
            elif "how to" in self.query:
                # Extract the part of the query after how to.
                parts = self.query.split("how to", 1)
                if len(parts) > 1:
                    query_words = parts[1].strip()
                    max_result = 1
                    results = search_wikihow(query_words, max_result)
                    if results and len(results) > 0:
                        # Print the result
                        results[0].print()
                        speak(results[0].summary)
                    else:
                        speak("Sorry, I couldn't find any information on that topic.")
                else:
                    speak("Please specify what you want to know.")        
                

    # To get weather forecast
            elif "temperature" in self.query:
                speak("wait a second i am in process")
                parts = self.query.split("in")
                if len(parts) == 2:
                    city = parts[1].strip()
                else:
                    speak("City name is missing.")
                    return
                # Construct the search query
                search = f"temperature in {city}"
                url = f"https://www.google.com/search?q={search}"
                r = requests.get(url)
                data = BeautifulSoup(r.text,"html.parser")
                temp = data.find("div",class_="BNeawe").text
                speak(f"current {search} is {temp}")

            
            #elif "play " in self.query:
                #parts = self.query.split( 1)
                #kit.playonyt("see you again")

    # To terminate 
            elif "you can sleep" in self.query:
                speak("okey, thanks for using me, have a great day")
                sys.exit()
                

    # To find jokes
            elif "tell me jokes" in self.query:
                joke = pyjokes.get_joke()
                speak(joke)
                

    #To Calculation.
            elif "calculate" in self.query:
                result = calculate(self.query)
                speak(result)
                      

    # To find my location using Ip Address
            elif "where i am" in self.query or "where we are" in self.query:
                speak("please wait, im gathering information")
                # Define URL for IP geolocation service
                url = "https://ipinfo.io/json"
                
                try:
                    # Send request to get location data
                    r = requests.get(url)
                    data = r.json()  # Parse JSON response

                    # Extract location information
                    region = data.get("region", "unknown region")
                    country = data.get("country", "unknown country")
                    
                    speak(f"Now you are in {region} State of {country}")
                except requests.RequestException as e:
                    speak(f"An error occurred while fetching location: {e}")
                               
            
    # To check a instagram profile
            elif "instagram profile" in self.query or "profile on instagram" in self.query:
                speak("please  input the user name correctly in terminal.")
                name = input("Enter the username here:")
                webbrowser.open(f"www.instagram.com/{name}")
                speak(f"here is the profile of the user {name}")
                time.sleep(5)
                    
                
    # To Take screenshot
            elif "take screenshot" in self.query or "take a screenshot" in self.query:
                speak("Please tell me the name for this screenshot file")
                
                name = self.takecommand().lower()
                
                speak("Please hold the screen for a few seconds, I am taking the screenshot")
                
                time.sleep(3)
                downloadspath = os.path.join(os.path.expanduser("~"), "Downloads")
                screenshot_path = os.path.join(downloadspath, f"{name}.png")
                img = pyautogui.screenshot()
                img.save(screenshot_path)
                
                speak(f"I am done. The screenshot is saved in your Downloads folder as {name}.png.")

                
    # To read PDF file
            elif "read pdf" in self.query:
                pdf_reader()
                
                
    # To hide folders
            elif "hide all files" in self.query or "hide this folder" in self.query or "visible for everyone" in self.query:
                speak("please tell me you want to hide this folder or make it visible for everyone")
                condition = self.takecommand().lower()
                if "hide" in condition:
                    os.system("attrib +h /s /d")
                    speak("all the files in this folder are now hidden.")      

                elif "visible" in self.query:
                    os.system("attrib -h /s /d")
                    speak("all the files in this folder are now visible to everyone.")
                    
                elif "leave it" in self.query or "leave for now" in self.query:
                    speak("okey")            
                        
    

            elif "tell me news" in self.query:
                speak("please wait, featching the latest news")
                
                

    # To do some athematical operation
            



    
            # speak("sir, do you have any other work")

    #To shutdown,restart and sleep the system
            elif "shutdown the system" in self.query:
                os.system("shutdown /s/ t/ 5")
                

            elif"restart the system" in self.query:
                
                os.system("shutdown /r /t 5")
                

            elif "sleep the system" in self.query:
                os.system("rundll32.exe powrprof.dll,SetSuspendState 0,1,0")



    #To talk with assist. 
            elif "Hello" in self.query or "hey" in self.query:
                speak("hello, may i help you.")
                

            elif "how are you" in self.query:
                speak("i am fine, what about you.")
                

            elif "good" in self.query or "i am fine" in self.query or "fine" in self.query:
                speak("that's great to hear from you.")
                

            elif"Thank you" in self.query or "Thanks" in self.query:
                speak("it's my pleasure.")
                

            elif "you can sleep" in self.query or "you can sleep now" in self.query:
                speak("okey, i am going to sleep you can call me anytime.")
                break


    def Wake_up(self):
        while True:
                self.permission = self.takecommand()

                if "hey alex" in self.permission:
                 self.TaskExecution()
                 

                elif"sleep" in self.permission:
                    speak("thanks for using me. have a good day.")
                    sys.exit()
        
                


startExecution = MainThread()


class Main(QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_Alex()
        self.ui.setupUi(self)
        self.ui.pushButton.clicked.connect(self.startTask)
        self.ui.pushButton_2.clicked.connect(self.close)
        
        
    def startTask(self):
        self.ui.movie = QtGui.QMovie("E:/alex/Sound wave for AI.gif")
        self.ui.label.setMovie(self.ui.movie)
        self.ui.movie.start()
        startExecution.start()


app = QApplication(sys.argv)
alex = Main()
alex.show()
(app.exec_())
