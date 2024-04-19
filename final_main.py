import os
import logging
import pyttsx3
logging.disable(logging.WARNING)
os.environ['TF_CPP_MIN_LOG_LEVEL'] = '3' # disabling warnings for gpu requirements
from keras_preprocessing.sequence import pad_sequences
import numpy as np
from keras.models import load_model
from pickle import load
import speech_recognition as sr
import sys
import pyautogui
sys.path.insert(0, os.path.expanduser('~')+"/Virtual-Voice-Assistant") # adding voice assistant directory to system path
import google.auth
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import requests
import googleapiclient.discovery
import webbrowser
import math
import psutil
import time
from random import randint
import subprocess
import AppOpener
from pynput.keyboard import Key, Controller
from PIL import ImageGrab
import wmi
import webbrowser
import re
import wikipedia
import speedtest
from youtubesearchpython import VideosSearch
import pandas as pd


websites_dict = {
    "google": "https://www.google.com/",
    "gmail": "https://mail.google.com/mail/u/0/#inbox",
    "myclass": "https://myclass.lpu.in/",
    "lpulive": "https://lpulive.lpu.in/login",
    "ums": "https://ums.lpu.in/lpuums/",
    "github": "https://github.com/",
    "youtube": "https://www.youtube.com/",
    "sanfoundry": "https://www.sanfoundry.com/",
    "hackerrank": "https://www.hackerrank.com/dashboard",
    "hackerearth": "https://www.hackerearth.com/challenges/",
    "codeforces": "https://codeforces.com/enter?back=%2Fproblemset",
    "codechef": "https://www.codechef.com/",
    "codecademy": "https://www.codecademy.com/catalog",
    "leetcode": "https://leetcode.com/",
    "udemy": "https://www.udemy.com/",
    "edx": "https://www.edx.org/",
    "coursera": "https://www.coursera.org/",
    "w3resource": "https://www.w3resource.com/",
    "w3schools": "https://www.w3schools.com/",
    "geeksforgeeks": "https://www.geeksforgeeks.org/",
    "stackoverflow": "https://stackoverflow.com/",
    "instagram": "https://www.instagram.com/?hl=en",
    "facebook": "https://www.facebook.com/",
    "linkedin": "https://www.linkedin.com/",
    "twitter": "https://twitter.com/LOGIN",
    "imdb": "https://imdb.com/",
    "netflix": "https://netflix.com/",
    "amazon prime": "https://primevideo.com/",
    "whatsapp": "https://web.whatsapp.com/",
    #"telegram": "https://telegram.org/",
    "hotstar": "https://www.hotstar.com/",
    "cricbuzz": "https://www.cricbuzz.com/",
    "google maps": "https://maps.google.com/",
    "flipkart": "https://www.flipkart.com/",
    "amazon": "https://www.amazon.com/",
    "nykaa": "https://www.nykaa.com/",
    "myntra": "https://www.myntra.com/",
    "amazon flex": "https://flex.amazon.in/",
    #"discord": "https://discord.com/",
    "swiggy": "https://www.swiggy.com/",
    "zomato": "https://www.zomato.com/",
    "spotify": "https://open.spotify.com/"
}

def add_data(speech, intent):
    df = pd.read_csv("C:/Users/admin/OneDrive/Documents/VIT/Capstone/testing.csv")
    new_data = {
    'recognized_speech': speech,
    'ground_truth': '',
    'recognized_intent': intent,
    'expected_intent': '',
    'intent_based_on_recognized_speech':intent}
    df = df.append(new_data,ignore_index=True)
    df.to_csv("C:/Users/admin/OneDrive/Documents/VIT/Capstone/testing.csv",index=False)
    return
    
class SystemTasks:
    def __init__(self):
        self.keyboard = Controller()

    def write(self, text):
        self.keyboard.type(text)

    def select(self):
        self.keyboard.press(Key.ctrl)
        self.keyboard.press('a')
        self.keyboard.release('a')
        self.keyboard.release(Key.ctrl)

    def hitEnter(self):
        self.keyboard.press(Key.enter)
        self.keyboard.release(Key.enter)
    
    def tab_key(self):
        self.keyboard.press(Key.tab)
        self.keyboard.release(Key.tab)
    
    def left(self):
        self.keyboard.press(Key.left)
        self.keyboard.release(Key.left)
    
    def right(self):
        self.keyboard.press(Key.right)
        self.keyboard.release(Key.right)

    def up(self):
        self.keyboard.press(Key.up)
        self.keyboard.release(Key.up)
    
    def down(self):
        self.keyboard.press(Key.down)
        self.keyboard.release(Key.down)
    
    def delete(self):
        self.select()
        self.keyboard.press(Key.backspace)
        self.keyboard.release(Key.backspace)

    def copy(self):
        self.select()
        self.keyboard.press(Key.ctrl)
        self.keyboard.press('c')
        self.keyboard.release('c')
        self.keyboard.release(Key.ctrl)

    def paste(self):
        self.keyboard.press(Key.ctrl)
        self.keyboard.press('v')
        self.keyboard.release('v')
        self.keyboard.release(Key.ctrl)

    def new_file(self):
        self.keyboard.press(Key.ctrl)
        self.keyboard.press('n')
        self.keyboard.release('n')
        self.keyboard.release(Key.ctrl)

    def save(self, name):
        self.keyboard.press(Key.ctrl)
        self.keyboard.press('s')
        self.keyboard.release('s')
        self.keyboard.release(Key.ctrl)
        time.sleep(0.2)
        self.write(name)
        self.hitEnter()


class TabOpt:
    def __init__(self):
        self.keyboard = Controller()

    def switchTab(self):
        self.keyboard.press(Key.ctrl)
        self.keyboard.press(Key.tab)
        self.keyboard.release(Key.tab)
        self.keyboard.release(Key.ctrl)

    def closeTab(self):
        self.keyboard.press(Key.ctrl)
        self.keyboard.press('w')
        self.keyboard.release('w')
        self.keyboard.release(Key.ctrl)

    def newTab(self):
        self.keyboard.press(Key.ctrl)
        self.keyboard.press('t')
        self.keyboard.release('t')
        self.keyboard.release(Key.ctrl)


class WindowOpt:
    def __init__(self):
        self.keyboard = Controller()

    def closeWindow(self):
        self.keyboard.press(Key.alt_l)
        self.keyboard.press(Key.f4)
        self.keyboard.release(Key.f4)
        self.keyboard.release(Key.alt_l)

    def minimizeWindow(self):
        for i in range(2):
            self.keyboard.press(Key.cmd)
            self.keyboard.press(Key.down)
            self.keyboard.release(Key.down)
            self.keyboard.release(Key.cmd)
            time.sleep(0.05)

    def maximizeWindow(self):
        self.keyboard.press(Key.cmd)
        self.keyboard.press(Key.up)
        self.keyboard.release(Key.up)
        self.keyboard.release(Key.cmd)

    def switchWindow(self):
        self.keyboard.press(Key.alt_l)
        self.keyboard.press(Key.tab)
        self.keyboard.release(Key.tab)
        self.keyboard.release(Key.alt_l)

    def Screen_Shot(self):
        im = ImageGrab.grab()
        im.save(f'../Data/Screenshots/ss_{randint(1, 100)}.jpg')

recognizer = sr.Recognizer()

engine = pyttsx3.init()
engine.setProperty('rate', 185)

sys_ops = SystemTasks()
tab_ops = TabOpt()
win_ops = WindowOpt()

# load trained model
model = load_model('C:/Users/admin/OneDrive/Documents/VIT/Capstone/model')

# load tokenizer object
with open('C:/Users/admin/OneDrive/Documents/VIT/Capstone/model/tokenizer.pickle', 'rb') as handle:
    tokenizer = load(handle)

# load label encoder object
with open('C:/Users/admin/OneDrive/Documents/VIT/Capstone/model/label_encoder.pickle', 'rb') as enc:
    lbl_encoder = load(enc)

def speak(text):
    print("ASSISTANT -> " + text)
    try:
        engine.say(text)
        engine.runAndWait()
    except KeyboardInterrupt or RuntimeError:
        return
def app_path(app):
    app_paths = {'access': 'C://Program Files//Microsoft Office//root//Office16//ACCICONS.exe',
                 'powerpoint': 'C://Program Files//Microsoft Office//root//Office16//POWERPNT.exe',
                 'word': 'C://Program Files//Microsoft Office//root//Office16//WINWORD.EXE',
                 'excel': 'C://Program Files//Microsoft Office//root//Office16//EXCEL.exe',
                 'outlook': 'C://Program Files//Microsoft Office//root//Office16//OUTLOOK.exe',
                 'onenote': 'C://Program Files//Microsoft Office//root//Office16//ONENOTE.exe',
                 'publisher': 'C://Program Files//Microsoft Office//root//Office16//MSPUB.exe',
                 'sharepoint': 'C://Program Files//Microsoft Office//root//Office16//GROOVE.exe',
                 'infopath designer': 'C://Program Files//Microsoft Office//root//Office16//INFOPATH.exe',
                 'infopath filler': 'C://Program Files//Microsoft Office//root//Office16//INFOPATH.exe'}
    try:
        return app_paths[app]
    except KeyError:
        return None


def open_app(query):
    ms_office = ('access', 'powerpoint', 'word', 'excel', 'outlook', 'onenote', 'publisher', 'sharepoint', 'infopath designer',
                 'infopath filler')
    for app in ms_office:
        if app in query:
            path = app_path(app)
            subprocess.Popen(path)
            return True
    AppOpener.run(query[5:])
    return True

def chat(text):
    # parameters
    max_len = 20
    while True:
        result = model.predict(pad_sequences(tokenizer.texts_to_sequences([text]),
                                                                          truncating='post', maxlen=max_len), verbose=False)
        intent = lbl_encoder.inverse_transform([np.argmax(result)])[0]
        return intent

def record():
    with sr.Microphone() as mic:
        recognizer.adjust_for_ambient_noise(mic)
        recognizer.dynamic_energy_threshold = True
        print("Listening...")
        audio = recognizer.listen(mic)
        try:
            text = recognizer.recognize_google(audio, language='us-in').lower()
        except:
            return None
    print("USER -> " + text)
    return text
    
def listen_audio():
    try:
        while True:
            response = record()
            if response is None:
                continue
            else:
                main_search(response)
                speak("Listening..")
    except KeyboardInterrupt:
        return
    
def open_specified_website(query):
    website = query[5:] #re.search(r'[a-zA-Z]* (.*)', query)[1]
    if website in websites_dict:
        url = websites_dict[website]
        webbrowser.open(url)
        return True
    else:
        return None

help_dict = {
    'one': 1,
    'two': 2,
    'three': 2,
    'four': 4,
    'five': 5,
    'six': 6,
    'seven': 7,
    'eight': 8,
    'nine': 9,
    'ten': 10
}

API_KEY = 'AIzaSyA3h6ezsah6gcrhzJqdCJBIPuIoyAarmVM'
SEARCH_ENGINE_ID = "502d094dbad334e45"

urls = [] 

def googleSearch(query):
    if 'image' in query:
        query += "&tbm=isch"
    query = query.replace('images', '')
    query = query.replace('image', '')
    query = query.replace('search', '')
    query = query.replace('find', '')
    query = query.replace('find about', '')
    query = query.replace('show', '')
    query = query.replace('google', '')
    query = query.replace('tell me about', '')
    query = query.replace('for', '')
    webbrowser.open("https://www.google.com/search?q=" + query)
    return "Here you go..."
   
def search_results(query):
    
    page = 1
    
    # calculating start, (page=2) => (start=11), (page=3) => (start=21)
    start = (page - 1) * 10 + 1
    url = f"https://www.googleapis.com/customsearch/v1?key={API_KEY}&cx={SEARCH_ENGINE_ID}&q={query}&start={start}"
    
    data = requests.get(url).json()
    search_items = data.get("items")
    url=[]
    for i, search_item in enumerate(search_items, start=1):
        try:
            long_description = search_item["pagemap"]["metatags"][0]["og:description"]
        except KeyError:
            long_description = "N/A"
        title = search_item.get("title")
        snippet = search_item.get("snippet")
        html_snippet = search_item.get("htmlSnippet")
        link = search_item.get("link")
        url.append(link)
        print("="*10, f"Result #{i+start-1}", "="*10)
        print("Title:", title)
        print("Description:", snippet)
        print("Long description:", long_description)
        print("URL:", link, "/n")
    return url

def main_search(query):
        # add_data(query)
        
        intent = chat(query)
        done = False
        print(intent)
        if ("google" in query and "search" in query) or ("google" in query and "how to" in query) or "google" in query:
            speak("Google Search executing")
            googleSearch(query)
            urls = search_results(query)
            done = True
            speak("Google Search executed")
            intent = "google_search"
        elif (query == "open url" and "url" in query) or (query == "open urls" and "url" in query):
            print(urls)
            webbrowser.open(f"{urls[0]}")
            done = True
        elif ("youtube" in query and "search" in query) or "play" in query or ("how to" in query and "youtube" in query):
            speak("Opening Youtube")
            duplicate_query = query
            duplicate = query
            duplicate = duplicate.replace('play', ' ')
            duplicate = duplicate.replace(' on youtube', ' ')
            duplicate = duplicate.replace('youtube', ' ')
            duplicate = duplicate.replace('search', ' ')
            print("Searching for videos...")
            videosSearch = VideosSearch(duplicate, limit=1)
            results = videosSearch.result()['result']
            print("Finished searching!")
            if "play" in duplicate_query:
                webbrowser.open('https://www.youtube.com/watch?v=' + results[0]['id'])
            else:
                webbrowser.open('https://www.youtube.com/results?search_query='+query)
            done = True
            intent = "youtube_search"
        elif "distance" in query or "map" in query:
            webbrowser.open(f'https://www.google.com/maps/search/{query}')
            done = True
        elif intent =="write":
            speak("what would you like to take down?")
            note = record()
            pyautogui.write(note)
            done = True
        elif intent == "hit_enter":
            speak("hitting enter")
            sys_ops.hitEnter()
            done = True
        elif intent == "save_file":
            speak("What would the name of the file be?")
            name = record()
            sys_ops.save(name)
            done = True
        elif intent=="left":
            speak("moving left")
            sys_ops.left()
            done = True
        elif intent=="right":
            speak("moving right")
            sys_ops.right()
            done = True
        elif intent=="up":
            speak("moving up")
            sys_ops.up()
            done = True
        elif intent=="down":
            speak("moving down")
            sys_ops.down()
            done = True
        elif intent == "tab_key":
            sys_ops.tab_key()
            done = True
        elif intent == "select_text" and "select" in query:
            speak("selecting text")
            sys_ops.select()
            done = True
        elif intent == "copy_text" and "copy" in query:
            speak("copying text")
            sys_ops.copy()
            done = True
        elif intent == "paste_text":
            speak("pasting text")
            sys_ops.paste()
            done = True
        elif intent == "delete_text" and "delete" in query:
            speak("deleting text")
            sys_ops.delete()
            done = True
        elif intent == "new_file" and "new" in query:
            speak("opening new file")
            sys_ops.new_file()
            done = True
        elif intent == "switch_tab" and "switch" in query and "tab" in query:
            speak("switching tab")
            tab_ops.switchTab()
            done = True
        elif intent == "close_tab" and "close" in query and "tab" in query:
            speak("closing tab")
            tab_ops.closeTab()
            done = True
        elif intent == "new_tab" and "new" in query and "tab" in query:
            speak("opening new tab")
            tab_ops.newTab()
            done = True
        elif intent == "close_window" and "close" in query:
            speak("closing window")
            win_ops.closeWindow()
            done = True
        elif intent == "switch_window" and "switch" in query:
            speak("switching window")
            win_ops.switchWindow()
            done = True
        elif intent == "minimize_window" and "minimize" in query:
            speak("minimizing window")
            win_ops.minimizeWindow()
            done = True
        elif intent == "maximize_window" and "maximize" in query:
            speak("maximizing window")
            win_ops.maximizeWindow()
            done = True
        elif intent == "screenshot" and "screenshot" in query:
            p = pyautogui.screenshot()
            p.save(r'C://Users//admin//Pictures//ScreenShots//p.png')
            done = True
        elif intent == "stopwatch":
            pass
        elif intent == "wikipedia" and ("tell" in query or "about" in query):
            description = None
            try:
                topic = query.replace("tell me about ", "") #re.search(r'([A-Za-z]* [A-Za-z]* [A-Za-z]*)$', query)[1]
                result = wikipedia.summary(topic, sentences=3)
                result = re.sub(r'/[.*]', '', result)
                description = result
            except (wikipedia.WikipediaException, Exception) as e:
                description = None
            if description:
                speak(description)
            else:
                googleSearch(query)
            done = True
        elif intent == "open_app":
            completed = open_app(query)
            if completed:
                done = True
        elif intent == "open_website":
            completed = open_specified_website(query)
            if completed:
                done = True
        elif intent == "note" and "note" in query:
            speak("what would you like to take down?")
            note = record()
            open_app("open notepad")
            time.sleep(0.2)
            sys_task = SystemTasks()
            sys_task.write(note)
            # sys_task.save(f'note_{randint(1, 100)}')
            done = True
        elif intent == "exit" and ("exit" in query or "terminate" in query or "quit" in query):
            exit(0)
        if not done:
            speak("Sorry, not able to answer your query")
            add_data(query,"None")
        else:
            add_data(query,intent)
        return

if __name__ == "__main__":
    try:
        speak("Hello I am online. How can I help you? ")
        listen_audio()
    except:
        speak("Quitting program")
        print("Exit")