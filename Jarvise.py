import cv2
import datetime
import os
# import Tools.scripts.google
import smtplib
import webbrowser
from time import ctime

import pyttsx3
import speech_recognition as sr
import wikipedia
import win32com.servers.dictionary
from googlesearch import search
import math

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
# print(voices[0].id)
engine.setProperty('voice', voices[0].id)


def speak(audio):
    engine.say(audio)
    engine.runAndWait()


def wishMe():
    hour = int(datetime.datetime.now(). hour)
    if hour >= 0 and hour < 12:
        speak("Good morning")

    elif hour >= 12 and hour < 18:
        speak("Good Afternoon")

    else:
        speak("Good Evening")

    speak('my name is Jarvice. how can i help you')


def takeCommand(ask=False):
    # It takes microphone input from the user and returns string output

    r = sr.Recognizer()
    with sr.Microphone()as source:
        speak("Hearing Sir")
        print("Hearing Sir...")
        r.pause_threshold = 1
        if ask:
            print(ask)
        audio = r.listen(source)

    try:
        print("recognizing...")
        quary = r.recognize_google(audio, language='hi=in')
        print(f"User said: {quary}\n") 

    except Exception as e:
        # print(e)
        speak("Sorry Sir can you repeat that one")
        print("Say that again Please...")
        return "None"
    return quary


def sendEmailpapa(to, content):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    server.login('anchitlahkar0202@gmail.com', 'manyata222007')
    server.sendmail('gautamlahkar@gmail.com', to, content)
    server.close()


def sendEmailma(to, content):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    server.login('anchitlahkar0202@gmail.com', 'manyata222007')
    server.sendmail('paramitachakravarty.com', to, content)
    server.close()


if __name__ == '__main__':
    wishMe()
    while True:
        query = takeCommand().lower()

        # logic for executing task based on query
        if 'wikipedia' in query:
            speak("Searching Wikipedia....")
            query = query.replace("wikipedia", "")
            url = 'https://www.wikipedia.org/'
            results = wikipedia.summary(query, sentences=5)
            speak("according to wikipedia")
            webbrowser.get().open(url)
            print(results)
            speak(results)
        
        # elif 'open youtube' in query:
        #     speak("opening Youtube")
        #     webbrowser.open("youtube.com")

        elif 'open youtube' in query:
            speak('Opening youtube')
            url = 'https://www.youtube.com/'
            webbrowser.get().open(url)

        # elif 'open google' in query:
        #     speak("opening google")
        #     webbrowser.open("google.com")

        elif 'open google' in query:
            speak('Opening google')
            url = 'https://www.google.com/'
            webbrowser.get().open(url)

        # elif 'open nasa' in query:
        #     speak("opening Nasa.gov")
        #     webbrowser.open("nasa.gov")

        elif 'open nasa' in query:
            speak('Opening nasa')
            url = 'https://www.nasa.gov/'
            webbrowser.get().open(url)
            

    #    elif 'play music' in query:
    #        music_dic = "D:\\ wikimusic"
    #        songs = os.listdir(music_dic)
    #        print(songs)
    #        os.startfile(os.path.join(music_dic, songs[0]))

        elif 'play music' in query:
            speak('playing music')
            url = 'https://wynk.in/music'
            webbrowser.get().open(url)

        elif 'the time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"Sir, The time is {strTime}")


        elif 'open microsoft team' in query :
            speak('opening Microsoft team')
            codePath = "C:\\Users\\lenevo\\AppData\\Local\\Microsoft\\Teams"
            os.startfile(codePath)


        elif "open WhatApp" in query :
            speak("opening whatapp")
            codePath = "C:\\Users\\lenevo\\AppData\\Local\\WhatsApp\\WhatsApp.exe"
            os.startfile(codePath)

        
        elif "open zoom" in query :
            speak("opening whatapp")
            codePath = "C:\\Users\\lenevo\\AppData\\Roaming\\Zoom\\bin\\Zoom.exe"
            os.startfile(codePath)

        elif "open rammarly" in query :
            speak("opening Grammarly")
            codePath = "C:\\Users\\lenevo\\AppData\\Local\\GrammarlyForWindows\\GrammarlyForWindows.exe"
            os.startfile(codePath)

        elif "open office" in query :
            speak("opening Office")
            codePath = "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome_proxy.exe"
            os.startfile(codePath)



        elif 'open programming' in query:
            speak("opening visiul studio code")
            codePath = "C:\\Users\\lenevo\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"
            os.startfile(codePath)


        elif 'open microsoft' in query:
            speak('opening Microsoft Edge')
            searchPath = 'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe'
            os.startfile(searchPath)

        elif 'open unity' in query:
            speak("opening unity")
            codePath = "F:\\Unitygaming\\2019.4.2f1\\Editor\\Unity.exe"
            os.startfile(codePath)

        # elif 'play music' in query:
        #     speak("playing Music")
        #     chromePath = "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome_proxy.exe  --profile-directory=Profile 1 --app-id=emacdpakoihlgkpbekbfnhinbipjbepd"
        #     os.startfile(chromePath)

        elif 'send email to papa' in query:
            try:
                speak("what should i send???")
                content = takeCommand()
                to = "gautamlahkar@gmail.com"
                sendEmailpapa(to, content)
                speak("Email has been send")
            except Exception as e:
                print(e)
                speak("Sorry, the email cant be send.")

        elif 'send email to ma' in query:
            try:
                speak("what should i send???")
                content = takeCommand()
                to = "paramitachakravarty@gmail.com"
                sendEmailma(to, content)
                speak("Email has been send")
            except Exception as e:
                print(e)
                speak("Sorry, the email cant be send.")

        elif 'what is your name' in query:
            speak('Sir My name is Jarvice.. And I am an Artificial Itelligents, I only work for my Sir Anchit Lahkar.')

        elif 'who are you' in query:
            speak('Sir My name is Jarvice.. And I am an Artificial Itelligents I only work for my Sir Anchit Lahkar.')

        elif 'what is python' in query:
            speak('sir please import codes')

        elif 'the date' in query:
            print(ctime())
            speak(ctime())
        elif 'search google' in query:
            speak('what do you want to search for???')
            search = takeCommand('what do you want to search for???')
            url = 'http://google.com/search?q=' + search
            webbrowser.get().open(url)
            speak('Here is what I found for' + search)
            print('Here is what I found for' + search)
         


        elif 'get location' in query:
            speak('What is the location???')
            location = takeCommand('What is the location???')
            url = 'http://google.nl/maps/place/' + location + '/&amp;'
            webbrowser.get().open(url)
            print('Here is the location' + location)

        elif 'answer' in query:
            speak('what do you want to search for???')
            search = takeCommand('what do you want to search for???')
            speak('Searching Wikipedia')
            url = 'https://en.wikipedia.org/wiki/' + search
            webbrowser.get().open(url)
            speak('Here is what I found for' + search)
            print('Here is what I found for' + search)
        
        elif 'search microsoft' in query:
            speak('what do you want to search for???')
            search = takeCommand('what do you want to search for???')
            url = 'https://www.bing.com/search?q=' + search
            webbrowser.get().open(url)
            speak('Here is what I found for' + search)
            print('Here is what I found for' + search)


        elif 'open spotify' in query:
            speak('opening spotify.com')
            url = 'https://open.spotify.com/'
            webbrowser.get().open(url)
            speak('now you can hear music on spotify.')
        


        elif 'open online website' in query:
            speak('opening wix.com')
            url = 'https://www.wix.com/account/sites?utm_source=email_mkt&utm_campaign=em_welcome-to-wix_en&experiment_id=button_cta_1_desktop'
            webbrowser.get().open(url)
            speak('Now you can code websites')

        elif 'talk to me' in query :
            speak('activating selftalk')
            speak("how can i help you through google as search engine")
            search = takeCommand("how can i help you through google as search engine")
            url = 'http://google.com/search?q=' + search
            webbrowser.get().open(url)
            speak("hear's what i found")
            speak(webbrowser.open(url))




        elif 'multiply' in query :
           speak('please enter the value of a\n')
           a = input('what is the value of a\n')
           a = int(a)    

           speak('please enter the value of b\n')
           b = input('what is the value of b\n')
           b = int(b)

           speak('a multiply b equals')
           speak(a * b)
           print(a * b)

        elif 'addition' in query :
            speak('please enter the value of a\n')
            a = input('what is the value of a\n')
            a = int(a)
 
            speak('please enter the value of b\n')
            b = input('what is the value of b\n')
            b = int(b)
 
            speak(a + b)
            print(a + b)

        elif 'subtraction' in query :
           speak('please enter the value of a\n')
           a = input('what is the value of a\n')
           a = int(a)

           speak('please enter the value of b\n')
           b = input('what is the value of b\n')
           b = int(b)

           speak(a * b)
           print(a * b)

        elif 'divide' in query :
           speak('please enter the value of a\n')
           a = input('what is the value of a\n')
           a = int(a)

           speak('please enter the value of b\n')
           b = input('what is the value of b\n')
           b = int(b)

           print(a / b)
           speak(a / b)

        elif 'square' in query :
           speak('please enter the value \n')
           a = input('what is the value \n')
           a = int(a)

           speak(a*a)
           print(a*a)

 



        # elif 'search google' in query:
        #     speak("Searching ....")
        #     query = query.replace("search", "")
        #     for j in search(query, tld="co.in", num=10, stop=10, pause=2): 
        #      print(j)           

        elif 'google password'  in query:
               speak('google account for anchitlahkar0202@gmail.com')
               speak('password = manyata222007')

        elif 'microsoft password' in query:
               speak('microsoft account for anchitlahkar0202@gmail')
               speak('password = anchit222007')

        elif 'second google account password' in query:
               speak('google account for Alindustriescode2020@gmail.com')
               speak('password = anchit222007')


        elif 'shut down' in query:
            speak('Sir if you need any help please ask me, i will be ready to help you')
            exit()
