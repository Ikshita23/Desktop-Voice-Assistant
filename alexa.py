
import pyttsx3
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import os
import random
import pyjokes
from ecapture import ecapture as ec
import winshell
import ctypes


engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
# print(voices[1].id)
engine.setProperty('voice', voices[1].id)


def speak(audio):
    engine.say(audio)
    engine.runAndWait()


def wishMe():
    hour = int(datetime.datetime.now().hour)
    if(hour >= 0 and hour < 12):
        speak("Good Morning!")

    elif(hour >= 12 and hour < 18):
        speak("Good Afternoon!")

    else:
        speak("Good Evening!")

    speak("Hey, I am your Assistant, ALEXA")
    speak("How may I help you?")

def tellDay():
     
    # This function is for telling the day of the week
    day = datetime.datetime.today().weekday() + 1

    Day_dict = {1: 'Monday', 2: 'Tuesday',
                3: 'Wednesday', 4: 'Thursday',
                5: 'Friday', 6: 'Saturday',
                7: 'Sunday'}
     
    if day in Day_dict.keys():
        day_of_the_week = Day_dict[day]
        print(day_of_the_week)
        speak("The day is " + day_of_the_week)

def takeCommand():
    #It takes microphone input from the user and returns string output
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)

    try:     #in case of error
        print("Recognizing...")
        query = r.recognize_google(audio, language = 'en-in')
        print(f"User said: {query}\n")
    
    except Exception as e:
        print(e)
        print("Unable to recognize your voice")
        return "None"

    return query


if __name__ == "__main__":
    wishMe()

    while True:
        query = takeCommand().lower()

        # Logic for executing tasks based on query

        if 'how are you' in query:
            speak("I am fine, Thank you\n How are you?")

        elif 'fine' in query or 'good' in query or 'great' in query:
            speak("I am glad to know that")
            
        elif 'wikipedia' in query:
            speak('Searching Wikipedia...')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences = 2)
            speak('Alright!')
            speak("According to Wikipedia")
            print(results)
            speak(results)

        elif 'open youtube' in query:
            speak("Opening Youtube\n")
            webbrowser.open("youtube.com")
        
        elif 'open google' in query:
            speak("Opening Google\n")
            webbrowser.open("google.com")

        elif 'open stack overflow' in query:
            speak("Opening Stackoverflow\n")
            webbrowser.open("stackoverflow.com")

        elif 'play music' in query or 'play songs' in query:
            music_dir =  'D:\\voice assistant\\alexa voice assistant\\music'
            songs = os.listdir(music_dir)
            print(songs)
            os.startfile(os.path.join(music_dir, random.choice(songs)))

        elif 'the time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"The time is {strTime}")

        elif "which day it is" in query:
            tellDay()
            continue

        elif 'open vs code' in query:
            speak("Opening vs code\n")
            codePath = "C:\\Users\\Lenovo\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"
            os.startfile(codePath)

        elif 'open word' in query:
            speak("Opening Microsoft Word\n")
            codePath = "C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs\\Word.lnk"
            os.startfile(codePath)

        elif 'open excel' in query:
            speak("Opening Microsoft Excel\n")
            codePath = "C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs\\Excel.lnk"
            os.startfile(codePath)

        elif 'quit' in query or 'exit' in query:
            speak("Bye-Bye, Thanks for your time")
            exit()

        elif 'shut down' in query:
            check = input("Do you want to shutdown your computer ? (y/n): ")
            if check == 'y' or check == "Y":
                os.system("shutdown /s /t 1")

        elif 'restart' in query:
            check = input("Do you want to restart your computer ? (y/n): ")
            if check == 'y' or check == "Y":
                os.system("shutdown /r /t 1")

        elif 'who made you' in query or 'who created you' in query:
            speak("I was created as a project by IKSHITA")

        elif 'change background' in query:
            ctypes.windll.user32.SystemParametersInfoW(20,0,"D:\\voice assistant\\alexa voice assistant",0)
            speak("Background changed successfully")

        elif 'joke' in query:
            speak(pyjokes.get_joke())

        elif "write a note" in query:
            speak("What should i write?")
            note = takeCommand()
            file = open('note.txt', 'w')
            speak("Should i include date and time")
            snfm = takeCommand()
            if 'yes' in snfm or 'sure' in snfm:
                strTime = datetime.datetime.now().strftime("%H:%M:%S")
                file.write(strTime)
                file.write(" :- ")
                file.write(note)
            else:
                file.write(note)
         
        elif "show note" in query:
            speak("Showing Notes")
            file = open("note.txt", "r")
            print(file.read())
            speak(file.read(6))

        elif "camera" in query or "take a photo" in query:
            ec.capture(0, "Alexa Camera ", "img.jpg")

        elif 'empty recycle bin' in query:
            winshell.recycle_bin().empty(confirm = False, show_progress = False, sound = True)
            speak("Recycle Bin Recycled")


        else:
            speak("Application not available")
