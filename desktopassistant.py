import pyttsx3
import speech_recognition as sr
import datetime
import random
import wikipedia
import webbrowser
import os
import wolframalpha
import wmi
import pyautogui
import psutil

engine = pyttsx3.init('sapi5')

client = wolframalpha.Client('5YEH5L-X46GJXLQJ2')

voices = engine.getProperty('voices')

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

    speak("I am your Desktop Assistant Jarvis Sir. Please tell me how may I help you")

def takeCommand():


    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)

    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language='en-in')
        print(f"User said: {query}\n")

    except Exception as e:
        print("Say that again please...")
        query=takeCommand()
        return query
    return query


def secs2hours(secs):
    mm, ss = divmod(secs, 60)
    hh, mm = divmod(mm, 60)
    return "%dhour, %02d minute, %02s seconds" % (hh, mm, ss)




if __name__ == "__main__":
    wishMe()
    while True:

        query = takeCommand().lower()


        if 'wikipedia' in query:
            speak('Searching Wikipedia...')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=2)
            speak("According to Wikipedia")
            print(results)
            speak(results)

        elif 'decrease brightness' in query:
            dec=wmi.WMI(namespace='wmi')
            methods=dec.WmiMonitorBrightnessMethods()[0]
            methods.WmiSetBrightness(30,0)

        elif 'increase brightness' in query:
            inc=wmi.WMI(namespace='wmi')
            methods=inc.WmiMonitorBrightnessMethods()[0]
            methods.WmiSetBrightness(100,0)

        elif 'charge' in query:
            battery = psutil.sensors_battery()
            plugged = battery.power_plugged
            percent = int(battery.percent)
            time_left = secs2hours(battery.secsleft)
            print(percent)
            if percent < 40 and plugged == False:
                speak('sir, please connect charger because i can survive only ' + time_left)
            if percent < 40 and plugged == True:
                speak("don't worry, sir charger is connected")
            else:
                speak('sir, no need to connect the charger because i can survive ' + time_left)

        elif 'please remind' in query:
            global remind_speech
            speak('what should i remind?')
            remind_speech = takeCommand()
            speak('Reminded:' + remind_speech)

        elif 'any reminder' in query:
            if remind_speech is None:
                speak('you do not have any reminder for today')
            else:
                speak('you have one reminder'+remind_speech)

        elif 'screenshot' in query or 'snapshot' in query:
            speak('ok, sir let me take a snapshot ')
            speak('ok done')
            speak('check your desktop, i saved there')
            pic = pyautogui.screenshot()
            pic.save('C:/Users/Omkar/Desktop/Screenshot.png')

        elif 'open youtube' in query:
            webbrowser.open("youtube.com")

        elif 'open google' in query:
            webbrowser.open("google.com")

        elif 'open website' in query:
            speak("Which website sir")
            w=takeCommand()
            webbrowser.open(w)

        elif 'open stackoverflow' in query:
            webbrowser.open("stackoverflow.com")

        elif 'play music' in query:
            speak("What you want to play anything,or a specific song")
            ch=takeCommand()
            if 'specific' in ch:
                music_dir = 'E:\\All songs\\DJ Songs\\Desktop Assistant Songs'
                speak("Which song sir")
                s=takeCommand()
                speak("Playing "+s+"for you")
                os.startfile(os.path.join(music_dir,s+'.mp3'))
            else:
                music_dir = 'E:\\All songs\\DJ Songs\\Hindi Dance'
                songs = os.listdir(music_dir)
                print(songs)
                os.startfile(os.path.join(music_dir, random.choice(songs)))

        elif 'the time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"Sir, the time is {strTime}")

        elif 'quit jarvis' in query:
            speak("Goodbye Sir See you again")
            exit(0)

        elif 'hello' in query:
            speak('Hello Sir')

        elif "what\'s up" in query or 'how are you' in query:
            stMsgs = ['Just doing my thing!', 'I am fine!', 'Nice!', 'I am nice and full of energy']
            speak(random.choice(stMsgs))

        else:
            query = query
            speak('Searching...')
            try:
                try:
                    res = client.query(query)
                    results = next(res.results).text
                    speak('Got it.')
                    speak('WOLFRAM-ALPHA says - ')
                    print(results)
                    speak(results)

                except:
                    results = wikipedia.summary(query, sentences=2)
                    speak('Got it.')
                    speak('WIKIPEDIA says - ')
                    print(results)
                    speak(results)

            except:
                speak("Sorry sir no results found")
