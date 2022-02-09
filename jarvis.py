import ctypes
import datetime
import time
import speech_recognition as sr
import pyttsx3
import wikipedia
import webbrowser
import os
import smtplib
import pyjokes
import winshell
from ecapture import ecapture as ec
import wolframalpha
import requests
import subprocess
import json
from twilio.rest import Client


engine=pyttsx3.init('sapi5')
voices=engine.getProperty('voices')
engine.setProperty('voice',voices[0].id)


def speak(audio):
    engine.say(audio)
    engine.runAndWait()

def wishMe():
    hour=int(datetime.datetime.now().hour)
    if hour>=0 and hour<12:
        speak("Good Morning")
    elif hour>=12 and hour<18:
        speak("Good Afternoon")
    else:
        speak("Good Evening")

    speak("I am Jarvis madam,Please tell me how may I help you")
name="jarvis"


def take_command():
    r=sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold=1
        audio=r.listen(source)

    try:
        print("Recognizing....")
        query=r.recognize_google(audio,language='en-in')
        print(f"user said:{query}\n")


    except Exception as e:
        #print(e)
        print("Say that again please..")
        return "None"
    return query

def format_sentence(sentence):
    sentenceSplit=filter(None,sentence.split("."))
    for s in sentenceSplit:
        print(s.strip()+".")

def sendEmail(to,content):
    server=smtplib.SMTP('smtp.gmail.com',587)
    server.ehlo()
    server.starttls()
    server.login('prateekshakumtakar3750@gmail.com','prateeksha@123')
    server.sendmail('prateekshakumtakar3750@gmail.com',to,content)
    server.close()

if __name__ == "__main__":
    wishMe()
    while(True):
        query=take_command().lower()

        if 'wikipedia' in query:
            speak("Searching wikipedia....")
            print("Searching wikipedia...")
            query=query.replace("wikipedia","")
            results=wikipedia.summary(query,sentences=2)
            speak("according to wikipedia")
            format_sentence(results)
            speak(results)

        elif 'open youtube' in query:
            print("Opening Youtube..")
            webbrowser.open("youtube.com")

        elif 'open google' in query:
            print("Opening Google..")
            webbrowser.open("google.com")

        elif 'open stackoverflow' in query:
            print("Opening Youtube..")
            webbrowser.open("stackoverflow.com")

        elif 'play music' in query:
            music_dir="C:\\Users\\prateeksha\\Music\\mymusic"
            songs=os.listdir(music_dir)
            print(songs)
            os.startfile(os.path.join(music_dir,songs[1]))

        elif 'the time' in query:
            strTime=datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"Madam, the time is {strTime}")
            print(f"The time is {strTime}")

        elif 'open a code' in query:
            codePath="C:\\Users\\prateeksha\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"
            os.startfile(codePath)

        elif 'send email to' in query:
            try:
                speak("What should I say?")
                content=take_command()
                to="1da18cs112.cs@drait.edu.in"
                sendEmail(to,content)
                speak("Email has been sent")
                print("Email has been sent")
            except Exception as e:
                print(e)
                speak("Sorry my friend,I am not able to send this mail")

        elif 'joke' in query:
            jokes = pyjokes.get_joke()
            speak(jokes)
            print(jokes)

        elif 'news' in query:
            news=webbrowser.open_new_tab("https://timesofindia.com/home/headlines")
            speak("here are some headlines from the Times of India,happy reading")
            print("here are some headlines from the Times of India,happy reading")
            print("Happy reading!")

        elif "camera" in query or "take a photo" in query:
            ec.capture(0,"robo camera","img.jpg")


        elif  "ask" in query:  #its an api which can compute expert-level answers using Wolfram's algorithms,knowledgebase and AI technology
            speak("I can answer to compuational and geographical questions and what question do you want to ask now?")
            question=take_command()
            app_id="L49WA5-XXH8V56U47"
            client=wolframalpha.Client('L49WA5-XXH8V56U47')
            res=client.query(question)
            answer=next(res.results).text
            speak(answer)
            print(answer)


        elif "weather" in query:
            api_key="f013661fe38bafb9af920993e02d1749"
            base_url="https://api.openweathermap.org/data/2.5/weather?"
            speak("what is your city name?")
            city_name=take_command()
            complete_url=base_url+"appid="+api_key+"&q="+city_name
            response=requests.get(complete_url)
            x=response.json()
            if x["cod"]!=404:
                y=x["main"]
                current_temperature=y["temp"]
                current_humidity=y["humidity"]
                z=x["weather"]
                weather_description=z[0]["description"]
                speak("temperature in kelvin unit is"+
                      str(current_temperature)+
                      "\n humidity in percentage is"+
                      str(current_humidity)+
                      "\n description "+
                      str(weather_description))
                print("temperature in kelvin unit is" +
                      str(current_temperature) +
                      "\n humidity in percentage is" +
                      str(current_humidity) +
                      "\n description " +
                      str(weather_description))



        elif "change name" in query:
            speak("What would you like to call me madam?")
            print("What would you like to call me madam?")
            name=take_command()
            speak("thanks for naming me")
            print("thanks for naming me")



        elif "powerpoint presentation" in query:
            speak("opening power point presentation")
            power=r"C:\\Users\\prateeksha\\Desktop\\Project and Activities\\Iot.pptx"
            os.startfile(power)

        elif "change background" in query:
            ctypes.windll.user32.SystemParametersInfoW(20,0,"Location of walpaper",0)
            speak("Background changed successfully")

        elif "lock window" in query:
            speak("locking the device")
            ctypes.windll.user32.LockWorkStation()

        elif "shutdown system" in query:
            speak("Hold on a sec,your system is on it's way to shut down")
            subprocess.call('shutdown/p/f')

        elif "empty recycle bin" in query:
            winshell.recycle_bin().empty(confirm=False,show_progress=False,sound=True)
            speak("Recycle bin recycled")

        elif "where is" in query:
            query=query.replace("where is","")
            location=query
            speak("user asked to locate")
            speak(location)
            webbrowser.open("https://www.google.nl/maps/place/"+location+"")

        elif "restart" in query:
            subprocess.call(["shutdown","/r"])

        elif ("hibernate" or "sleep") in query:
            speak("Hibernating")
            subprocess.call("shutdown / h")

        elif ("log off" or "sign out") in query:
            speak("ok, your pc will log off in 10 sec make sure you exit from all applications")
            time.sleep(5)
            subprocess.call(["shutdown","/1"])

        elif "don't listen" in query or "stop listening" in query:
            speak("for how much time you want me to stop from listening commands")
            print("for how much time you want me to stop from listening commands")
            a=int(take_command())
            time.sleep(a)
            print(a)

        elif "write a note" in query:
            speak("what should I write,madam")
            note=take_command()
            file=open("jarvis.txt","w")
            speak("sir,should I include date and time?")
            snfm=take_command()
            if "yes" or "sure" in snfm:
                strTime=datetime.datetime.now().strftime("%H:&M:%S")
                file.write(strTime)
                file.write(" :- ")
                file.write(note)
            else:
                file.write(note)

        elif "show note" in query:
            speak("showing notes")
            file=open("jarvis.txt","r")
            print(file.read())
            speak(file.read(6))

        elif "send message" in query:
            account_sid='AC7147c6e2df93ecd962778dc0a0162df8'
            auth_token='3aa5d8562ecdde672d80860f812'
            client=Client(account_sid,auth_token)
            message=client.messages.create(body=take_command(),from_="7022928060",to="9448093136")
            print(message.sid)

        elif "who are you" in query or "what can you do" in query:
            speak("I am your personal assistant. I am programmed to minor tasks like"
                  "opening youtube,google chrome,gmail and stackoverflow,predict time,take a photo,search wikipedia,predict weather"
                  "in different cities ,get top headline news from times of india and you can ask me computational or geographical questions too!")
            print("I am your personal assistant. I am programmed to minor tasks like\n"
                  "opening youtube,google chrome,gmail and stackoverflow,predict time,take a photo,search wikipedia,predict weather\n"
                  "in different cities ,get top headline news from times of india and you can ask me computational or geographical questions too!\n")


        elif ("what's your name" in query)or("what is your name" in query):
            print(name)
            speak(name)


        elif "how are you" in query:
            speak("I am fine,Thank you")
            speak("How are you madam?")
            print("I am fine,Thank you")
            print("How are you madam?")

        elif "fine" in query or "good" in query:
            speak("It's good to know that you are fine")
            print("It's good to know that you are fine")


        elif "exit" in query:
            speak("Thanks for giving me your time")
            exit()










