from flask import Flask, render_template, request, redirect
import subprocess
import wolframalpha
import pyttsx3
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import os
import winshell
import pyjokes
import smtplib
import ctypes
import time
import requests
import shutil
from twilio.rest import Client
from clint.textui import progress
engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)
app = Flask(__name__)
def speak(audio):
    engine.say(audio)
    # engine.runAndWait()

def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour < 12:
        speak("Good Morning Sir. I am Your Assistant Nova. How May i Help You ?")

    elif hour >= 12 and hour < 18:
        speak("Good Afternoon Sir. I am Your Assistant Nova. How May i Help You ?")

    else:
        speak("Good Evening Sir. I am Your Assistant Nova. How May i Help You ?")

def username():
    speak("What should i call you sir")
    username = takeCommand()
    speak("Welcome Mister")
    speak(username)
    columns = shutil.get_terminal_size().columns

    print("#####################".center(columns))
    print("Welcome Mr.", username.center(columns))
    print("#####################".center(columns))

    speak("How can i Help you, Sir")


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
        print(e)
        print("Unable to Recognize your voice.")
        return "None"

    return query


def sendEmail(to, content):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    server.login('your email id', 'your email password')
    server.sendmail('your email id', to, content)
    server.close()


@app.route("/", methods=["GET", "POST"])
def index():
    transcript = ""
    if request.method == "POST":
        clear = lambda: os.system('cls')
        clear()
        while True:
            query = takeCommand().lower()
            if 'wikipedia' in query:
                speak('Searching Wikipedia...')
                query = query.replace("wikipedia", "")
                results = wikipedia.summary(query, sentences=3)
                speak("According to Wikipedia")
                print(results)
                speak(results)

            elif 'open youtube' in query:
                speak("Here you go to Youtube\n")
                webbrowser.open("youtube.com")

            elif 'open google' in query:
                speak("Here you go to Google\n")
                webbrowser.open("google.com")

            elif 'open stackoverflow' in query:
                speak("Here you go to Stack Over flow.Happy coding")
                webbrowser.open("stackoverflow.com")

            elif 'play music' in query or "play song" in query:
                speak("Here you go with music")
                music_dir = "E:\Music"
                songs = os.listdir(music_dir)
                print(songs)
                random = os.startfile(os.path.join(music_dir, songs[1]))


            elif 'time' in query:
                strTime = datetime.datetime.now().strftime("%H:%M:%S")
                speak(f"the time is {strTime}")

            elif 'open google chrome' in query:
                codePath = r"C:\Program Files\Google\Chrome\Application"
                os.startfile(codePath)


            elif 'email to sriharsha' in query:
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
                    speak("whom should i send")
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
                assist_name = query

            elif "change name" in query:
                speak("What would you like to call me, Sir ")
                assist_name = takeCommand()
                speak("Thanks for naming me")

            elif "what's your name" in query or "What is your name" in query:
                speak("My friends call me")
                speak(assist_name)
                print("My friends call me", assist_name)

            elif 'exit' in query:
                speak("Thanks for giving me your time")
                exit()

            elif "who made you" in query or "who created you" in query:
                speak("I have been created by Sriharsha.")

            elif 'joke' in query:
                speak(pyjokes.get_joke())

            elif "calculate" in query:

                app_id = "Wolframalpha api id"
                client = wolframalpha.Client(app_id)
                indx = query.lower().split().index('calculate')
                query = query.split()[indx + 1:]
                res = client.query(' '.join(query))
                answer = next(res.results).text
                print("The answer is " + answer)
                speak("The answer is " + answer)

            elif 'search' in query or 'play' in query:

                query = query.replace("search", "")
                query = query.replace("play", "")
                webbrowser.open(query)

            elif "who i am" in query:
                speak("If you talk then definitely your human.")

            elif "why you came to world" in query:
                speak("Thanks to Sriharsha. further It's a secret")

            elif 'power point presentation' in query:
                speak("opening Power Point presentation")
                power = r"C:\\Users\\SRIHARSHA\\Desktop\\Minor Project\\Presentation\\Voice Assistant.pptx"
                os.startfile(power)

            elif "open google" in query:
                speak("Opening Google ")
                webbrowser.open("www.google.com")


            elif ' what is love' in query:
                speak("It is 7th sense that destroy all other senses")

            elif "who are you" in query:
                speak('I am Nova version 1 point O your personal assistant. I am programmed to minor tasks like''opening youtube,google chrome,gmail and stackoverflow ,predict time,take a photo,search wikipedia,predict weather' 
                    'in different cities , get top headline news from times of india and you can ask me computational or geographical questions too!')

            elif 'reason for you' in query:
                speak("I was created as a Minor project by Mister Sriharsha ")

            elif 'change background' in query:
                ctypes.windll.user32.SystemParametersInfoW(20,0,"Location of wallpaper",0)
                speak("Background changed successfully")

            elif 'news' in query:
                news = webbrowser.open_new_tab("https://timesofindia.indiatimes.com/home/headlines")
                speak('Here are some headlines from the Times of India,Happy reading')


            elif 'lock window' in query:
                speak("locking the device")
                ctypes.windll.user32.LockWorkStation()

            elif 'shutdown system' in query:
                speak("Hold On a Sec ! Your system is on its way to shut down")
                subprocess.call('shutdown / p /f')

            elif 'empty recycle bin' in query:
                winshell.recycle_bin().empty(confirm=False, show_progress=False, sound=True)
                speak("Recycle Bin Recycled")

            elif "don't listen" in query or "stop listening" in query:
                speak("for how much time you want to stop Nova from listening commands")
                a = int(takeCommand())
                time.sleep(a)
                print(a)

            elif "where is" in query:
                query = query.replace("where is", "")
                location = query
                speak("User asked to Locate")
                speak(location)
                webbrowser.open("https://www.google.nl / maps / place/" + location + "")


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
                file = open("nova.txt", "r")
                print(file.read())
                speak(file.read(6))

            elif "update assistant" in query:
                speak("After downloading file please replace this file with the downloaded one")
                url = '# url after uploading file'
                r = requests.get(url, stream=True)

                with open("Voice.py", "wb") as Pypdf:

                    total_length = int(r.headers.get('content-length'))

                    for ch in progress.bar(r.iter_content(chunk_size=2391975),
                                        expected_size=(total_length / 1024) + 1):
                        if ch:
                            Pypdf.write(ch)


            elif "Nova" in query:

                wishMe()
                speak("Nova 1 point o in your service Mister")
                speak(assist_name)

            elif "weather" in query:

                api_key = "8ef61edcf1c576d65d836254e11ea420"

                base_url = "https://api.openweathermap.org/data/2.5/weather?"

                speak("whats the city name")

                city_name = takeCommand()

                complete_url = base_url + "appid=" + api_key + "&q=" + city_name

                response = requests.get(complete_url)

                x = response.json()

                if x["cod"] != "404":

                    y = x["main"]

                    current_temperature = y["temp"]

                    current_humidity = y["humidity"]

                    z = x["weather"]

                    weather_description = z[0]["description"]

                    speak(" Temperature in kelvin unit is " +

                        str(current_temperature) +

                        "\n humidity in percentage is " +

                        str(current_humidity) +

                        "\n description  " +

                        str(weather_description))

                    print(" Temperature in kelvin unit = " +

                        str(current_temperature) +

                        "\n humidity (in percentage) = " +

                        str(current_humidity) +

                        "\n description = " +

                        str(weather_description))


                else:

                    speak(" City Not Found ")

            elif "send message " in query:
                account_sid = 'Account Sid key'
                auth_token = 'Auth token'
                client = Client(account_sid, auth_token)

                message = client.messages \
                    .create(
                    body=takeCommand(),
                    from_='Sender No',
                    to='Receiver No'
                )

                print(message.sid)

            elif "wikipedia" in query:
                webbrowser.open("wikipedia.com")

            elif "Good Morning" in query:
                speak("A warm" + query)
                speak("How are you Mister")
                speak(assist_name)

            elif "will you be my gf" in query or "will you be my bf" in query:
                speak("I'm not sure about, may be you should give me some time")

            elif "how are you" in query:
                speak("I'm fine, thank you")

            elif "i love you" in query:
                speak("It's hard to understand")

            elif "what is" in query or "who is" in query:
                client = wolframalpha.Client("API_ID")
                res = client.query(query)

                try:
                    print(next(res.results).text)
                    speak(next(res.results).text)
                except StopIteration:
                    print("No results")
    return render_template('index.html', transcript=transcript)


if __name__ == "__main__":
    app.run(debug=True, threaded=True)