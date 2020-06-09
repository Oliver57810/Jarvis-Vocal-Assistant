from __future__ import print_function
import datetime
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import os
import time
import playsound
import pytz
import pyttsx3
import pyttsx3.drivers.sapi5
import speech_recognition as sr
import subprocess
import random
import webbrowser
from dateutil import parser
import urllib.request
import json     
import pyowm
import glob
from googlesearch import search
import MySQLdb
from tqdm.auto import tqdm
from time import sleep
import wikipedia


SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']
MONTHS = ["january", "february", "march", "april", "may", "june","july", "august", "september","october", "november", "december"]
DAYS = ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"]
DAY_EXTENTIONS = ["rd", "th", "st", "nd"]
QUICKSTART = [r"C:\Program Files\Sublime Text 3\sublime_text.exe", r"D:\Visual Studio\Common7\IDE\devenv.exe", r"C:\Users\olive\AppData\Local\Programs\Microsoft VS Code\Code.exe", r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe", r"C:\Program Files\NetBeans 8.2 RC\bin\netbeans64.exe", r"C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE", r"C:\Program Files (x86)\Microsoft Office\root\Office16\POWERPNT.EXE", r"C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE", r"C:\Program Files (x86)\Minecraft Launcher\MinecraftLauncher.exe", r"D:\Programmi\Origin\Origin.exe", r"D:\Programmi\Steam\Steam.exe", r"D:\Programmi\Ubisoft Game Launcher\Uplay.exe", r"D:\Call of Duty Warzone\Call of Duty Modern Warfare\ModernWarfare.exe", r"D:\GTA V\GTAVLauncher.exe"] 
CITY = ["Vicenza"]
USER = ["User:"]
PASSWORD = ["Jarvis"]



def speak(text):
    engine = pyttsx3.init()
    en_voice_id = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\Voices\Tokens\TTS_MS_EN-GB_HAZEL_11.0"
    engine.setProperty('voice', en_voice_id)
    engine.setProperty('rate', 130)
    engine.setProperty('volume', 1.2)

    print("Jarvis: " + text)
    engine.say(text)
    engine.runAndWait()

def get_wakeAudio():
	r = sr.Recognizer()
	with sr.Microphone() as source:
		audio = r.listen(source)
		said = ""
		try:
			said = r.recognize_google(audio)
		except Exception as e:
			return("Exception" + str(e))

	return said.lower()

def get_audio():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        audio = r.listen(source)
        said = ""
        try:
            said = r.recognize_google(audio)
            print("\n" + USER[0], said + "\n")
        except Exception as e:
            speak("I'm sorry I haven't heard you")
            return("Exception" + str(e))

    return said.lower()

def authenticate_google():
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('calendar', 'v3', credentials=creds)

    return service

def get_events(date, service):
    # Call the Calendar API
    date = datetime.datetime.combine(date, datetime.datetime.min.time())
    end_date = datetime.datetime.combine(date, datetime.datetime.max.time())
    utc = pytz.UTC
    date = date.astimezone(utc)
    end_date = end_date.astimezone(utc)


    events_result = service.events().list(calendarId='primary', timeMin=date.isoformat(), timeMax=end_date.isoformat(),
                                        singleEvents=True,
                                        orderBy='startTime').execute()
    events = events_result.get('items', [])

    if not events:
        speak('No upcoming events found.')
    else:
        speak(f"You have {len(events)} events on this day.")

        for event in events:
            start = event['start'].get('dateTime', event['start'].get('date'))
            start_time = str(start.split("T")[1].split("+")[0])
            if int(start_time.split(":")[0]) < 12:
                start_time = start_time + "am"
            else:
                start_time = str(int(start_time.split(":")[0])-12) + start_time.split(":")[1]
                start_time = start_time + "pm"

            speak(event["summary"] + " at " + start_time)

def get_date(text):
	today = datetime.date.today()

	if text.count("today") > 0:
		return today

	day = -1	
	day_of_week = -1
	month = -1
	year = today.year

	for word in text.split():
		if word in MONTHS:
			month = MONTHS.index(word) + 1
		elif word in DAYS:
			day_of_week = DAYS.index(word)
		elif word.isdigit():
			day = int(word)
		else:
			for ext in DAY_EXTENTIONS:
				found = word.find(ext)
				if found > 0:
					try:
						day = int(word[:found])
					except:
						pass

	if month < today.month and month != -1:
		year = year + 1

	if month == -1 and day != -1:
		if day < today.day:
			month = today.month + 1
		else:
			month = today.month

	if month == -1 and day == -1 and day_of_week != -1:
		current_day_of_week = today.weekday()
		dif = day_of_week - current_day_of_week

		if dif < 0:
			dif += 7
			if text.count("next") >= 1:
				dif += 7

		return today + datetime.timedelta(dif)

	if month == -1 or day == -1:
		return None


	return datetime.date(month = month, day = day, year = year)

def note(text):
    date = datetime.datetime.now()
    file_name = str(date).replace(":", "-") + "-note.txt"
    with open(file_name, "w") as f:
        f.write(text)

    subprocess.Popen(["notepad.exe", file_name])

def getCurrentWeather(): 
    owm = pyowm.OWM("your API key")
    place = str(CITY[0]) + ", IT"
    obs = owm.weather_at_place(place)
    weather = obs.get_weather()
    temperature = weather.get_temperature('celsius')['temp']
    temperrature = round(temperature, 1)
    status = weather.get_status()
    forecast = "Current weather conditions in Vicenza: " + str(status) + ". The temperature is about: " + str(temperature) + "degrees"
    speak(forecast)

def wiki(text):
    data = wikipedia.search(text)
    data.reverse()
    temp = data.pop()
    data.reverse()
    if len(data) > 1:
        speak("There are multiple results. This is the list of them, each result is separated by a comma. Please type the number of what you meant.")
        print(data)
        pick = input("Choose what you meant as " + text + ". Insert a number from the list:")
        pick = int(pick) - 1
        selection = str(data[pick])
        if selection != None:
            speak("According to Wikipedia: " + wikipedia.summary(selection, sentences = 3))
        else:
            speak("Sorry, I think you have typed a wrong number.")
    elif len(data) == 1:
        speak("According to Wikipedia: " + wikipedia.summary(selection, sentences = 3))
    else:
        speak("Something went wrong!")

def locate(lat, lon):
    speak("Based on Wikipedia geosearch here we have the most important places near that coordinates")
    print(wikipedia.geosearch(lat, lon))

def programIDE(text):
    if "c sharp" in text:
        speak("Fine, I'm opening Visual Studio 2019") 
        subprocess.Popen([r"D:\Visual Studio\Common7\IDE\devenv.exe"])
    elif "python" in text:
        speak("Ok, I'm opening Sublime Text 3")
        subprocess.Popen([r"C:\Program Files\Sublime Text 3\sublime_text.exe"])
    else:
        speak("In this case Visual Studio Code should be the right choice")
        subprocess.Popen([r"C:\Users\olive\AppData\Local\Programs\Microsoft VS Code\Code.exe"])

def googling(text):
    url = f"https://www.google.com.tr/search?q={text}"
    webbrowser.open(url)

def add_event(date_text, start_Time, end_Time, title):
    sTimeStr = ""
    eTimeStr = ""
    sBool = False;
    eBool = False;
    date_text = get_date(date_text)
    dateStr = date_text.strftime("%Y%m%d")
    for word in start_Time.split(":"):
        if word.isdigit():
            sTimeStr = sTimeStr + word
            sBool = True;

    if sBool != True:
        if start_Time.count("a.m.") > 0:
            sTimeStr = ""
            for word in start_Time.split():
                if word.isdigit():
                    sTimeStr = word

        elif start_Time.count("p.m.") > 0:
            sTimeStr = ""
            for word in start_Time.split():
                if word.isdigit():
                    intTime = int(word) + 12
                    sTimeStr = str(intTime) 

        sTimeStr = sTimeStr + "00"

    if len(sTimeStr) < 4:
        sTimeStr = "0" + sTimeStr

    for word in end_Time.split(":"):
        if word.isdigit():
            eTimeStr = eTimeStr + word
            eBool = True;

    if eBool != True:
        if end_Time.count("a.m.") > 0:
            eTimeStr = ""
            for word in end_Time.split():
                if word.isdigit():
                    eTimeStr = word

        elif end_Time.count("p.m.") > 0:
            eTimeStr = ""
            for word in end_Time.split():
                if word.isdigit():
                    intTime = int(word) + 12
                    eTimeStr = str(intTime) 

        eTimeStr = eTimeStr + "00"

    if len(sTimeStr) < 4:
        sTimeStr = "0" + sTimeStr

    start = dateStr + "T" + sTimeStr + "00"
    end = dateStr + "T" + eTimeStr + "00"
    link = f"https://www.google.com/calendar/event?action=TEMPLATE&dates={start}%2F{end}&text={title}"
    webbrowser.open(link)

def covid():
    with urllib.request.urlopen("https://raw.githubusercontent.com/pcm-dpc/COVID-19/master/dati-json/dpc-covid19-ita-andamento-nazionale.json") as url:
        data = json.loads(url.read().decode())
        if data != None: 
            date = str(data[-1]["data"])
            date = date[0] + date[1] + date[2] + date[3] + date[4] + date[5] + date[6] + date[7] + date[8] + date[9]                    
            total_positives = str(data[-1]["totale_positivi"])
            intensive_care = str(data[-1]["terapia_intensiva"])
            new_cases = str(data[-1]["variazione_totale_positivi"])
            healed = str(data[-1]["dimessi_guariti"])
            total_deaths = str(data[-1]["deceduti"])
            print("\n\n" + date + "\nTotal positives: " + total_positives + "\nIntensive care: "  + intensive_care + "\nHealed: " + healed + "\nTotal deaths: " + total_deaths + "\nNew cases: " + new_cases + " This represents the difference between new positives and healed + dead\n\n")
            speak("Data provided by Protezione Civile. First of all the difference between yesterday and today is: " + new_cases + ". Secondly the number of healed reaches: " + healed + ". The number of the total currently positives to covid19 are: " + total_positives + ". Still recovered in intensive care: " + intensive_care + ". Total deaths are: " + total_deaths)
            initialize()
        else:
            # Connessione al Database
            db = MySQLdb.connect("localhost","root","","covid")
            # Ottenimento del cursore
            cursor = db.cursor()
            # Esecuzione di una query SQL
            sql = "SELECT data, SUM(R_Sintomi) as ricoverati_con_sintomi, SUM(T_Intensiva) as terapia_intensiva, SUM(I_Domiciliare) as isolamento_domiciliare, SUM(guariti) as dimessi_guariti, SUM(deceduti) as deceduti, SUM(tamponi) as tamponi from datiregionitot group by data;"
            try:
                # Esecuzione della query
                cursor.execute(sql)
                # Risultati della query
                results = cursor.fetchall()
                date = str(results[len(results) - 1][0])
                p1 = results[len(results) - 1][1]
                p2 = results[len(results) - 1][2]
                p3 = results[len(results) - 1][3]
                tp = p1 + p2 + p3
                total_positives = str(tp)
                intensive_care = str(p2)
                healed = str(results[len(results) - 1][4])
                total_deaths = str(results[len(results) - 1][5])
                new_cases = "345"
            except:
                print("Errore")
                db.close()

            sql2 = "SELECT data, SUM(R_Sintomi + T_Intensiva + I_Domiciliare) as variazione_positivi from dati_regioni GROUP BY data"

            try:
                # Esecuzione della query
                cursor.execute(sql2)
                # Risultati della query
                results = cursor.fetchall()
                new_cases = str(results[len(results) - 1][1])
            except:
                print("Errore")
                db.close()

            print("\n\n" + date + "\nTotal positives: " + total_positives + "\nIntensive care: "  + intensive_care + "\nHealed: " + healed + "\nTotal deaths: " + total_deaths + "\nNew cases: " + new_cases  + " This represents the difference between new positives and healed + dead\n\n")
            speak("Data provided by Protezione Civile. First of all the difference between yesterday and today is: " + new_cases + ". Secondly the number of healed reaches: " + healed + ". The number of the total currently positives to covid19 are: " + total_positives + ". Still recovered in intensive care: " + intensive_care + ". Total deaths are: " + total_deaths)

def news():
    query = "notizie italia"
    my_results_list = []
    for i in search(query,          
                    tld = 'com',    
                    lang = 'it',    
                    num = 5,        
                    start = 0,      # First result to retrieve
                    stop = 5,       # Last result to retrieve
                    pause = 1.0,    # Lapse between HTTP requests
                   ):
        my_results_list.append(i)
        
    for link in my_results_list:
        webbrowser.open(link)

def cls():
    print("\n" * 100)

#main methods
def loading():
    for _ in tqdm(range(100), unit = " Operations", desc = "Loading packages..."):
        sleep(0.1)
    authenticateName()

def authenticateName():
    speak("Please insert your username")
    temp = input("Username:")       
    USER[0] = temp + ":"
    for _ in tqdm(range(50), unit = " Operations", desc = "Checking data..."):
        sleep(0.1)
    authenticate()

def authenticate():
    speak("Please insert your password")
    password = input("Password:")    
    for _ in tqdm(range(50), unit = " Operations", desc = "Checking data..."):
        sleep(0.1)
    if password == PASSWORD[0]:
        cls()
        print("Just A Rather Very Intelligent System, by Oliver Urbani ITIS A. Rossi 2020®          Beta Version: 27.0520a")
        speak("Welcome back!")
        initialize()
    else:
        speak("Password is wrong.")
        authenticate()

def initialize():
    WAKE = "jarvis"

    print("Waiting the keyword")

    while True:
        text = get_wakeAudio()

        if text.count(WAKE) > 0:
            speak("How can I help you?")
            start()

def start():
    text = get_audio()
    PROGRAMMING = ["i want to program", "i wanna program"]
    for phrase in PROGRAMMING:
        if phrase in text:
            speak("In which language would you like to program with?")
            response_text = get_audio()
            programIDE(response_text)
            initialize()

    GOOGLING = ["check this for me", "look up this for me", "search on web", "google this"]
    for phrase in GOOGLING:
        if phrase in text:
            speak("What am I supposed to look up?")
            response_text = get_audio()
            googling(response_text)
            initialize()

    CALENDAR_STRS = ["what do i have", "do i have", "am i busy", "do i have plans"]
    for phrase in CALENDAR_STRS:
        if phrase in text:
            date = get_date(text)
            if date:
                get_events(get_date(text), SERVICE)
                initialize()

    NOTE_STRS = ["make a note", "write it down", "remember this"]
    for phrase in NOTE_STRS:
        if phrase in text:
            speak("What would you like me to write down?")
            note_text = get_audio()
            note(note_text)
            speak("I've made a note of that.")
            initialize()

    NAME = ["what is your name", "what's your name", "how should I call you", "who are you"]
    for phrase in NAME:
        if phrase in text:
            speak("I am, Just A Rather Very Intelligent System, but you can simply call me Jarvis") 
            initialize()            

    GETHOURS = ["what time is it"]
    for phrase in GETHOURS:
        if phrase in text:
            time = datetime.datetime.now()
            speak("It's " + time.strftime("%I:%M %p"))
            initialize()

    GETTODAY = ["what day is it"]
    for phrase in GETTODAY:
        if phrase in text:
            time = datetime.datetime.now()
            speak(time.strftime("%a, %b %d, %Y"))
            initialize()
            
    SING = ["sing something"]
    for phrase in SING:
        if phrase in text:
            speak("We all live in a yellow submarine Yellow submarine, yellow submarine We all live in a yellow submarine Yellow submarine, yellow submarine. Wait, I'm not good at all!")
            initialize()

    EVENT = ["add this to my calendar", "could you remind me this", "add this to the agenda", "write this on my calendar", "create an event"]
    for phrase in EVENT:
        if phrase in text:
            speak("Ok. Tell me the date of your appointment")
            date_text = get_audio()
            get_date(date_text)
            speak("When is the event start?")
            start_Time = get_audio()
            speak("When is the event end?")
            end_Time = get_audio()
            speak("What is the title for this?")
            title = get_audio()
            add_event(date_text, start_Time, end_Time, title)
            speak("Got it. Your event is in your Google Calendar")
            initialize()

    WEATHERNOW = ["how is the weather now", "how is the weather"]
    for phrase in WEATHERNOW:
        if phrase in text:
            getCurrentWeather()
            initialize()

    COVID = ["coronavirus", "tell me something about coronavirus", "information about coronavirus"]
    for phrase in COVID:
        if phrase in text:                                            
            covid()
            initialize()

    WIKI = ["wikipedia"]
    for phrase in WIKI:
        if phrase in text:
            speak("Tell me just the name of what you are looking for")
            text = get_audio()                                            
            wiki(text)
            initialize()

    NEWS = ["give me the latest news", "latest news", "any news", "some news"]            
    for phrase in NEWS:
        if phrase in text:
            news()
            speak("Ok, here we have the first five results by Google")
            initialize()

    GOODBYE = ["bye", "shut up", "close yourself", "turn off", "bye bye"]
    for phrase in GOODBYE:
        if phrase in text:
            exit()

    GREETING = ["hi", "hello", "what's up", "what's good"]
    for phrase in GREETING:
        if phrase in text:
            speak("Nice to see you again")
            initialize()

    LOCATE = ["locate", "start locating function"]
    for phrase in LOCATE:
        if phrase in text:
            speak("Please for a better accuracy digit the coordinates using latitude and longitude as decimal numbers")
            lat = input("Latitude:")
            lon = input("Longitude:")
            locate(lat, lon)
            initialize()

    initialize()

#main
SERVICE = authenticate_google()
loading()