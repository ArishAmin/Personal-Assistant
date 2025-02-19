import speech_recognition as sr 
import os
import win32com.client as wincl
import webbrowser
import datetime
from config import apikey
import openai

speaker= wincl.Dispatch("SAPI.SpVoice")

def say(text):
    speaker.Speak(text)

def takecommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold= 1
        audio = r.listen(source, timeout=5)
        try:
            print("Recognizing...")
            query= r.recognize_google(audio, language='en-in')
            print(f"user said: {query}")
            return query
        except Exception as e:
            return "Some error occured. Sorry from Jarvis"

def ai(prompt):
    openai.api_key= apikey

    response= openai.Completion.create(
        prompt= prompt,
        model="gpt-3.5-turbo",
        max_tokens=150,
        temperature=0.5,
    )
    
    text += response["choices"][0]["text"]
    if not os.path.exists("GPT"):
        os.makedirs("GPT")

    with open(f"GPT/{' '.join(prompt.split('GPT')[1:]) }.txt", "w") as f:
        f.write(text)

def chat(query):
    openai.api_key= apikey
    chatstr= f"User: {query}\n Jarvis:"
    response= openai.Completion.create(
        engine="dgpt-3.5-turbo",
        prompt= chatstr,
        max_tokens=150,
        temperature=0.5,
    )
    chatstr += f"{response['choices'][0]['text']}\n"
    say(response['choices'][0]['text'])

    with open(f"GPT/{' '.join(prompt.split('GPT')[1:]) }.txt", "w") as f:
        f.write(text)   

if __name__ == "__main__":
   print('Python')
   say("Hello I am Jarvis AI")
   while True:
        print("listening...")
        query= takecommand()
        sites=[
            ["youtube", "www.youtube.com"], 
            ["google", "www.google.com"], 
            ["facebook", "www.facebook.com"], 
            ["instagram", "www.instagram.com"], 
            ["twitter", "www.twitter.com"], 
            ["linkedin", "www.linkedin.com"]
            ]
        for site in sites:
            if f"Open {site[0]}".lower() in query.lower():
                say(f"Opening {site[0]} bro...")
                webbrowser.open(site[1])

        Apps=[
            ["chrome", "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"],
            ['notepad', 'C:\\Windows\\System32\\notepad.exe'],
            ['word', 'C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE'],
            ['excel', 'C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE'],
            ['powerpoint', 'C:\\Program Files\\Microsoft Office\\root\\Office16\\POWERPNT.EXE'],
            ['outlook', 'C:\\Program Files\\Microsoft Office\\root\\Office16\\OUTLOOK.EXE'],
            ['vs code', 'C:\\Users\\User\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe'],
            ['whatsapp', 'C:\\Users\\User\\AppData\\Local\\WhatsApp\\WhatsApp.exe'],
            ['file explorer', 'C:\\Windows\\explorer.exe'],
            ['anaconda navigator', 'C:\\Users\\User\\anaconda3\\pythonw.exe'],
            ['recycle bin', 'C:\\Windows\\explorer.exe'],
        ]
        for App in Apps:
            if f"Open {App[0]}".lower() in query.lower():
                say(f"Opening {App[0]} bro...")
                webbrowser.open(App[1])

        if " the time" in query:
            strTime= datetime.datetime.now().strftime("%H:%M")
            say(f"Bro, the time is {strTime}")

        elif "search" in query:
            say("What do you want to search bro?")
            search= takecommand()
            webbrowser.open(f"https://www.google.com/search?q={search}")

        elif "play" in query:
            say("What do you want to play bro?")
            song= takecommand()
            webbrowser.open(f"https://www.youtube.com/results?search_query={song}")

        elif "Using GPT".lower() in query.lower():
            ai(prompt=query)

        elif "Jarvis Quit".lower() in query.lower():
            say("Goodbye Bro...")
            exit()

        elif "reset chat".lower() in query.lower():
            say("Resetting Chat...")
            chatstr=''

        else:
            print("Chatting...")
            chat(query)