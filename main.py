'''import win32com.client
speaker=win32com.client.Dispatch('SAPI.SpVoice')
while 1:
    print('Enter the word you want to speak it out by the computer')
    s=input()
    speaker.Speak(s)'''
import speech_recognition as sr
import pyttsx3
import webbrowser
import openai
import os 
import datetime
import screen_brightness_control as sbc

def increase_brightness(step=10):
    try:
        current = sbc.get_brightness()[0]
        new_brightness = min(current + step, 100)
        sbc.set_brightness(new_brightness)
        say(f"Brightness increased to {new_brightness} percent")
    except Exception as e:
        print("Error:", e)
        say("I could not change the brightness.")


def decrease_brightness(step=10):
    try:
        current = sbc.get_brightness()[0]
        new_brightness = max(current - step, 0)
        sbc.set_brightness(new_brightness)
        say(f"Brightness reduced to {new_brightness} percent")
    except Exception as e:
        print("Error:", e)
        say("I could not change the brightness.")
def set_brightness_level(level):
    try:
        level = int(level)
        if 0 <= level <= 100:
            sbc.set_brightness(level)
            say(f"Brightness set to {level} percent")
        else:
            say("Please say a value between 0 and 100.")
    except:
        say("Invalid brightness value.")

def say(text):
    engine = pyttsx3.init()
    engine.say(text)
    engine.runAndWait()

def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.adjust_for_ambient_noise(source)
        audio = r.listen(source, phrase_time_limit=5)
    try:
        query = r.recognize_google(audio, language='en-IN')
        query = query.strip()
        query = ''.join(c for c in query if c.isprintable())
        print(f"User Said: {query}")
        return query
    except sr.UnknownValueError:
        return ''
    except Exception as e:
        print(f"Error: {e}")
        return ''

if __name__ == '__main__':
    print('Starting...')
    say("Greetings, I am Jarvis AI.")
    while True:
        query = takeCommand()
        sites=[['youtube','https://youtube.com'],['instagram', 'https://instagram.com'],['wikipedia','https://www.wikipedia.org/'],['google','https://www.google.com'],['gmail','https://mail.google.com']]
        for site in sites:
            if f'open {site[0]}'.lower() in query.lower():
                webbrowser.open(site[1])
                say(f'Opening {site[0]} ....')
        if "open ai chat gpt" in query.lower() or "open chat g p t" in query.lower():
            webbrowser.open("https://chat.openai.com/")
            say("Opening Chat G P T ....")
        # if "Whatsapp".lower() in query.lower():
        #     codePath: str = "C:\Program Files\WindowsApps\5319275A.WhatsAppDesktop_2.2587.9.0_x64__cv1g1gvanyjgm"
        #     os.startfile(codePath)
        #     say("Opening Whatsapp....")
        
        if "play music" in query.lower():
            musicPath: str = "E:\ENGG\My Projects\Jarvis\music.mp3"
            os.startfile(musicPath)
            say("Playing Music ....")
        if "exit" in query.lower() or "stop" in query.lower() or "bye" in query.lower():
            say("Goodbye , have a nice day.")
            break
        if "time" in query.lower():
            from datetime import datetime
            now = datetime.now()
            current_time = now.strftime("%I:%M:%S %p")
            print("Current Time =", current_time)
            say(f"The current time is {current_time}")
        if "date" in query.lower():
            from datetime import date
            today = date.today()
            print("Today's date:", today)
            say(f"Today's date is {today}")
        if "search for" in query.lower():
            search_query = query.lower().replace("search", "").strip()
            if search_query:
                url = f"https://www.google.com/search?q={search_query}"
                webbrowser.open(url)
                say(f"Searching for {search_query} on Google ....")
        if "open" in query.lower():
            app_name = query.lower().replace("open", "").strip()
            if app_name:
                try:
                    os.startfile(app_name)
                    say(f"Opening {app_name} ....")
                except Exception as e:
                    print(f"Could not open {app_name}: {e}")
                    say(f"Sorry, I couldn't open {app_name}.")
        if "increase brightness" in query.lower():
            increase_brightness()
        if "decrease brightness" in query.lower():
            decrease_brightness()
        if "set brightness to" in query.lower():
            level = query.lower().split("set brightness to")[-1].strip()
            set_brightness_level(level)
       
      
       
    





