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
    say("Namaskar, I am Jarvis AI.")
    while True:
        query = takeCommand()
        sites=[['youtube','https://youtube.com'],['instagram', 'https://instagram.com'],['wikipedia','https://www.wikipedia.org/'],['google','https://www.google.com'],['gmail','https://mail.google.com']]
        for site in sites:
            if f'open {site[0]}'.lower() in query.lower():
                webbrowser.open(site[1])
                say(f'Opening {site[0]} sir....')
        if "open ai chat gpt" in query.lower():
            webbrowser.open("https://chat.openai.com/")
            say("Opening Chat G P T sir....")
        if "exit" in query.lower() or "stop" in query.lower() or "bye" in query.lower():
            say("Goodbye sir, have a nice day.")
            break
        if "play music" in query.lower():
            musicPath: str = "E:\ENGG\My Projects\Jarvis\music.mp3"
            os.startfile(musicPath)
            say("Playing Music Sir....")





