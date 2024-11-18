import speech_recognition as sr
import win32com.client
import os
import webbrowser
import time
import google.generativeai as genai
from dotenv import load_dotenv

load_dotenv()  # Load environment variables from .env file
Api_key = os.getenv('API_KEY')  # Access the API key

speaker = win32com.client.Dispatch("SAPI.SpVoice")
print("Initializing...")

def chatbot(query):
    api_key = Api_key
    genai.configure(api_key=api_key)

    if not api_key:
        raise ValueError("API_KEY not found. Please set it in the .env file.")

    model = genai.GenerativeModel('gemini-pro')
    chat = model.start_chat(history=[])
    while True:
        question = query
        if question.strip() == '':
            break
        response = chat.send_message(question)
        cleaned_text = response.text.replace('*', '')
        print(cleaned_text)
        speaker.Speak(cleaned_text)
        return

def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.pause_threshold = 1
        audio = r.listen(source)
        try:
            query = r.recognize_google(audio, language="en-in")
            print(f"User  said: {query}")
            return query.lower()
        except sr.UnknownValueError:
            print("Sorry, your request cannot be fulfilled!")
            return "Sorry, your request cannot be fulfilled!"
        except sr.RequestError:
            print("Sorry, the service is down!")
            return "Sorry, the service is down!"




if __name__ == '__main__':
    speaker.Speak("Hello, I am Jarvis.")
    while True:
        print("Listening...")
        query = takeCommand()
        
        

        speaker.Speak(query)

        if "stop" in query:
            break

        elif "open discord" in query.lower():
            speaker.Speak("Opening Discord")
            os.startfile("C:/Users/depan/Desktop/Discord.lnk")

        if "the time" in query:
            current_time = time.strftime("%H:%M:%S")
            print(current_time)
            speaker.Speak(current_time)

        elif "the date" in query:
            current_date = time.strftime("%d/%m/%Y")
            print(current_date)
            speaker.Speak(current_date)

        elif "the day" in query:
            current_day = time.strftime("%A")
            print(current_day)
            speaker.Speak(current_day)

        elif "open" in query:
            sites = [["Youtube", "https://www.youtube.com/"],
                     ["Google", "https://www.google.com/"],
                     ["Gmail", "https://mail.google.com/"],
                     ["Wikipedia", "https://www.wikipedia.org/"]]

            for site in sites:
                if f"open {site[0]}".lower() in query:
                    speaker.Speak(f"Opening {site[0]}")
                    webbrowser.open(site[1])
                    break
        else:
            chatbot(query)