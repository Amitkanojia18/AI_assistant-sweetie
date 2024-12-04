import openai
import os
import win32com.client
import speech_recognition as sr
import cv2
import pyttsx3
from datetime import datetime
import requests
import webbrowser
# speaker = win32com.client.Dispatch("SAPI.spVoice")

engine = pyttsx3.init()

# Open AI key setting

openai.api_key = "my secret key"

# Implemented another Voice

zira_voice_id = "HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Speech_OneCore\\Voices\\Tokens\\TTS_MS_EN-US_ZIRA_11.0"
engine.setProperty('voice', zira_voice_id)

def speak(text):
    engine.say(text)
    engine.runAndWait()

#speak with greeting

speak("Welcome! Your sweetie is here.")
speak("How can I help you.")

# Initialize the speech-to-text engine
def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening....")
        r.pause_threshold = 1
        audio = r.listen(source)
        print("captured")

    try:
        text = r.recognize_google(audio)
        print(f"You said: {text}")
        return text
    except sr.UnknownValueError:
        print("Sorry, I could not understand the audio.")
        speak("Sorry, I could not understand the audio.")
        return None
    except sr.RequestError:
        print("Sorry, My speech service is down.")
        speak("Sorry, My speech service is down.")
        return None


# website command

def open_website(site_name):
    url = f"https://www.{site_name}.com"
    webbrowser.open(url)
    speak(f"Opening {site_name}")

def getDate():
    now = datetime.now()
    date = now.strftime("%B %d, %Y")
    print(date)
    speak(f"Today's date is {date}")

def getTime():
    now = datetime.now()
    time = now.strftime("%I:%M %p")
    print(time)
    speak(f"The current time is {time}")

def fetch_duckduckgo_answer(query):
    url = f"https://api.duckduckgo.com/?q={query}&format=json"
    response = requests.get(url).json()

    if 'AbstractText' in response and response['AbstractText']:
        return response['AbstractText']
    else:
        return "I couldn't find an answer to that question."    

# def generate_response(prompt):
#     response = openai.ChatCompletion.create(
#         model="gpt-3.5-turbo",
#         messages=[
#             {"role": "system", "content": "You are a helpful assistant."},
#             {"role": "user", "content": prompt}
#         ]
#     )
#     return response['choices'][0]['message']['content']

def opencamera():
    speak("Opening Camera")
    cap = cap = cv2.VideoCapture(0, cv2.CAP_DSHOW)

    if not cap.isOpened():
        speak("Sorry, I can't access the camera right now!")
        return

    cv2.namedWindow("Camera", cv2.WND_PROP_FULLSCREEN)
    cv2.setWindowProperty("Camera", cv2.WND_PROP_FULLSCREEN, cv2.WINDOW_FULLSCREEN)

    while True:
        ret, frame = cap.read()
        if not ret:
            speak("Failed to grab frame.")
            break

        cv2.imshow("Camera", frame)

        if cv2.waitKey(1) & 0xFF == ord('q'):
            break

    cap.release()
    cv2.destroyAllWindows()


def process_command(command):
    if 'open' in command:
        if 'camera' in command:
            opencamera()
        else:
            site_name = command.replace('open ', '').strip()
            open_website(site_name)

    elif 'date' in command or 'time' in command:
        getDate()
        speak("and")
        getTime()

    else:
        # Treat everything else as a query and fetch an answer using DuckDuckGo
        answer = fetch_duckduckgo_answer(command)
        print(f"{answer}")
        speak(answer)


while True:
    command = takeCommand()
    
    if command:
        process_command(command)




# Continuous loop to listen and speak
'''
while True:
    command = takeCommand()
    if command:
        if 'open' in command:
            if 'camera' in command:
                opencamera()

            else:
                site_name = command.replace('open ', '').strip()
                open_website(site_name)


        elif "ask" in command.lower():
            prompt = command.replace('ask', '').strip()
            answer = generate_response(prompt)
            speak(answer)

        elif 'date' and 'time' in command:
            getDate()
            speak("and")
            getTime()

        elif 'date' in command:
            getDate()

        elif 'time' in command:
            getTime()

        else:
            speak(command)
'''

