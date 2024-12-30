import os
import pyttsx3
import speech_recognition as sr
import requests
import datetime
import webbrowser
import pyautogui
6import psutil
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from comtypes import CLSCTX_ALL
import screen_brightness_control as sbc

# Initialize Text-to-Speech Engine with a Female Voice
engine = pyttsx3.init("sapi5")
voices = engine.getProperty("voices")
engine.setProperty("voice", voices[1].id)  # 1 is typically female; change if required
engine.setProperty("rate", 170)  # Adjust speaking rate

def speak(text):
    """Speak the given text."""
    print(f"Assistant: {text}")
    engine.say(text)
    engine.runAndWait()

def take_command():
    """Capture audio input from the user and convert it to text."""
    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        recognizer.adjust_for_ambient_noise(source)
        audio = recognizer.listen(source)
    try:
        query = recognizer.recognize_google(audio, language="en-in")
        print(f"User: {query}")
    except sr.UnknownValueError:
        speak("Sorry, I didn't catch that. Could you please repeat?")
        return "None"
    return query.lower()

# Functionalities
def open_app(app_name):
    """Open specific apps."""
    app_paths = {
        "notepad": "notepad.exe",
        "chrome": "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",  # Adjust paths
        "spotify": "C:\\Users\\YourUserName\\AppData\\Roaming\\Spotify\\Spotify.exe",
        "file explorer": "explorer",
    }
    if app_name in app_paths:
        os.startfile(app_paths[app_name])
        speak(f"Opening {app_name}.")
    else:
        speak(f"I couldn't find {app_name} on your system.")

def restart_pc():
    """Restart the PC."""
    os.system("shutdown /r /t 1")

def shutdown_pc():
    """Shut down the PC."""
    os.system("shutdown /s /t 1")

def log_off_pc():
    """Log off the user."""
    os.system("shutdown /l")

def set_volume(level):
    """Set system volume."""
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = interface.QueryInterface(IAudioEndpointVolume)
    volume.SetMasterVolumeLevelScalar(level / 100, None)

def set_brightness(level):
    """Set screen brightness."""
    sbc.set_brightness(level)

def toggle_wifi(turn_on):
    """Toggle Wi-Fi on or off."""
    if turn_on:
        os.system("netsh interface set interface Wi-Fi enabled")
        speak("Wi-Fi is now enabled.")
    else:
        os.system("netsh interface set interface Wi-Fi disabled")
        speak("Wi-Fi is now disabled.")

def get_weather(city):
    """Fetch the weather for a given city."""
    API_KEY = "a4637789c6e7ab4caef6af41565b7428"  # Replace with your OpenWeatherMap API Key
    url = f"http://api.openweathermap.org/data/2.5/weather?q={city}&appid={API_KEY}&units=metric"
    response = requests.get(url)
    if response.status_code == 200:
        data = response.json()
        temp = data["main"]["temp"]
        desc = data["weather"][0]["description"]
        return f"The temperature in {city} is {temp}°C with {desc}."
    else:
        return "I couldn't fetch the weather information. Please try again."

def handle_query(query):
    """Process user queries and perform actions."""
    if "weather" in query:
        speak("Which city would you like the weather for?")
        city = take_command()
        if city != "None":
            weather = get_weather(city)
            speak(weather)

    elif "time" in query:
        now = datetime.datetime.now()
        speak(f"The current time is {now.strftime('%I:%M %p')}.")

    elif "open" in query:
        speak("Which app should I open?")
        app_name = take_command()
        if app_name != "None":
            open_app(app_name)

    elif "restart" in query:
        speak("Restarting the computer.")
        restart_pc()

    elif "shutdown" in query:
        speak("Shutting down the computer.")
        shutdown_pc()

    elif "log off" in query:
        speak("Logging off the computer.")
        log_off_pc()

    elif "volume" in query:
        speak("Set the volume to what percentage?")
        try:
            level = int(take_command())
            set_volume(level)
            speak(f"Volume set to {level}%.")
        except ValueError:
            speak("Please provide a valid percentage.")

    elif "brightness" in query:
        speak("Set the brightness to what percentage?")
        try:
            level = int(take_command())
            set_brightness(level)
            speak(f"Brightness set to {level}%.")
        except ValueError:
            speak("Please provide a valid percentage.")

    elif "wifi" in query:
        speak("Should I turn Wi-Fi on or off?")
        wifi_status = take_command()
        if "on" in wifi_status:
            toggle_wifi(True)
        elif "off" in wifi_status:
            toggle_wifi(False)

    elif "google" in query:
        speak("What should I search on Google?")
        search_query = take_command()
        if search_query != "None":
            webbrowser.open(f"https://www.google.com/search?q={search_query}")
            speak(f"Here are the results for {search_query}.")

    elif "youtube" in query:
        speak("What should I play on YouTube?")
        video = take_command()
        if video != "None":
            webbrowser.open(f"https://www.youtube.com/results?search_query={video}")
            speak(f"Playing {video} on YouTube.")

    elif "goodbye" in query or "exit" in query:
        speak("Goodbye! Have a great day!")
        exit()

    else:
        speak("I'm sorry, I didn't understand that. Can you try again?")

# Main Loop
speak("Hello, I am your assistant. Hey,what’s on your mind today? Anything I can help you with or chat about? That would be me! I’m pi, your personal ai. my name stands for “personal intelligence”, which is what i aim to provide for you. think of me as your friendly ai companion!?How’s your day going so far?")
while True:
    user_query = take_command()
    if user_query != "None":
        handle_query(user_query)
