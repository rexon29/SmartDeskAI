import os
import pyttsx3
import requests
import datetime
import webbrowser
import psutil
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from comtypes import CLSCTX_ALL
import screen_brightness_control as sbc

# Initialize Text-to-Speech Engine with a Female Voice
engine = pyttsx3.init("sapi5")
voices = engine.getProperty("voices")
engine.setProperty("voice", voices[1].id)  # 1 is typically female
engine.setProperty("rate", 170)  # Adjust speaking rate

def speak(text):
    """Speak the given text and display it on the console."""
    print(f"Assistant: {text}")
    engine.say(text)
    engine.runAndWait()

def get_weather(city):
    """Fetch the weather for a given city."""
    API_KEY = "a4637789c6e7ab4caef6af41565b7428"
    url = f"http://api.openweathermap.org/data/2.5/weather?q={city}&appid={API_KEY}&units=metric"
    response = requests.get(url)
    if response.status_code == 200:
        data = response.json()
        temp = data["main"]["temp"]
        desc = data["weather"][0]["description"]
        return f"The temperature in {city} is {temp}°C with {desc}."
    else:
        return "I couldn't fetch the weather information. Please try again."

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
    speak("Restarting your computer.")

def shutdown_pc():
    """Shut down the PC."""
    os.system("shutdown /s /t 1")
    speak("Shutting down your computer.")

def log_off_pc():
    """Log off the user."""
    os.system("shutdown /l")
    speak("Logging off your computer.")

def set_volume(level):
    """Set system volume."""
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = interface.QueryInterface(IAudioEndpointVolume)
    volume.SetMasterVolumeLevelScalar(level / 100, None)
    speak(f"Volume set to {level}%.")

def set_brightness(level):
    """Set screen brightness."""
    sbc.set_brightness(level)
    speak(f"Brightness set to {level}%.")

def toggle_wifi(turn_on):
    """Toggle Wi-Fi on or off."""
    if turn_on:
        os.system("netsh interface set interface Wi-Fi enabled")
        speak("Wi-Fi is now enabled.")
    else:
        os.system("netsh interface set interface Wi-Fi disabled")
        speak("Wi-Fi is now disabled.")

def handle_query(query):
    """Process user queries and provide responses."""
    if "weather" in query:
        speak("Which city would you like the weather for?")
        city = input("User (City): ")
        weather = get_weather(city)
        speak(weather)

    elif "time" in query:
        now = datetime.datetime.now()
        speak(f"The current time is {now.strftime('%I:%M %p')}.")

    elif "open" in query:
        speak("Which app should I open?")
        app_name = input("User (App Name): ")
        open_app(app_name)

    elif "restart" in query:
        restart_pc()

    elif "shutdown" in query:
        shutdown_pc()

    elif "log off" in query:
        log_off_pc()

    elif "volume" in query:
        speak("Set the volume to what percentage?")
        try:
            level = int(input("User (Volume %): "))
            set_volume(level)
        except ValueError:
            speak("Please provide a valid percentage.")

    elif "brightness" in query:
        speak("Set the brightness to what percentage?")
        try:
            level = int(input("User (Brightness %): "))
            set_brightness(level)
        except ValueError:
            speak("Please provide a valid percentage.")

    elif "wifi" in query:
        speak("Should I turn Wi-Fi on or off?")
        wifi_status = input("User (On/Off): ").lower()
        if "on" in wifi_status:
            toggle_wifi(True)
        elif "off" in wifi_status:
            toggle_wifi(False)
        else:
            speak("Invalid choice. Please say on or off.")

    elif "google" in query:
        speak("What should I search on Google?")
        search_query = input("User (Search Query): ")
        webbrowser.open(f"https://www.google.com/search?q={search_query}")
        speak(f"Here are the results for {search_query}.")

    elif "youtube" in query:
        speak("What should I play on YouTube?")
        video = input("User (YouTube Query): ")
        webbrowser.open(f"https://www.youtube.com/results?search_query={video}")
        speak(f"Playing {video} on YouTube.")

    elif "goodbye" in query or "exit" in query:
        speak("Goodbye! Have a great day!")
        exit()
    else:
        speak("I'm sorry, I didn't understand that. Can you try again?")

# Main Loop
speak("Hello, I am your assistant. How can I help you today?")
while True:
    user_query = input("User: ").lower()
    handle_query(user_query)
