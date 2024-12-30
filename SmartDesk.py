import pyttsx3
import datetime
import speech_recognition as sr
import wikipedia
import webbrowser as wb
import os
import random
import pyautogui
import requests
import ctypes
import subprocess

# Initialize the Text-to-Speech Engine
engine = pyttsx3.init()
voices = engine.getProperty("voices")
engine.setProperty("voice", voices[0].id)  # Set male voice
engine.setProperty("rate", 180)  # Adjust speaking rate


def speak(audio):
    print(f"Assistant: {audio}")
    engine.say(audio)
    engine.runAndWait()
def time_():
    current_time = datetime.datetime.now().strftime("%I:%M %p")
    speak(f"The current time is {current_time}.")
    print(f"The current time is {current_time}.")


def date_():
    today = datetime.datetime.now()
    speak(f"Today's date is {today.strftime('%A, %d %B %Y')}.")
    print(f"Today's date is {today.strftime('%A, %d %B %Y')}.")


def greeting():
    hour = datetime.datetime.now().hour
    if 5 <= hour < 12:
        return "Good Morning!"
    elif 12 <= hour < 17:
        return "Good Afternoon!"
    else:
        return "Good Evening!"


def intro():
    greet = greeting()
    speak(f"{greet} I am Jarvis, your personal assistant. I am here to make your day easier. "
          "You can ask me to perform tasks, fetch information, or just chat. How can I assist you today?")
    print(f"{greet} I am Jarvis, your personal assistant. I am here to make your day easier. "
          "You can ask me to perform tasks, fetch information, or just chat. How can I assist you today?")


def take_command():
    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        recognizer.pause_threshold = 1
        audio = recognizer.listen(source)

    try:
        print("Recognizing...")
        query = recognizer.recognize_google(audio, language="en-in")
        print(f"User: {query}")
        return query.lower()
    except sr.UnknownValueError:
        speak("Sorry, I didn't catch that. Please repeat.")
        return "None"
    except sr.RequestError:
        speak("Network error. Please check your connection.")
        return "None"


def weather(city):
    api_key = "a4637789c6e7ab4caef6af41565b7428"
    url = f"http://api.openweathermap.org/data/2.5/weather?q={city}&appid={api_key}&units=metric"
    response = requests.get(url)
    if response.status_code == 200:
        data = response.json()
        temperature = data['main']['temp']
        description = data['weather'][0]['description']
        speak(f"The weather in {city} is {description} with a temperature of {temperature} degrees Celsius.")
    else:
        speak("I couldn't retrieve the weather details. Please try again.")


def random_news():
    news = [
        "India wins the Cricket World Cup!",
        "Scientists discover water on Mars.",
        "New advancements in AI technology unveiled.",
        "Electric cars to dominate the market by 2030.",
        "Global leaders commit to combating climate change."
    ]
    speak("Here's a random headline: " + random.choice(news))


def system_command(action):
    if action == "lock":
        speak("Locking the computer.")
        ctypes.windll.user32.LockWorkStation()
    elif action == "shutdown":
        speak("Shutting down the system.")
        os.system("shutdown /s /t 5")
    elif action == "restart":
        speak("Restarting the system.")
        os.system("shutdown /r /t 5")


def open_website(site_name, search_query=None):
    websites = {
        "youtube": "https://www.youtube.com",
        "google": "https://www.google.com",
        "spotify": "https://www.spotify.com",
        "instagram": "https://www.instagram.com",
        "chatgpt": "https://chat.openai.com",
        "netflix": "https://www.netflix.com",
        "prime video": "https://www.primevideo.com",
        "disney hotstar": "https://www.hotstar.com"
    }
    if site_name in websites:
        if search_query:
            if site_name == "youtube":
                url = f"{websites['youtube']}/results?search_query={search_query}"
            elif site_name == "google":
                url = f"{websites['google']}/search?q={search_query}"
            else:
                speak(f"Search functionality is not available for {site_name}. Opening the main site.")
                url = websites[site_name]
            speak(f"Searching for {search_query} on {site_name}.")
        else:
            url = websites[site_name]
            speak(f"Opening {site_name}.")
        wb.open(url)
    else:
        speak("I couldn't find the website you mentioned.")


def open_app(app_name):
    app_paths = {
        "notepad": "notepad",
        "calculator": "calc",
        "paint": "mspaint",
        "word": r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE",
        "excel": r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE",
    }
    if app_name in app_paths:
        speak(f"Opening {app_name}.")
        subprocess.run(app_paths[app_name], shell=True)
    else:
        speak("I couldn't find the application you mentioned. Please check the name.")


def adjust_system_settings(action, value=None):
    if action == "brightness":
        try:
            sbc = __import__('screen-brightness-control')
            sbc.set_brightness(value)
            speak(f"Brightness set to {value} percent.")
        except Exception as e:
            speak("Unable to adjust brightness. Please check your system settings.")
    elif action == "volume up":
        pyautogui.press("volumeup", presses=5)
        speak("Volume increased.")
    elif action == "volume down":
        pyautogui.press("volumedown", presses=5)
        speak("Volume decreased.")
    elif action == "mute":
        pyautogui.press("volumemute")
        speak("Volume muted.")


if __name__ == "__main__":
    intro()

    while True:
        query = take_command()

        if "time" in query:
            time_()
        elif "date" in query:
            date_()
        elif "wikipedia" in query:
            speak("What should I search for on Wikipedia?")
            topic = take_command()  # Wait for the user's input
            if topic != "None" and topic.strip():
                try:
                    speak(f"Searching for {topic} on Wikipedia...")
                    results = wikipedia.summary(topic, sentences=2)
                    speak(f"According to Wikipedia, {results}")
                    print(f"According to Wikipedia, {results}")
                    wb.open(f"https://en.wikipedia.org/wiki/{topic.replace(' ', '_')}")
                except Exception as e:
                    speak("An error occurred while fetching data from Wikipedia.")
                    print(f"Error: {e}")
        elif "weather" in query:
            speak("Which city's weather would you like to know?")
            city = take_command()
            weather(city)
        elif "news" in query:
            random_news()
        elif "lock my computer" in query:
            system_command("lock")
        elif "shutdown my computer" in query:
            system_command("shutdown")
        elif "restart my computer" in query:
            system_command("restart")
        elif "open" in query:
            if "search on" in query:
                words = query.split()
                site_name = words[words.index("on") + 1]
                search_query = " ".join(words[words.index("on") + 2:])
                open_website(site_name, search_query)
            else:
                site_name = query.replace("open", "").strip()
                open_website(site_name)
        elif "launch" in query:
            app_name = query.replace("launch", "").strip()
            open_app(app_name)
        elif "set brightness" in query:
            speak("What level should I set the brightness to?")
            try:
                level = int(take_command())
                adjust_system_settings("brightness", value=level)
            except ValueError:
                speak("Please provide a valid number.")
        elif "volume up" in query:
            adjust_system_settings("volume up")
        elif "volume down" in query:
            adjust_system_settings("volume down")
        elif "mute" in query:
            adjust_system_settings("mute")
        elif "exit" in query or "bye" in query:
            speak("Goodbye! Have a wonderful day!")
            break
        else:
            speak("I didn't understand that. Could you please repeat?")
