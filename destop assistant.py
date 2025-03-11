import numpy as np
import random
import os
import pickle
import speech_recognition as sr
import pyttsx3
import webbrowser
import requests
from sklearn.feature_extraction.text import CountVectorizer
from sklearn.model_selection import StratifiedKFold
from sklearn.ensemble import RandomForestClassifier
from sklearn.metrics import accuracy_score, classification_report
from collections import Counter
from faker import Faker
import nltk
import time
import threading
from plyer import notification
from datetime import datetime
import wikipediaapi
from newspaper import Article
import openai
import pyautogui
import pywhatkit
from bs4 import BeautifulSoup
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
import screen_brightness_control as sbc
import psutil
import re
import pygame



# Sample dataset (commands and corresponding labels)
data = [
    ("open youtube", "open_website"),
    ("open gmail", "open_gmail"),
    ("increase volume", "adjust_volume"),
    ("turn up the volume", "adjust_volume"),
    ("make it louder", "adjust_volume"),
    ("decrease volume", "adjust_volume"),
    ("lower the volume", "adjust_volume"),
    ("mute sound", "adjust_volume"),
    ("unmute sound", "adjust_volume"),
    ("what time is it", "tell_time"),
    ("play music", "play_music"),
    ("stop music", "stop_music"),
    ("shutdown", "shut_down"),
    ("open notepad", "open_application"),
    ("open calculator", "open_application"),
    ("weather in", "get_weather"),
    ("open photos", "open_photos"),
    ("open powerpoint", "open_powerpoint"),
    ("open paint", "open_paint"),
    ("open pdf reader", "open_pdf_reader"),
    ("search for", "search_google"),
    ("introduce yourself", "introduce_yourself"),
    ("who created you", "who_created_you"),
    ("what is your name", "what_is_your_name"),
    ("set an alarm for", "set_alarm"),
    ("close notepad", "close_application"),
    ("close calculator", "close_application"),
    ("exit", "exit_program"),
]

commands, labels = zip(*data)

# Use CountVectorizer for feature representation
vectorizer = CountVectorizer()
X = vectorizer.fit_transform(commands)

# Stratified K-Fold Cross Validation for better model evaluation
kf = StratifiedKFold(n_splits=5, shuffle=True, random_state=42)
model = RandomForestClassifier(n_estimators=100, random_state=42)

accuracies = []
y = np.array(labels)
class_distribution = Counter(y)  # Check class distribution
print(f"Class distribution: {class_distribution}")

# Ensure there are enough samples for K-Fold
if len(class_distribution) < 2:
    print("Not enough classes for K-Fold. Please add more unique labels.")
else:
    for train_index, test_index in kf.split(X, y):
        X_train, X_test = X[train_index], X[test_index]
        y_train, y_test = y[train_index], y[test_index]
        model.fit(X_train, y_train)
        y_pred = model.predict(X_test)
        accuracy = accuracy_score(y_test, y_pred)
        accuracies.append(accuracy)
        print(f"Fold Accuracy: {accuracy:.4f}")

    print(f"Mean Accuracy: {np.mean(accuracies):.4f}")
    print(classification_report(y_test, y_pred, zero_division=0))

# Save model and vectorizer
with open('rf_model.pkl', 'wb') as model_file:
    pickle.dump(model, model_file)

with open('vectorizer.pkl', 'wb') as vec_file:
    pickle.dump(vectorizer, vec_file)

# Functions
def get_greeting():
    hour = datetime.now().hour
    if 0 <= hour < 12:
        return "Good Morning! I am your personal desktop assistant. How can I help you?"
    elif 12 <= hour < 18:
        return "Good Afternoon! I am your personal desktop assistant. How can I help you?"
    else:
        return "Good Evening! I am your personal desktop assistant. How can I help you?"


def say_greeting():
    engine = pyttsx3.init()
    greeting = get_greeting()
    engine.say(greeting)
    engine.runAndWait()
    print(greeting)


def open_gmail_compose():
    engine = pyttsx3.init()
    engine.say("Opening Gmail for composing a new email.")
    engine.runAndWait()
    compose_url = "https://mail.google.com/mail/u/0/?view=cm&fs=1&tf=1"
    webbrowser.open(compose_url)


def tell_time():
    current_time = datetime.now().strftime("%I:%M %p")
    engine = pyttsx3.init()
    engine.say(f"The time is {current_time}.")
    engine.runAndWait()
    print(f"The time is {current_time}.")


def open_website(website):
    engine = pyttsx3.init()
    engine.say(f"Opening {website}.")
    engine.runAndWait()
    web_urls = {
        "google": "https://www.google.com",
        "wikipedia": "https://www.wikipedia.org",
        "gmail": "https://mail.google.com",
        "youtube": "https://www.youtube.com",
        "chatgpt": "https://chat.openai.com",
        "google classroom": "https://classroom.google.com"
    }
    webbrowser.open(web_urls.get(website, "https://www.google.com"))


def shut_down():
    engine = pyttsx3.init()
    engine.say("Shutting down the system. Goodbye.Have a nice day")
    engine.runAndWait()
    os.system("shutdown /s /t 1")


current_song_process = None


def play_music():
    global current_song_process
    music_directory = "C:\\Users\\Shruti\\OneDrive\\Music\\Playlists\\songs"
    songs = [song for song in os.listdir(music_directory) if song.endswith(('.mp3', '.wav'))]
    if songs:
        song_to_play = random.choice(songs)
        current_song_process = os.startfile(os.path.join(music_directory, song_to_play))
        engine = pyttsx3.init()
        engine.say(f"Playing {song_to_play}.")
        engine.runAndWait()
    else:
        engine = pyttsx3.init()
        engine.say("No music files found in the specified directory.")
        engine.runAndWait()


def stop_music():
    global current_song_process
    if current_song_process is not None:
        os.system("taskkill /im wmplayer.exe /f")
        engine = pyttsx3.init()
        engine.say("Stopping the music.")
        engine.runAndWait()
        current_song_process = None
    else:
        engine = pyttsx3.init()
        engine.say("No music is currently playing.")
        engine.runAndWait()


def search_file(file_name):
    for root, dirs, files in os.walk("C:\\"):
        if file_name in files:
            engine = pyttsx3.init()
            engine.say(f"I found {file_name} in {root}.")
            engine.runAndWait()
            return
    engine = pyttsx3.init()
    engine.say(f"Sorry, I couldn't find {file_name}.")
    engine.runAndWait()


# Fun Facts API
FUN_FACTS_API = "https://uselessfacts.jsph.pl/random.json?language=en"

# Joke API
JOKE_API = "https://v2.jokeapi.dev/joke/Any"

def tell_joke():
    """Fetch and tell a random joke"""
    response = requests.get("https://v2.jokeapi.dev/joke/Any").json()

    if response.get("https://v2.jokeapi.dev/joke/Any") == "twopart":
        joke = f"{response['setup']} ... {response['delivery']}"
    else:
        joke = response.get("joke", "Sorry, I couldn't find a joke.")

    print(joke)
    engine.say(joke)
    engine.runAndWait()


def rock_paper_scissors():
    """Play Rock-Paper-Scissors with the user"""
    choices = ["rock", "paper", "scissors"]
    user_choice = input("Choose Rock, Paper, or Scissors: ").lower()
    bot_choice = random.choice(choices)

    if user_choice not in choices:
        print("Invalid choice! Please select Rock, Paper, or Scissors.")
        return

    print(f"Zynox chose {bot_choice}")
    engine.say(f"I chose {bot_choice}")

    if user_choice == bot_choice:
        result = "It's a tie!"
    elif (user_choice == "rock" and bot_choice == "scissors") or \
            (user_choice == "scissors" and bot_choice == "paper") or \
            (user_choice == "paper" and bot_choice == "rock"):
        result = "You win! 🎉"
    else:
        result = "I win! 😜"

    print(result)
    engine.say(result)
    engine.runAndWait()


def suggest_movie():
    """Recommend a random movie"""
    movies = ["Inception", "Interstellar", "The Matrix", "Avengers: Endgame", "The Dark Knight", "Titanic",
              "Harry Potter", "The Shawshank Redemption"]
    suggestion = random.choice(movies)

    print(f"I suggest you watch {suggestion} 🎬")
    engine.say(f"I suggest you watch {suggestion}")
    engine.runAndWait()





def fun_fact():
    """Fetch and share a random fun fact"""
    response = requests.get("https://uselessfacts.jsph.pl/random.json?language=en").json()
    fact = response.get("text", "Sorry, I couldn't find a fun fact.")

    print(f"Did you know? {fact}")
    engine.say(f"Did you know? {fact}")
    engine.runAndWait()

def get_weather(city):
    api_key = "3bcdc13bece71b3f5efbad2c4756bf3d"
    base_url = f"http://api.openweathermap.org/data/2.5/weather?q={city}&appid={api_key}&units=metric"
    response = requests.get(base_url)
    data = response.json()

    if data["cod"] != "404":
        main = data["main"]
        temperature = main["temp"]
        weather_description = data["weather"][0]["description"]
        engine = pyttsx3.init()
        engine.say(f"The temperature in {city} is {temperature} degrees Celsius with {weather_description}.")
        engine.runAndWait()
    else:
        engine = pyttsx3.init()
        engine.say(f"City {city} not found.")
        engine.runAndWait()


def introduce_yourself():
    engine = pyttsx3.init()
    introduction = (
        "Hello! I am your personal desktop assistant. I can assist you with tasks like opening applications, "
        "searching the web, and much more. "
        "How can I help you today?"
    )
    engine.say(introduction)
    engine.runAndWait()
    print(introduction)


def what_is_your_name():
    engine = pyttsx3.init()
    response = "My name is Zynox. You can call me Assistant or Zynox! How can I assist you today?"
    engine.say(response)
    engine.runAndWait()
    print(response)


def who_created_you():
    engine = pyttsx3.init()
    response = "I was created by my developer Shruti Dekate to assist you with tasks and make your life easier! How can I help you today?"
    engine.say(response)
    engine.runAndWait()
    print(response)


def set_alarm(alarm_time):
    """
    Function to set an alarm and run it in a separate thread.
    :param alarm_time: The time for the alarm in HH:MM format.
    """
    def alarm_thread():
        while True:
            current_time = datetime.now().strftime("%H:%M")
            if current_time == alarm_time:
                engine = pyttsx3.init()
                engine.say("Wake up! This is your alarm notification. Wake up!")
                engine.runAndWait()
                print("Alarm ringing!")
                break
            time.sleep(1)

    threading.Thread(target=alarm_thread, daemon=True).start()


def open_application(app_name):
    app_paths = {
        "notepad": "notepad.exe",
        "calculator": "calc.exe",
        "paint": "mspaint.exe",
        "powerpoint": r"C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs\\Microsoft Office\\PowerPoint",
    }

    if app_name in app_paths:
        app_path = app_paths[app_name]
        os.startfile(app_path)
        engine = pyttsx3.init()
        engine.say(f"Opening {app_name}.")
        engine.runAndWait()
        print(f"{app_name.capitalize()} opened.")
    else:
        engine = pyttsx3.init()
        engine.say(f"Sorry, I couldn't find the application {app_name}.")
        engine.runAndWait()
        print(f"Application {app_name} not found.")
def take_screenshot():
    folder_path = "C:\\Users\\Shruti\\OneDrive\\Pictures\\Screenshots"  # Change this path
    os.makedirs(folder_path, exist_ok=True)  # Create folder if it doesn't exist

    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    filename = f"{folder_path}/screenshot_{timestamp}.png"

    screenshot = pyautogui.screenshot()
    screenshot.save(filename)

    speak(f"Screenshot saved successfully in Screenshots folder .Please check it.")


def close_application(app_name):
    app_processes = {
        "notepad": "notepad.exe",
        "calculator": "calc.exe",
        "paint": "mspaint.exe",
        "powerpoint":r"C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs\\Microsoft Office\\PowerPoint",
    }

    if app_name in app_processes:
        os.system(f"taskkill /f /im {app_processes[app_name]}")
        engine = pyttsx3.init()
        engine.say(f"Closing {app_name}.")
        engine.runAndWait()
        print(f"{app_name.capitalize()} closed.")
    else:
        engine = pyttsx3.init()
        engine.say(f"Sorry, I couldn't find any running process for {app_name}.")
        engine.runAndWait()
        print(f"No running process found for {app_name}.")


# Function to fetch a short summary from Wikipedia
def search_wikipedia(query):
    wiki_wiki = wikipediaapi.Wikipedia('en')
    page = wiki_wiki.page(query)
    if page.exists():
        summary = page.summary[:500]  # Limit summary to 500 characters
        print(f"Wikipedia Summary for {query}:\n{summary}")
        speak(summary)
    else:
        print("No Wikipedia page found.")
        speak("No Wikipedia page found.")


# Function to fetch latest news headlines
def get_news():
    API_KEY = "6ea9321cef9544dc8fe0dd93e258c72d"  # Replace with your NewsAPI key
    url = f"https://newsapi.org/v2/top-headlines?country=us&apiKey={API_KEY}"
    response = requests.get(url).json()
    articles = response.get("articles", [])[:5]  # Get top 5 headlines

    if articles:
        print("Latest News Headlines:")
        for idx, article in enumerate(articles, start=1):
            print(f"{idx}. {article['title']}")
            speak(article['title'])
    else:
        print("No news articles found.")
        speak("No news articles found.")


# Function to fetch and read aloud an article
def read_article(url):
    article = Article(url)
    article.download()
    article.parse()
    article.nlp()
    print(f"Reading article: {article.title}\n")
    speak(article.text[:500])  # Read first 500 characters


# Function to perform Google Search
def search_google(query):
    search_url = f"https://www.google.com/search?q={query}"
    print(f"Searching Google for: {query}")
    engine.say(f"Searching Google for {query}.")
    #speak(f"Searching Google for {query}")
    webbrowser.open(search_url)

def take_screenshot():
    folder_path = "C:\\Users\\Shruti\\OneDrive\\Pictures\\Screenshots"  # Change this path
    os.makedirs(folder_path, exist_ok=True)  # Create folder if it doesn't exist

    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    filename = f"{folder_path}/screenshot_{timestamp}.png"

    screenshot = pyautogui.screenshot()
    screenshot.save(filename)

    speak(f"Screenshot saved successfully in Screenshots folder Please check it.")

def play_youtube(video):
    try:
        print(f"Searching and playing '{video}' on YouTube...")
        pywhatkit.playonyt(video)
        print("YouTube should open now.")
    except Exception as e:
        print(f"Error: {e}")


def check_system_status():
    ram = psutil.virtual_memory().percent
    cpu = psutil.cpu_percent(interval=1)
    return f"RAM Usage: {ram}%\nCPU Usage: {cpu}%"

# Example Usage
print(check_system_status())


# Text-to-speech function
def speak(text):
    engine = pyttsx3.init()
    engine.say(text)
    engine.runAndWait()


# Main program loop
say_greeting()

while True:
    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        recognizer.adjust_for_ambient_noise(source, duration=1)
        print("Listening for commands...")
        audio = recognizer.listen(source)

    try:
        command = recognizer.recognize_google(audio).lower()
        print(f"You said: {command}")

        # Interpret command and perform action
        if "open" in command:
            if "gmail" in command:
                open_gmail_compose()
            elif "youtube" in command:
                open_website("youtube")
            elif "notepad" in command:
                open_application("notepad")
            elif "calculator" in command:
                open_application("calculator")
            elif "paint" in command:
                open_application("paint")
            else:
                print("Unknown open command.")
        elif "tell me the time" in command or "what time is it" in command:
            tell_time()
        elif "screenshot" in command:
            take_screenshot()
        elif "search for" in command:
            query = command.replace("search for", "").strip()
            webbrowser.open(f"https://www.google.com/search?q={query}")
        elif "introduce yourself" in command:
            introduce_yourself()
        elif "your name" in command:
            what_is_your_name()
        elif "who created you" in command:
            who_created_you()
        elif "weather in" in command:
            city = command.replace("weather in", "").strip()
            get_weather(city)
        elif "set alarm" in command:
            alarm_time = command.replace("set an alarm for", "").strip()
            set_alarm(alarm_time)
        elif "close" in command:
            if "notepad" in command:
                close_application("notepad")
            elif "calculator" in command:
                close_application("calculator")
        elif "play music" in command:
            play_music()
        elif "stop music" in command:
            stop_music()
        if "search wikipedia for" in command:
            query = command.replace("search wikipedia for", "").strip()
            search_wikipedia(query)
        elif "latest news" in command:
            get_news()
        elif "read article from" in command:
            url = command.replace("read article from", "").strip()
            read_article(url)
        elif "search google for" in command:
            query = command.replace("search google for", "").strip()
            search_google(query)
        elif "screenshot" in command:
            take_screenshot()
        elif "youtube" in command:
            play_youtube(command)
        elif "check system status" in command or "check ram" in command:
            print(check_system_status())
        elif "fun Fact" in command :
            fun_fact(command)
        elif "play game" in command:
            rock_paper_scissors()
        elif "joke" in command :
            tell_joke()
        elif "suggest movie" in command:
            suggest_movie()
        elif "exit" in command:
            print("Exiting...")
            engine = pyttsx3.init()
            engine.say("Goodbye! Have a great day!")
            engine.runAndWait()
            break
    except sr.UnknownValueError:
        speak("Sorry, I could not understand your request.")
    except Exception as e:
        print(f"An error occurred: {e}")
        speak("An error occurred while processing your request.")
    except sr.UnknownValueError:
        print("Sorry, I could not understand the audio.")
    except Exception as e:
        print(f"An error occurred: {e}")