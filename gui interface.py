import numpy as np
import random
import os
import datetime
import pickle
import speech_recognition as sr
import pyttsx3
import webbrowser
import requests
from sklearn.feature_extraction.text import CountVectorizer
from sklearn.model_selection import train_test_split, StratifiedKFold
from sklearn.ensemble import RandomForestClassifier
from sklearn.metrics import accuracy_score, classification_report
from collections import Counter
import tkinter as tk
from threading import Thread
from PIL import Image, ImageTk

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
    ("search for", "search_google")
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

# Function for assistant action
def get_greeting():
    hour = datetime.datetime.now().hour
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
    webbrowser.open(compose_url, new=2)
    print(f"opening gmail ")
    # **Updated to open in default browser**

def tell_time():
    current_time = datetime.datetime.now().strftime("%I:%M %p")
    engine = pyttsx3.init()
    engine.say(f"The time is {current_time}.")
    engine.runAndWait()
    print(f"The time is {current_time}.")

def open_website(website):
    engine = pyttsx3.init()
    engine.say(f"Opening {website}.")
    engine.runAndWait()
    print(f"opning {website}")
    web_urls = {
        "google": "https://www.google.com",
        "wikipedia": "https://www.wikipedia.org",
        "gmail": "https://mail.google.com",
        "youtube": "https://www.youtube.com",
        "chatgpt": "https://chat.openai.com",
        "google classroom": "https://classroom.google.com"
    }
    webbrowser.open(web_urls.get(website, "https://www.google.com"), new=2)  # **Updated to ensure default browser is used**

def shut_down():
    engine = pyttsx3.init()
    engine.say("Shutting down the system. Goodbye.")
    engine.runAndWait()
    os.system("shutdown /s /t 1")
    print(f"shutting down the system")

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
    print(f" playing {play_music()}")

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
    print({engine})

def get_weather(city):
    api_key = "3bcdc13bece71b3f5efbad2c4756bf3d"  # Replace with your OpenWeatherMap API key
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
    print({engine.say})

import subprocess

def open_application(app_name):
    app_paths = {
        "notepad": r"C:\\Users\\Shruti\\AppData\\Local\\Microsoft\\WindowsApps\\notepad.exe",
        "calculator": "calc.exe",
        "paint": "mspaint.exe",
        "powerpoint": r"C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\POWERPNT.EXE",
        "telegram desktop": r"C:\\Users\\Shruti\\AppData\\Roaming\\Telegram Desktop\\Telegram.exe",
        "photos": r"C:\\Users\\Shruti\\OneDrive\\Pictures\\Saved Pictures",
        "whatsapp": r"C:\\Program Files\\WindowsApps\\WhatsApp.exe",
    }

    if app_name in app_paths:
        app_path = app_paths[app_name]
        try:
            os.startfile(app_path)
            engine = pyttsx3.init()
            engine.say(f"Opening {app_name}.")
            print({engine.say})
        except Exception as e:
            engine.say(f"Failed to open {app_name}. Error: {str(e)}")
    else:
        engine = pyttsx3.init()
        engine.say(f"Sorry, I don't know how to open {app_name}.")
    engine.runAndWait()


def close_application(app_name):
    app_processes = {
        "notepad": "notepad.exe",
        "calculator": "calc.exe",
        "paint": "mspaint.exe",
        "powerpoint": "powerpnt.exe",
        "telegram": "Telegram.exe",
        "whatsapp": "WhatsApp.exe"
    }

    if app_name in app_processes:
        app_process = app_processes[app_name]
        os.system(f"taskkill /im {app_process} /f")
        engine = pyttsx3.init()
        engine.say(f"Closing {app_name}.")
        engine.runAndWait()
    else:
        engine = pyttsx3.init()
        engine.say(f"Sorry, I can't close {app_name}.")
        engine.runAndWait()


def process_command(command):
    command = command.lower()
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
        elif "photos" in command:
            open_application("photos")
        elif "powerpoint" in command:
            open_application("powerpoint")
        elif "whatsapp" in command:
            open_application("whatsapp")
        else:
            website = command.split("open ")[-1]
            open_website(website)
    elif "close" in command:
        app_name = command.split("close ")[-1]
        close_application(app_name)
    elif "time" in command:
        tell_time()
    elif "play music" in command:
        play_music()
    elif "stop music" in command:
        stop_music()
    elif "shutdown" in command:
        shut_down()
    elif "weather in" in command:
        city = command.split("in ")[-1]
        get_weather(city)
    elif "search for" in command:
        search_query = command.split("for ")[-1]
        webbrowser.open(f"https://www.google.com/search?q={search_query}")
    else:
        print("Command not recognized.")
# Function to start the voice recognition in a separate thread

def process_command(command):
    command = command.lower()
    if "open" in command:
        if "gmail" in command:
            open_gmail_compose()
        elif "youtube" in command:
            open_website("youtube")
        else:
            website = command.split("open ")[-1]
            open_website(website)
    elif "time" in command:
        tell_time()
    else:
        print("Command not recognized.")


# Function to start the voice recognition in a separate thread
def listen_command():
    with sr.Microphone() as source:
        print("Listening...")
        recognizer = sr.Recognizer()
        audio = recognizer.listen(source)

    try:
        command = recognizer.recognize_google(audio)
        print(f"You said: {command}")
        process_command(command)
        output_text.set(f"Assistant: {command}")

    except sr.UnknownValueError:
        output_text.set("Assistant: Sorry, I could not understand the audio.")
    except sr.RequestError:
        output_text.set("Assistant: Could not request results from Google Speech Recognition service.")


# Function to exit the application
def exit_app():
    engine = pyttsx3.init()
    engine.say("Goodbye! have a nice day .")
    engine.runAndWait()
    root.destroy()


from PIL import Image, ImageTk

# GUI Design
root = tk.Tk()
root.title("Desktop Assistant")
root.geometry("800x500")

# Adding a background image
bg_image = Image.open("background.jpg")  # Replace with the path to your background image
from PIL import Image, ImageTk  # Ensure the imports are correct

bg_image = bg_image.resize((800, 500), Image.Resampling.LANCZOS)

bg_photo = ImageTk.PhotoImage(bg_image)

bg_label = tk.Label(root, image=bg_photo)
bg_label.place(relwidth=1, relheight=1)

# Title Frame
title_frame = tk.Frame(root, bg="#ff9800", pady=10, bd=3, relief="ridge")
title_frame.pack(fill="x", pady=10)

title_label = tk.Label(
    title_frame,
    text="My Personal Desktop Assistant",
    font=("Helvetica", 20, "bold"),
    fg="white",
    bg="#ff9800"
)
title_label.pack()

# Output Frame
output_frame = tk.Frame(root, bg="#ffffff", padx=20, pady=20, bd=3, relief="ridge")
output_frame.pack(expand=True, fill="both", pady=10, padx=10)

output_text = tk.StringVar()
output_text.set("Assistant: Hello! How can I help you today?")

output_label = tk.Label(
    output_frame,
    textvariable=output_text,
    font=("Arial", 14),
    fg="#000000",
    bg="#ffffff",
    wraplength=700,
    justify="left"
)
output_label.pack()

# Button Frame
button_frame = tk.Frame(root, bg="#2F4F4F", bd=3, relief="ridge")
button_frame.pack(fill="x", pady=10, padx=10)

# Button Images
listen_img = Image.open("C:\\Users\\Shruti\\PycharmProjects\\pythonProject3\\.venv\\Include\\listen.jpg")  # Replace with the path to your button image
listen_img = listen_img.resize((100, 40), Image.ANTIALIAS)
listen_photo = ImageTk.PhotoImage(listen_img)

exit_img = Image.open("C:\\Users\\Shruti\\PycharmProjects\\pythonProject3\\.venv\\Include\\exit.jpg")  # Replace with the path to your button image
exit_img = exit_img.resize((100, 40), Image.ANTIALIAS)
exit_photo = ImageTk.PhotoImage(exit_img)

# Buttons with Images
listen_button = tk.Button(button_frame, image=listen_photo, command=lambda: Thread(target=listen_command).start(), borderwidth=0)
listen_button.pack(side="left", padx=20)

exit_button = tk.Button(button_frame, image=exit_photo, command=exit_app, borderwidth=0)
exit_button.pack(side="right", padx=20)

# ** Add the greeting functionality here **
def display_greeting():
    greeting = get_greeting()
    output_text.set(f"Assistant: {greeting}")  # Set the greeting message in the GUI
    say_greeting()  # Speak the greeting

# Call the greeting function immediately after GUI setup
display_greeting()

root.mainloop()