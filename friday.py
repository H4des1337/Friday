import pyttsx3
import datetime
import speech_recognition as sr
import wikipedia
import webbrowser as wb
import os
import random
import pyautogui
import pyjokes
import subprocess
import requests
import json
import time

# ---------- TTS Engine Setup ----------
engine = pyttsx3.init()
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[2].id)
engine.setProperty('rate', 150)
engine.setProperty('volume', 1.0)

def speak(text):
    engine.say(text)
    engine.runAndWait()

# ---------- Mistral AI Integration ----------
def ask_mistral_streaming(prompt):
    try:
        response = requests.post(
            "http://localhost:11434/api/chat",
            json={"model": "mistral", "messages": [{"role": "user", "content": prompt}], "stream": True},
            stream=True
        )

        print("AI:", end=" ")
        full_reply = ""
        buffer = ""
        for line in response.iter_lines():
            if line:
                try:
                    json_data = line.decode("utf-8").removeprefix("data: ")
                    parsed = json.loads(json_data)
                    token = parsed.get("message", {}).get("content", "")
                    print(token, end="", flush=True)
                    full_reply += token
                    buffer += token

                    # Speak sentence-by-sentence
                    if any(p in buffer for p in ['.', '?', '!']):
                        speak(buffer.strip())
                        buffer = ""

                except Exception as e:
                    print("[stream error]", e)
        if buffer:
            speak(buffer.strip())
        print()
    except Exception as e:
        speak("I'm having trouble connecting to my local AI brain.")
        print("Streaming error:", e)

# ---------- Core Assistant Features ----------
def time_now():
    current_time = datetime.datetime.now().strftime("%I:%M:%S %p")
    speak("The current time is " + current_time)
    print("Time:", current_time)

def date_now():
    now = datetime.datetime.now()
    speak("The current date is")
    speak(f"{now.day} {now.strftime('%B')} {now.year}")
    print(f"Date: {now.day}/{now.month}/{now.year}")

def wishme():
    speak("Welcome back sir!")
    hour = datetime.datetime.now().hour
    if 4 <= hour < 12:
        speak("Good morning!")
    elif 12 <= hour < 16:
        speak("Good afternoon!")
    elif 16 <= hour < 24:
        speak("Good evening!")
    else:
        speak("Good night!")
    assistant_name = load_name()
    speak(f"{assistant_name} at your service. Please tell me how may I assist you.")

def screenshot():
    img = pyautogui.screenshot()
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    img_path = f"screenshot_{timestamp}.png"
    img.save(img_path)
    speak(f"Screenshot saved as {img_path}.")
    print(f"Screenshot saved as {img_path}.")

def play_music(song_name=None):
    song_dir = "D:\\code\\friday\\songs"
    songs = os.listdir(song_dir)
    if song_name:
        songs = [s for s in songs if song_name.lower() in s.lower()]
    if songs:
        song = random.choice(songs)
        os.startfile(os.path.join(song_dir, song))
        speak(f"Playing {song}.")
    else:
        speak("No song found.")

def set_name():
    speak("What would you like to name me?")
    name = listen_and_return_text()
    if name:
        with open("assistant_name.txt", "w") as file:
            file.write(name)
        speak(f"Alright, I will be called {name} from now on.")

def load_name():
    try:
        with open("assistant_name.txt", "r") as file:
            return file.read().strip()
    except FileNotFoundError:
        return "Friday"

def search_wikipedia(query):
    try:
        speak("Searching Wikipedia...")
        result = wikipedia.summary(query, sentences=2)
        speak(result)
        print(result)
    except wikipedia.exceptions.DisambiguationError:
        speak("Multiple results found. Please be more specific.")
    except Exception:
        speak("I couldn't find anything on Wikipedia.")

def open_app(app_name: str):
    apps = {
        "chrome": r"C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
        "vs code": r"C:\\Users\\nrish\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe",
        "notepad": r"C:\\Windows\\System32\\notepad.exe",
        "calculator": r"C:\\Windows\\System32\\calc.exe",
        "blender": r"C:\\Program Files\\Blender Foundation\\Blender 4.3\\blender-launcher.exe",
        "steam": r"C:\\Program Files (x86)\\Steam\\Steam.exe",
        "discord": r"C:\\Users\\nrish\\AppData\\Local\\Discord\\Update.exe --processStart Discord.exe",
        "spotify": "start spotify:",
        "whatsapp": "start whatsapp:",
        "microsoft teams": "start teams:",
        "microsoft word": r"C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE",
        "microsoft excel": r"C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE",
        "microsoft powerpoint": r"C:\\Program Files\\Microsoft Office\\root\\Office16\\POWERPNT.EXE",
        "vlc": r"C:\\Program Files\\VideoLAN\\VLC\\vlc.exe",
        "file explorer": r"C:\\Windows\\explorer.exe",
        "task manager": r"C:\\Windows\\System32\\taskmgr.exe",
        "control panel": r"C:\\Windows\\System32\\control.exe",
        "vpn": r"C:\\Program Files\\Proton\\VPN\\ProtonVPN.Launcher.exe",
        "camera": r"C:\\Windows\\System32\\Camera\\Camera.exe",
        "brave": r"C:\\Program Files\\BraveSoftware\\Brave-Browser\\Application\\brave.exe"
    }

    for key in apps:
        if key in app_name:
            try:
                path = apps[key]
                if path.startswith("start"):
                    os.system(path)
                elif "--" in path:
                    subprocess.Popen(path, shell=True)
                else:
                    subprocess.Popen([path])
                speak(f"Opening {key}")
                print(f"Opened {key}")
            except Exception as e:
                speak(f"Failed to open {key}.")
                print(f"Error while opening {key}:", e)
            return
    speak("Sorry, I don't know that application.")

def handle_command(query):
    if "time" in query:
        time_now()
    elif "date" in query:
        date_now()
    elif "wikipedia" in query:
        query = query.replace("wikipedia", "").strip()
        search_wikipedia(query)
    elif "play music" in query:
        song_name = query.replace("play music", "").strip()
        play_music(song_name)
    elif "open youtube" in query:
        wb.open("https://youtube.com")
    elif "open google" in query:
        wb.open("https://google.com")
    elif "change your name" in query:
        set_name()
    elif "screenshot" in query:
        screenshot()
    elif "tell me a joke" in query:
        joke = pyjokes.get_joke()
        speak(joke)
        print(joke)
    elif "shutdown" in query:
        speak("Shutting down the system, goodbye!")
        os.system("shutdown /s /f /t 1")
        exit()
    elif "restart" in query:
        speak("Restarting the system, please wait!")
        os.system("shutdown /r /f /t 1")
        exit()
    elif "exit" in query or "offline" in query:
        speak("Going offline. Have a good day!")
        exit()
    elif "open" in query:
        app_name = query.replace("open", "").strip()
        open_app(app_name)
    else:
        ask_mistral_streaming(query)

def listen_and_process():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening for wake word and command...")
        r.pause_threshold = 1
        try:
            audio = r.listen(source, timeout=5, phrase_time_limit=6)
            query = r.recognize_google(audio, language="en-in").lower()
            print("Heard:", query)

            if query.startswith(("friday", "hey friday", "yo friday")):
                command = query.replace("friday", "", 1).replace("hey", "").replace("yo", "").strip()
                if command:
                    handle_command(command)
                else:
                    speak("Yes, I am here!")
        except sr.WaitTimeoutError:
            pass
        except sr.UnknownValueError:
            pass
        except sr.RequestError:
            speak("Speech recognition service is unavailable.")
        except Exception as e:
            print("Error:", e)

# ---------- Manual Text Listener ----------
def listen_and_return_text():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        try:
            audio = r.listen(source, timeout=5)
            return r.recognize_google(audio, language="en-in").lower()
        except:
            return None

# ---------- Main ----------
if __name__ == "__main__":
    wishme()
    while True:
        try:
            listen_and_process()
        except KeyboardInterrupt:
            print("Assistant terminated by user.")
            break
