"""
Jarvis Voice Assistant
Author: Villan
Description:
A Python-based voice assistant that can recognize commands, tell jokes,
play local music, report the time/date, search Wikipedia, and talk to ChatGPT.
"""

import os
import random
import datetime
import webbrowser
import traceback
import re

import speech_recognition as sr
import pyjokes
import wikipedia
import openai
import pythoncom
import win32com.client

# ------------------- Config -------------------
WAKE_WORD = "jarvis"
MUSIC_DIR = r"C:\path\to\your\music"  # Change this to your folder path
DEFAULT_CHAT_MODEL = "gpt-3.5-turbo"

# Initialize TTS engine (SAPI)
pythoncom.CoInitialize()
speaker = win32com.client.Dispatch("SAPI.SpVoice")


def speak(text: str):
    """Speak the given text out loud."""
    print(f"SPEAKING: {text}")
    speaker.Speak(text)


def ask_chatgpt(prompt: str) -> str:
    """Send a user prompt to ChatGPT and return the response."""
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        speak("OpenAI API key not set. Please set the OPENAI_API_KEY environment variable.")
        return "API key missing."
    openai.api_key = api_key
    try:
        response = openai.chat.completions.create(
            model=DEFAULT_CHAT_MODEL,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.7,
            max_tokens=700
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        print("OpenAI error:", e)
        traceback.print_exc()
        speak("Sorry, I couldn't get an answer from ChatGPT.")
        return "Error contacting ChatGPT."


def get_time() -> str:
    return datetime.datetime.now().strftime("%H:%M:%S")


def get_date() -> str:
    return datetime.date.today().strftime("%B %d, %Y")


def listen_once(timeout=5, phrase_time_limit=7) -> str:
    """Listen for a single phrase and return recognized text."""
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Adjusting for ambient noise...")
        r.adjust_for_ambient_noise(source, duration=1)
        print("Listening...")
        try:
            audio = r.listen(source, timeout=timeout, phrase_time_limit=phrase_time_limit)
        except sr.WaitTimeoutError:
            print("Timeout waiting for phrase.")
            return ""
    try:
        text = r.recognize_google(audio).lower()
        print(f"Recognized: {text}")
        return text
    except sr.UnknownValueError:
        print("Could not understand audio.")
        return ""
    except sr.RequestError as e:
        print(f"Speech Recognition error: {e}")
        speak("Speech recognition service unavailable.")
        return ""


def open_website(url: str):
    """Open a website in the default browser."""
    try:
        webbrowser.open(url)
        speak(f"Opening {url}")
    except Exception as e:
        print("Error opening website:", e)
        speak("Sorry, I couldn't open the website.")


def play_random_music(folder_path: str):
    """Play a random song from the given folder."""
    if not os.path.exists(folder_path):
        speak("Music folder not found.")
        return
    songs = [f for f in os.listdir(folder_path)
             if f.lower().endswith(('.mp3', '.wav', '.aac', '.flac', '.ogg'))]
    if not songs:
        speak("No music found in the folder.")
        return
    song = random.choice(songs)
    speak(f"Playing {os.path.splitext(song)[0]}")
    os.startfile(os.path.join(folder_path, song))


def search_wikipedia(topic: str):
    """Search for a topic on Wikipedia and read the summary."""
    if not topic.strip():
        speak("What do you want to search on Wikipedia?")
        topic = listen_once()
    try:
        speak("Searching Wikipedia...")
        summary = wikipedia.summary(topic, sentences=2)
        speak("According to Wikipedia,")
        print(summary)
        speak(summary)
    except Exception:
        speak("Sorry, I couldn't find information on that topic.")


def wish_me():
    hour = datetime.datetime.now().hour
    if 0 <= hour < 12:
        speak("Good Morning!")
    elif 12 <= hour < 18:
        speak("Good Afternoon!")
    else:
        speak("Good Evening!")
    speak(f"I am your assistant. Say '{WAKE_WORD}' to wake me.")


def words_to_math(expr: str) -> str:
    """Convert spoken math words into symbols."""
    expr = expr.lower()
    replacements = {
        "plus": "+", "minus": "-", "times": "*", "divided by": "/", "over": "/", "mod": "%", "power of": "**"
    }
    for word, symbol in replacements.items():
        expr = expr.replace(word, symbol)
    numbers_map = {
        "zero": "0", "one": "1", "two": "2", "three": "3", "four": "4",
        "five": "5", "six": "6", "seven": "7", "eight": "8", "nine": "9", "ten": "10"
    }
    for word, digit in numbers_map.items():
        expr = re.sub(r'\b'+word+r'\b', digit, expr)
    expr = re.sub(r'[^0-9+\-*/%.() ]', '', expr)
    return expr.strip()


def calculate_expression(expression: str) -> str:
    """Evaluate a math expression safely."""
    expr = expression.replace(" ", "")
    try:
        result = eval(expr, {"__builtins__": None}, {})
        return f"The result is {result}"
    except Exception:
        return "Sorry, I couldn't calculate that."


def process_command(q: str) -> bool:
    """Process recognized user commands."""
    math_expr = words_to_math(q)
    if any(op in math_expr for op in "+-*/%"):
        speak(calculate_expression(math_expr))
        return True

    if "youtube" in q:
        open_website("https://youtube.com")
    elif "google" in q:
        open_website("https://google.com")
    elif "play music" in q:
        play_random_music(MUSIC_DIR)
    elif "time" in q:
        speak(f"The time is {get_time()}")
    elif "date" in q:
        speak(f"Today is {get_date()}")
    elif "joke" in q:
        speak(pyjokes.get_joke())
    elif "wikipedia" in q:
        search_wikipedia(q.replace("wikipedia", "").strip())
    elif "ask chatgpt" in q or "ask ai" in q:
        speak("What would you like to ask ChatGPT?")
        user_prompt = listen_once(timeout=20, phrase_time_limit=40)
        if user_prompt:
            speak("Thinking...")
            speak(ask_chatgpt(user_prompt))
    elif any(x in q for x in ["exit", "quit", "stop listening"]):
        speak("Goodbye!")
        return False
    else:
        speak("I don't have a command for that yet.")
    return True


def main():
    wish_me()
    try:
        while True:
            phrase = listen_once(timeout=15, phrase_time_limit=4)
            if WAKE_WORD in phrase:
                speak("Yes?")
                active = True
                while active:
                    command = listen_once()
                    if not command:
                        speak("I didn't catch that. Please say it again.")
                        continue
                    active = process_command(command.lower())
    except KeyboardInterrupt:
        speak("Shutting down. Bye.")


if __name__ == "__main__":
    main()
