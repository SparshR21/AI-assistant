import speech_recognition as sr
import os
import win32com.client
import datetime
import webbrowser
import subprocess
import spacy
import asyncio
from vaderSentiment.vaderSentiment import SentimentIntensityAnalyzer

speaker = win32com.client.Dispatch("SAPI.SpVoice")

current_rate = speaker.Rate
new_rate = current_rate + 2.5
speaker.Rate = new_rate

class NameExtractorAssistant:
    def __init__(self):
        self.nlp = spacy.load("en_core_web_sm")

    def extract_name(self, sentence):
        doc = self.nlp(sentence)
        names = [ent.text for ent in doc.ents if ent.label_ == "PERSON"]
        return names

    def wishme(self):
        hour = datetime.datetime.now().hour

        if 0 <= hour < 12:
            greeting = "Good Morning!"
        elif 12 <= hour < 18:
            greeting = "Good Afternoon!"
        else:
            greeting = "Good Evening!"

        print(greeting)
        # Assuming you have a 'speaker' object and a 'takeCommand' function
        # You may need to adjust these based on your actual implementation
        speaker.Speak(greeting)

        print("I am here to assist you. What should I call you?")
        speaker.Speak("I am here to assist you. What should I call you?")

        name_input = takeCommand()  # Assuming you have a function like takeCommand to get user input
        extracted_names = self.extract_name(name_input)

        if extracted_names:
            name = ', '.join(extracted_names)
        else:
            name = "Sir"

        print(f"Hope you are doing well, {name}\nHow may I help you?")
        speaker.Speak(f"Hope you are doing well, {name}\nHow may I help you?")

async def sleep_for_seconds(seconds):
    print(f"AI Assistant: Sleeping for {seconds} seconds...")
    await asyncio.sleep(seconds)
    print("AI Assistant: Awake now!")

def search_google(query):
    search_url = f"https://www.google.com/search?q={query.replace(' ', '+')}"
    webbrowser.open(search_url)

def quitApp():
    print("Goodbye sir")
    speaker.Speak("Goodbye sir")
    print("Offline")
    exit(0)

def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.pause_threshold = 0.5
        audio = r.listen(source)
        try:
            print("Recognizing")
            query = r.recognize_google(audio, language="en-in")
            print(f"User said: {query}\n")
        except Exception as e:
            print(e)
            print("Unable to Recognize your voice.")
            return "None"

        return query

def analyze_sentiment(text):
    analyser = SentimentIntensityAnalyzer()
    sentiment_scores = analyser.polarity_scores(text)
    compound_score = sentiment_scores['compound']

    if compound_score >= 0.05:
        return 'Positive'
    elif compound_score <= -0.05:
        return 'Negative'
    else:
        return 'Neutral'

def open_application(app_name):
    try:
        subprocess.Popen([app_name], shell=True)
        print(f"Opening {app_name}")
    except Exception as e:
        print(f"Error opening {app_name}: {e}")

# def send_email(subject, body, to_email):
#     # Get user email and password
#     from_email = input("Enter your email address: ")
#     password = input("Enter your email password: ")
#
#     # Set up the email server
#     smtp_server =
#     smtp_port =
#
#     # Compose the email message
#     message = f"Subject: {subject}\n\n{body}"
#
#     try:
#         # Log in to the email server
#         with smtplib.SMTP(smtp_server, smtp_port) as server:
#             server.starttls()  # Use this line for TLS, remove it for SSL
#             server.login(from_email, password)
#
#             # Send the email
#             server.sendmail(from_email, to_email, message)
#
#         print("AI Assistant: Email sent successfully!")
#     except Exception as e:
#         print(f"AI Assistant: Unable to send email. Error: {e}")

if __name__ == '__main__':
    assistant = NameExtractorAssistant()
    assistant.wishme()
    # wishme()
    while True:
        print("listening....")
        query = takeCommand()
        if "how are you" in query.lower() or "how r you" in query.lower():
            print("I'm fine, glad you asked")
            speaker.Speak("I'm fine, glad you asked")

        elif "hello" in query.lower() or "hi" in query.lower():
            print("hey, whatsup")
            speaker.Speak("hey, whatsup")

        elif "what is your name" in query.lower() or "what's your name" in query.lower() or "what should i call you" in query.lower() or "whats your name" in query.lower():
            print("people call me with different names but i like being called assistant")
            speaker.Speak("people call me with different names but i like being called assistant")

        elif "open" in query.lower() and "site" in query.lower() or "website" in query.lower():
            print("Please let me know the domain with the top-level domain.")
            speaker.Speak("Please let me know the domain with the top-level domain.")
            site = takeCommand()
            if site.lower() == "terminate" or site.lower() == "cancel":
                print("Execution terminated.")
                speaker.Speak("Execution terminated.")
            else:
                print("Here is your required site.")
                speaker.Speak("Here is your required site.")
                webbrowser.open('https://' + site)

        elif "search" in query.lower():
            search_query = query.replace("search", "").strip()
            print(f"Searching Google for: {search_query}")
            speaker.Speak(f"Searching Google for: {search_query}")
            search_google(search_query)

        elif "play" in query.lower() :
            song_name = query.lower().replace("play", "").strip()

            if song_name:
                song_name_with_extension = song_name + ".mp3"
                music_folder = 'C:\\Users\\rspar\\Music'  # Replace with the path to your music folder
                song_path = os.path.join(music_folder, song_name_with_extension)

                if os.path.exists(song_path):
                    print(f"Sure! playing {song_name}")
                    speaker.Speak(f"Sure! playing {song_name}")
                    os.startfile(song_path)
                else:
                    print(f"Sorry, {song_name} not found in the playlist. Do you want me to open any other music platform?")
                    speaker.Speak(f"Sorry, {song_name} not found in the playlist. Do you want me to open any other music platform?")
                    user_choice = takeCommand().lower()
                    if user_choice == 'yes' or 'okay' or 'sure':
                        print("Please choose an online music platform (e.g., YouTube, Spotify, etc.)")
                        speaker.Speak("Please choose an online music platform (e.g., YouTube, Spotify, etc.)")
                        platform_choice = takeCommand().lower()
                        if platform_choice == 'terminate' or 'cancel':
                            print("Execution terminated")
                            speaker.Speak("execution terminated")
                        elif platform_choice == 'youtube':
                            # Searching for the song on YouTube
                            search_url = f"https://www.youtube.com/results?search_query={song_name.replace(' ', '+')}"
                            webbrowser.open(search_url)
                        elif platform_choice == 'spotify':
                            # Opening Spotify web player
                            open_application("spotify")
                        else:
                            print("Unsupported platform. Please choose a different platform.")
                            speaker.Speak("Unsupported platform. Please choose a different platform.")
                    else:
                        print("you chose not to search online.")
                        speaker.Speak("You chose not to search online.")
            else:
                print("Sorry, I didn't catch the song name. Please try again.")
                speaker.Speak("Sorry, I didn't catch the song name. Please try again.")

        elif 'the time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            speaker.Speak(f"Sir, the time is {strTime}")

        elif "analyse" in query.lower():
            print("let me know how you are feeling")
            speaker.Speak("let me know how you are feeling")
            text = takeCommand()
            compound_score = analyze_sentiment(text)
            print("sentiment:", compound_score)
            if compound_score == 'Positive':
                print("I'm glad to hear that you're feeling positive!")
                speaker.Speak("I'm glad to hear that you're feeling positive!")
            elif compound_score == 'Negative':
                print("I'm sorry to hear that you're not feeling well. Is there anything I can do to help?")
                speaker.Speak("I'm sorry to hear that you're not feeling well. Is there anything I can do to help?")
            else:
                print("Thanks for sharing. If you need assistance or have any questions, feel free to let me know.")
                speaker.Speak("Thanks for sharing. If you need assistance or have any questions, feel free to let me know.")

        elif "open app" in query.lower():
            print("which one")
            speaker.Speak("which one")
            app = takeCommand()
            open_application(app)

        elif "weather" in query.lower():
            print("To fetch up the weather, i don't have an API KEY to do so!\nSorry for the inconvenience")
            speaker.Speak("To fetch up the weather, i don't have an API KEY to do so!\nSorry for the inconvenience")

        elif "who developed you" in query.lower() or "who made you" in query.lower():
            print("I am being developed by Sparsh Rastogi")
            speaker.Speak("I am being developed by Sparsh Rastogi")

        elif "can you" in query.lower():
            print("Currently i am under development to perform tasks which includes usage of API keys")
            speaker.Speak("Currently i am under development to perform tasks which includes usage of API keys")

        elif "languages" and "you" in query.lower():
            print("The language that i can fetch and speak is the Indian-English language, it depends what language is set for me")
            speaker.Speak("The language that i can fetch and speak is the Indian-English language, it depends what language is set for me")


        elif "stop listening" in query.lower():
            asyncio.run(sleep_for_seconds(5))

        elif "exit" in query.lower() or "bye" in query.lower():
            print("It's a pleasure helping you and I am always here to help you out!")
            speaker.Speak("It's a pleasure helping you and I am always here to help you out!")
            quitApp()



