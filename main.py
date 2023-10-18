import requests
import json
import win32com.client as wincom

speaker = wincom.Dispatch("SAPI.SpVoice")

# speaker.Speak("hello")

city = input("Enter the name of the city :\n")

url = f"https://api.weatherapi.com/v1/current.json?key=9838bb0960784510ad391128230210&q={city}&aqi=no"

r = requests.get(url)

wdic = json.loads(r.text)
w = wdic["current"]["temp_c"]
rain = wdic["current"]["condition"]["text"]
feels = wdic["current"]["feelslike_c"]
last = wdic["current"]["last_updated"]
wind = wdic["current"]["wind_mph"]
humid = wdic["current"]["humidity"]

speaker.Speak(f"the current weather in {city} is {w} degrees but it feels like {feels}  and climate is {rain} ")
speaker.Speak(f"Also the speed of wind is {wind} meters per hour and humidity in air is {humid}")
speaker.Speak(f"this weather forecast is last updated on {last}")

print(r.text)