import requests
import json

from win32com.client import Dispatch
url = 'http://newsapi.org/v2/top-headlines?sources=the-times-of-india&apiKey=d175a5dbf380463a8bdde6cc821aacf7'

news = requests.get(url)
text = news.text
response = json.loads(text)

article = response["articles"]
headlines = []

for items in article:
    headlines.append(items["title"])

for i in range(len(headlines)):
    print(i + 1, headlines[i])

def speak(str):
    speak = Dispatch("SAPI.SpVoice")
    speak.Speak(str)


if __name__ == '__main__':
    speak(headlines)
