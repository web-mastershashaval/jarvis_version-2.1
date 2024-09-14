from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import re
import pyttsx3

def speak(text):
    engine = pyttsx3.init('sapi5')
    voices = engine.getProperty('voices')
    engine.setProperty('voices', voices[0].id)
    engine.setProperty('rate', 180)  # Moved this up for consistency
    engine.say(text)
    engine.runAndWait()

def google_search(command):
    try:
        reg_ex = re.search('search google for (.*)', command)
        if reg_ex:
            search_for = reg_ex.group(1)
        else:
            search_for = command.split("for", 1)[1].strip()
        
        speak("Okay sir!")
        speak(f"Searching for {search_for}")
        
        driver = webdriver.Chrome(executable_path='C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe')
        driver.get('https://www.google.com')
        
        search = driver.find_element(By.NAME, 'q')
        search.send_keys(search_for)
        search.send_keys(Keys.RETURN)
    
    except Exception as e:
        speak(f"An error occurred: {str(e)}")
    
    finally:
        driver.quit()


