import speech_recognition as sr
import pyttsx3 as speech
import pywhatkit as py
import wikipedia as wk
import webbrowser
from bs4 import BeautifulSoup
import smtplib
import requests
import datetime
import os
import pyautogui
from time import sleep,time
import subprocess
r=sr.Recognizer()
r.energy_threshold=500
engine=speech.init()
micro=sr.Microphone()
voices=engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)
engine.setProperty('rate', 170)
query=""
def speak(your_voice):
    engine.say(your_voice)
    engine.runAndWait()
def text(your_voice):
    print(your_voice)
def timee():
    p = datetime.datetime.now()
    if int(p.strftime("%H"))<12:
        speak('Hi iam your assistant good morning')
    elif int(p.strftime("%H"))>12:
        speak('Hi iam your assistant good evening')
timee()
dictapp={"commandprompt":"cmd","spotify":"spotify","paint":"paint","excel":"excel","google chrome":"chrome"}
print("tell me")
def openapp(query):
    if "chrome" in query:
        subprocess.call("C://Program Files (x86)//Google//Chrome//Application//chrome.exe")
        speak("here it is")
        sleep(0.3)
    elif "blogging" in query:
        webbrowser.open("http://jntuablogging.liveblog365.com/")
        speak("Launched successfully")
    elif "spotify" in query:
        subprocess.call("spotify.exe")
        sleep(0.3)
    elif "file explorer" in query:
        subprocess.call("File explorer.exe")
def closeapp(query):
    speak("Closing sir!")
    if "one tab" in query or "1 tab" in query:
        pyautogui.hotkey("ctrl","w")
        speak("all tabs are closed")
    elif "2 tab" in query or "1 tab" in query:
        pyautogui.hotkey("ctrl","w")
        sleep(0.5)
        pyautogui.hotkey("ctrl","w")
        speak("all tabs are closed")
    elif "3 tab" in query or "1 tab" in query:
        pyautogui.hotkey("ctrl","w")
        sleep(0.5)
        pyautogui.hotkey("ctrl", "w")
        sleep(0.5)
        pyautogui.hotkey("ctrl","w")
        sleep(0.5)
        speak("all tabs are closed")
    else:
        keys=list(dictapp.keys())
        for app in keys:
            if app in query:
                os.system(f"taskkill /f /im {dictapp[app]}.exe")
def temperature():
    city=query.split("in",1)
    soup=BeautifulSoup(requests.get(f"https://www.google.com/search?q=weather+in+{city[1]}").text,"html.parser")
    region=soup.find("span",class_="BNeawe tAd8D AP7Wnd")
    temp=soup.find("div",class_="BNeawe iBp4i AP7Wnd")
    day=soup.find("div",class_="BNeawe tAd8D AP7Wnd")
    wheather=day.text.split("m", 1)
    temp_r=temp.text.split("C",1)
    your_voice="Its Currently" +wheather[1]+" and "+temp_r[0]+" celcius "+" in "+region.text
    speak(your_voice)
    text(your_voice)
def send_email(to,subject,content):
    sender_email="example@gamil.com"
    password="1225654"
    server=smtplib.SMTP('smtp.gmail.com',587)
    server.starttls()
    server.login(sender_email,password)
    message=f"Subject:{subject}\n\n{content}"
    server.sendmail(sender_email,to,message)
    server.close()
def browser():
    global query
    speak("Please wait Launching... ")
import pandas as pd
# Load the Excel sheet into a pandas DataFrame
# Initialize the text-to-speech engine
def another(det):
    while True:
        if ("yes" in det or "another student" in det or "another name" in det):
            print("User:",det)
            speak("okay tell me the name of the student")
            det=command().lower()
            print(det.split())
            if len((det.split()))!=1:
                try:
                    nm = det.split("is", 1)
                    get_student_details(nm[1][1:])
                    speak("would like to find another name")
                    query=command().lower()
                    another(query)
                except:
                    speak("tell me in the correct format")
                    break
            elif len((det.split()))==1:
                get_student_details(det)
                speak("would you like to find another name")
                det=command().lower()
                another(det)
            else:
                speak("sorry i did not get you")
                break
        elif "no" in det:
            print("okay iam closing")
            speak("okay iam closing")
            break
        else:
            print("could not recognised iam closing")
            speak("could not recognised iam closing")
            break
        break
def get_student_details(name):
    # Search for the row containing the student's name
    df = pd.read_excel("C:/Users/gouth/OneDrive/Desktop/Studentdetails.xlsx")
    student_row = df.loc[df['Name'] == name]
    # If the student is not found, return an error message
    if student_row.empty:
        print('Sorry, the student could not be found.')
        speak("Sorry, the student could not be found.")
        speak("do you want to search another name")
        det=command().lower()
        print(det)
        another(det)
    # Otherwise, extract the details of the student
    else:
        student_details = {
            'Name': student_row['Name'].iloc[0],
            'Age': student_row['Age'].iloc[0],
            'Gender': student_row['Gender'].iloc[0],
            'Grade': student_row['Grade'].iloc[0],
            'Placement': student_row['Placement'].iloc[0],
            'Company': student_row["Company"].iloc[0]
        }
    # Convert the details into a text string
        details_text = f"The student's name is {student_details['Name']}."
        if student_details["Gender"]=="male":
            details_text += f"He is {student_details['Age']} years old and {student_details['Gender']} "
            details_text += f"His grade level is {student_details['Grade']} "
            details_text += f"His placement status is {student_details['Placement']} in {student_details['Company']}."
        else:
            details_text += f"She is {student_details['Age']} years old and {student_details['Gender']} "
            details_text += f"Her grade level is {student_details['Grade']} "
            details_text += f"Her placement status is {student_details['Placement']} in {student_details['Company']}."
        print(details_text)
    # Speak the details using the text-to-speech engine
        speak(details_text)
        speak("do you want to search another name")
        det=command().lower()
        another(det)
# Example usage: Get the details for a student named "Alice"
def command():
    global query
    with micro as source:
        print(".......Listening.......")
        r.adjust_for_ambient_noise(source, duration=1)
        audio=r.listen(source)
        try:
            print("______recognizing_____")
            query=r.recognize_google(audio)
            print("User:",query)
        except:
            your_voice="iam unable to recognize you"
            speak(your_voice)
            return "None"
        return query
if __name__=="__main__":
    while True:
        query=command().lower()
        if "hello" in query:
            speak("Hello ")
        elif "how are you" in query:
            speak("Iam good thank you")
        elif "weather" in query or "temperature" in query :
            temperature()
        elif "bye" in query or "exit" in query:
            speak("Ok Bye!, Iam always there for you")
            exit()
        elif "send email" in query:
                send_email()
        elif "wikipedia" in query:
            query=query.replace("wikipedia","")
            summ=wk.summary(query,sentences=2)
            speak(summ)
            text(summ)
        elif "principal" in query or ("principal" and "jntua") in query:
            print("DISPLAY_IMAGE!D:/Personal/final yr Project/Prof Pics/principle.png")
            print("\nName: Dr.P.Sujatha \nDesignation: Principal, JNTUA CEA")
            speak("\nName: Dr.P.Sujatha \nDesignation: Principal, JNTUA CEA")
            print("Our principal is an outstanding leader, whose vision, dedication, and commitment to excellence")
            speak("Our principal is an outstanding leader, whose vision, dedication, and commitment to excellence")
            sleep(1)
            webbrowser.open("https://www.jntuacea.ac.in/principal_administration.php")
            speak("you can also verify through this official website.")
            break
        elif "time table first year" in query or "first year" in query or ("first year and time table") in query:
            speak("opening")
            print("DISPLAY_IMAGE!D:/Personal/final yr Project/Prof Pics/time table1.png")
            print("First year Time table")
            break
        elif "time table second year" in query or "second year" in query or ("second year and time table") in query:
            speak("opening")
            print("DISPLAY_IMAGE!D:/Personal/final yr Project/Prof Pics/time table1.png")
            print("second year Time table")
            break
        elif "time table third year" in query or "third year" in query or ("third year and time table") in query:
            speak("opening")
            print("DISPLAY_IMAGE!D:/Personal/final yr Project/Prof Pics/time table1.png")
            print("third year Time table")
            break
        elif "time table fourth year" in query or "final year" in query or ("final year" and "time table") in query:
            speak("opening")
            print("DISPLAY_IMAGE!D:/Personal/final yr Project/Prof Pics/time table1.png")
            print("Fourth year Time table")
            break
        elif "electronics" in query or "ece" in query or "ec" in query or "communication" in query or "easy" in query:
            if "hod" in query or "head of the department" in query or "hyd" in query:
                print("DISPLAY_IMAGE!D:/Personal/final yr Project/Prof Pics/hodsir.png")
                print("\nName: Dr.D.Vishnu Vardhan\nDesignation: Associate Professor & Head\nDepartment: Electronics & Communication Engineering")
                speak("Name: Dr.D.Vishnu Vardhan\nDesignation: Associate Professor & Head\nDepartment: Electronics & Communication Engineering")
                text("Dr.D.Vishnu Vardhan is an exceptional educator, whose guidance, knowledge, and enthusiasm have a profound impact on academic")
                speak("Dr.D.Vishnu Vardhan is an exceptional educator, whose guidance, knowledge, and enthusiasm have a profound impact on academic")
                sleep(1)
                speak("you can also verify through this official website")
                webbrowser.open("https://www.jntuacea.ac.in/hodece.php")
                break
            elif "where" in query:
                webbrowser.open("shorturl.at/wyZ34")
                text("here it is in 14.651380 lattitude and 77.606391 longitude")
                speak('here it is in 14.651380 lattitude and 77.606391 longitude')
                sleep(2)
                speak("click get on direction to get the directions")
                break
            elif "ramana" in query or "ramana reddy" in query:
                print("DISPLAY_IMAGE!D:/Personal/final yr Project/Prof Pics/ramana reddy sir.png")
                print("\nName: Dr.P.Ramana Reddy\nDesignation: Professor\nDepartment: Electronics & Communication Engineering")
                speak("\nName: Dr.P.Ramana Reddy\nDesignation: Professor\nDepartment: Electronics & Communication Engineering")
                text("Dr.P.Ramana Reddy is an exceptional educator, whose guidance, knowledge, and enthusiasm have had a profound impact on our academics and personal growth.")
                speak('Dr.P.Ramana Reddy is an exceptional educator, whose guidance, knowledge, and enthusiasm have had a profound impact on our academics and personal growth')
                sleep(1)
                speak("you can also verify through this official website")
                webbrowser.open("https://www.jntuacea.ac.in/pdfs/P.Ramana%20Reddy%20ece.pdf")
            elif "mamatha" in query or "mamata" in query:
                print("DISPLAY_IMAGE!D:/Personal/final yr Project/Prof Pics/mamatha madam.png")
                print("\nName: Dr.G.Mamatha\nDesignation: Assistant Professor\nDepartment: Electronics & Communication Engineering")
                speak("Name: Dr.G.Mamatha\nDesignation: Assistant Professor\nDepartment: Electronics & Communication Engineering")
                text('Dr.G. Mamatha is a remarkable educator, with her passion for teaching,and caring')
                speak('Dr.G. Mamatha is a remarkable educator, with her passion for teaching,and caring')
                sleep(1)
                speak("you can also verify through this official website")
                webbrowser.open("https://www.jntuacea.ac.in/pdfs/Mamatha%20Madam.pdf")
                break
            elif "aruna" in query or "aruna mastani" in query or "mastani" in query:
                print("DISPLAY_IMAGE!D:/Personal/final yr Project/Prof Pics/arunamastanimdm.png")
                print("\nName: Dr.S.Aruna Mastani\nDesignation: Assistant Professor\nDepartment: Electronics & Communication Engineering")
                speak("\nName: Dr.S.Aruna Mastani\nDesignation: Assistant Professor\nDepartment: Electronics & Communication Engineering")
                text("Dr.S.Aruna Mastani is an exceptional educator, whose guidance, knowledge, and enthusiasm have had a profound impact on our academic and personal growth.")
                speak('Dr.S.Aruna Mastani is a remarkable educator, with her passion for teaching,')
                sleep(1)
                speak("you can also verify through this official website")
                webbrowser.open("https://www.jntuacea.ac.in/pdfs/aruna.pdf")
                break
            elif "sumalatha" in query or "suma" in query:
                print("DISPLAY_IMAGE!D:/Personal/final yr Project/Prof Pics/sumalatha mdm.png")
                print("\nName: Dr.V. Sumalatha\nDesignation: Director, Academic & Planning\nDepartment: Electronics & Communication Engineering")
                speak("\nName: Dr.V. Sumalatha\nDesignation: Director, Academic & Planning\nDepartment: Electronics & Communication Engineering")
                text("Dr.V. Sumalatha is an exceptional educator, whose guidance, knowledge, and enthusiasm have had a profound impact on our academics")
                speak('Dr.V. Sumalatha is an exceptional educator, whose guidance, knowledge, and enthusiasm have had a profound impact on our academics')
                sleep(1)
                speak("you can also verify through this official website")
                webbrowser.open("https://www.jntuacea.ac.in/pdfs/Sumalatha.pdf")
                break
            elif "lalitha" in query or "lalitha kumari" in query:
                print("DISPLAY_IMAGE!D:/Personal/final yr Project/Prof Pics/lalithamdm.png")
                print("\nName: Dr.D.Lalitha Kumari\nDesignation: Assistant Professor\nDepartment: Electronics & Communication Engineering")
                speak("\nName: Dr.D.Lalitha Kumari\nDesignation: Assistant Professor\nDepartment: Electronics & Communication Engineering")
                text("Dr.D.Lalitha Kumari is a remarkable educator, with her passion for teaching,and caring")
                speak('Dr.D.Lalitha Kumari is a remarkable educator, with her passion for teaching,and caring')
                sleep(1)
                speak("you can also verify through this official website")
                webbrowser.open("https://www.jntuacea.ac.in/pdfs/lalithakumari.pdf")
                break
            elif ("chandra" or "chandra mohan" or "chandra mohan reddy") in query:
                print("DISPLAY_IMAGE!D:/Personal/final yr Project/Prof Pics/cmrsir.png")
                print("\nName: Dr. S.CHANDRA MOHAN REDDY\nDesignation: Associate Professor \nDepartment: Electronics & Communication Engineering")
                speak("\nName: Dr. S.CHANDRA MOHAN REDDY\nDesignation: Associate Professor \nDepartment: Electronics & Communication Engineering")
                text("Dr. S.CHANDRA MOHAN REDDY is an exceptional educator, whose guidance, knowledge, and enthusiasm have a profound impact on academic")
                speak("Dr. S.CHANDRA MOHAN REDDY is an exceptional educator, whose guidance, knowledge, and enthusiasm have a profound impact on academic")
                sleep(1)
                speak("you can also verify through this official website")
                webbrowser.open("https://www.jntuacea.ac.in/pdfs/CMR.pdf")
                break
            elif ("about" or "ece department" in query) or ("electronics and communication" in query):
                webbrowser.open("https://www.jntuacea.ac.in/aboutece.php")
                speak("here is your results")
                break
        elif "cse" in query or "computer science" in query:
            if "hod" in query or "hyd" in query:
                print("DISPLAY_IMAGE!D:/Personal/final yr Project/Prof Pics/csehod.png")
                webbrowser.open("https://www.jntuacea.ac.in/hodc.php")
                speak("Dr.K. Madhavi")
                break
            elif "where" in query:
                webbrowser.open("https://tinyurl.com/3tvy5p8e")
                text("here it is in 14.651992348186713 lattitude and 77.6079703812279 longitude")
                speak("here it is in 14.651992348186713 lattitude and 77.6079703812279 longitude")
                speak("click get on direction to get the directions")
                break
            elif ("about" or "cse department" in query) or ("computer science" in query):
                webbrowser.open("https://www.jntuacea.ac.in/aboutcse.php")
                speak("here is your results")
                break
        elif "eee" in query or "electrical and electronics" in query:
            if "hod" in query:
                webbrowser.open("https://www.jntuacea.ac.in/hodeee.php")
                speak("Dr. N. Visali")
                break
            elif "where" in query:
                webbrowser.open("https://tinyurl.com/2p95n6x4")
                text("here it is in 14.650758607716535 lattitude and 77.60663913697456 longitude")
                speak("here it is in 14.650758607716535 lattitude and 77.60663913697456 longitude")
                speak("click get on direction to get the directions")
                break
            elif ("about" or "eee department" in query) or ("electrical and electronics" in query):
                webbrowser.open("https://www.jntuacea.ac.in/abouteee.php")
                speak("here is your results")
                break
        elif "who" in query:
            webbrowser.open(query)
        elif "play" in query:
            if "anthem" in query:
                webbrowser.open("https://www.youtube.com/watch?v=oJiATIs8vx0")
                break
            else:
                query=query.replace("play","")
                py.playonyt(query)
                speak("Yeah! its Playing ")
        elif "about jntua" in query or "jawaharlal nehru" in query:
            webbrowser.open("https://www.jntuacea.ac.in/")
            speak("here is your results")
            break
        elif ".com" in query or ".org" in query:
            query = query.replace("search", "")
            query = query.replace("for", "")
            webbrowser.open(query)
        elif "student details" in query or "details" in query:
            nm=query.split("of",1)
            get_student_details(nm[1][1:])
            break
        elif "open" in query:
            openapp(query)
            sleep(3)
        elif "close" in query:
            closeapp(query)
        elif "search" in query:
            webbrowser.open(query)
            speak("searching for that")
            break
        else:
            speak("Sorry command is invalid try again")
            continue