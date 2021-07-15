# modules imported
from gtts import gTTS
import speech_recognition as sr
import playsound
from playsound import _playsoundWin
import webbrowser as wb
import datetime
import time as t
import wikipedia as w
import os
from tkinter import *
import smtplib
from email.message import EmailMessage
import math

# use of tkinter module for creating the GUI for the virtual assistant
z = Tk()
z.configure(background='black')
z.geometry("600x400+350+140")
z.title("VIRTUAL ASSISTANT")


# method for converting text to speech
def speak(text):
    tts = gTTS(text=text, lang='en-us', slow=False)
    filename = 'greet.mp3'
    tts.save(filename)
    playsound.playsound(filename)
    os.remove(filename)


# method for converting speech to text
def get_audio():
    c = sr.Recognizer()  # recognizer class
    with sr.Microphone() as source:
        c.adjust_for_ambient_noise(source, duration=0.5)
        c.pause_threshold = 1
        print('LISTENING ...')
        a = c.listen(source)
        query = ""

        try:
            query = c.recognize_google(a, language='en-us').lower()
            print(query)
        except ImportError:
            print("SORRY : CAN'T RECOGNIZE UR VOICE")
        except sr.UnknownValueError:
            print("Could not understand audio.")

    return query


a = []


# greetings message
def hour():
    h = datetime.datetime.now().hour
    if (h >= 0) and (h < 12):
        speak('Good Morning')
    elif (h >= 12) and (h < 16):
        speak('Good afternoon')
    else:
        speak('Good evening')

    speak('I am Natasha. How can i help you.')


hour()


# mathematical calculations functions
def factorial(res):
    fact, val = 1, res
    while res > 0:
        fact = fact * res
        res = res - 1
    speak("The factorial of" + str(val) + "is" + str(fact))


def find_modulo(s):
    c = s.split()
    for i in c:
        if i.isdecimal():
            a.append(i)
        else:
            continue

    for i in range(0, len(a)):
        a[i] = int(a[i])

    m = a[0]
    n = a[1]

    speak("The modulo of " + str(m) + " and " + str(n) + " is " + str(m % n))
    a.clear()


def find_square(s):
    c = s.split()
    for i in c:
        if i.isdecimal():
            r = i
        else:
            continue

    r = int(r)
    f = math.sqrt(r)
    speak("The square root of " + str(r) + " is " + str(round(f, 2)))


def find_cube(s):
    c = s.split()
    for i in c:
        if i.isdecimal():
            p = i
        else:
            continue

    p = int(p)

    cube = 0

    while cube ** 3 < p:
        cube = cube + 1

    if cube ** 3 != p:
        speak("Sorry ! " + str(p) + " is not a perfect cube.")
    else:
        speak("The cube root of " + str(p) + " is " + str(cube))


# using Photoimage class we are importing a .png image to our tkinter GUI
i = PhotoImage(file="ironman-jarvis.png")
photo = Label(z, image=i)
photo.pack(pady=4)


# main interaction between user and virtual assistant starts from here
def talk():
    global c
    while True:

        text = get_audio()

        if 'hi' in text or 'hello' in text or 'how are you doing' in text:
            speak("Hello there ! ")

        elif 'are you a robot' in text or 'are you bot' in text or 'what you are' in text or 'who are you' in text:
            speak("I am Cassandra and I am the result of two innovative minds.")

        elif 'better than alexa' in text or 'better than siri' in text or 'better than google assistant' in text:
            speak('I feel she is quite better than me. She is a nice voice assistant with having lot of more features')

        elif 'who am i' in text or 'who i am' in text or 'my name' in text:
            speak('You are my friend.')

        elif 'ur name' in text or 'what is ur name' in text:
            speak('I am Cassandra!')

        elif 'god' in text or 'who is god' in text or 'what is god' in text:
            speak('I believe that, for me god is you who has gifted me with a life to sustain in this universe.')

# commands placing session
        elif 'show time' in text or 'time' in text or 'current time' in text or 'what is the time now' in text:
            speak("Okay ! So the current time is :" + str(t.strftime(" %r")) + "according to the Indian Standard Time")

        elif 'about' in text:
            speak('Searching Wikipedia for the results.')
            query = text.replace('about', "")
            r = w.summary(text, sentences=2)
            speak('Wikipedia says')
            print(r)
            speak(r)

        elif 'close my computer' in text or 'close' in text or 'shutdown' in text:
            speak("Okay mate ! Shutting down your computer right away")
            os.system('shutdown /s')

        elif 'word' in text or 'open word' in text or 'open microsoft word' in text:
            speak('Okay mate ! Opening word.')
            c = "C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\WINWORD.EXE"
            os.startfile(c)

        elif 'powerpoint' in text or 'open powerpoint' in text or 'open microsoft powerpoint' in text:
            speak('Okay mate ! Opening powerpoint.')
            c = "C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\POWERPNT.EXE"
            os.startfile(c)

        elif 'excel' in text or 'open excel' in text or 'open microsoft excel' in text:
            speak('Okay mate ! Opening excel.')
            c = "C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\EXCEL.EXE"
            os.startfile(c)

        elif 'pycharm' in text or 'open pycharm' in text:
            speak('Okay mate ! Opening pycharm.')
            c = "C:\\Program Files\\JetBrains\\PyCharm Community Edition 2019.2\\bin\\pycharm64.exe"
            os.startfile(c)

        elif 'settings' in text or 'open setting' in text:
            speak('Okay mate ! Opening settings.')
            c = "C:\\WINDOWS\\System32\\Control.exe"
            os.startfile(c)

        elif 'open' in text:
            l = text.lstrip('open')
            m = l.lstrip(' ')
            speak('Okay Mate ! Here we go !')
            wb.open('www.' + m + '.com')

        elif 'music' in text or 'song' in text or 'play song' in text or 'play music' in text:
            speak('Okay! Here we go.')
            p = 'C:\\Users\\ANURAG\\Music'
            s = os.listdir(p)
            if len(s) == 0:
                speak('Sorry ! There is no songs present in your system ! Please download some !')
            else:
                s.remove('desktop.ini')
                print(s)
                for k in range(0, len(s)):
                    _playsoundWin(os.path.join(p, s[k]))
                if get_audio():
                    speak('Yes please!')

# mailing
        elif 'mail' in text or 'email' in text or 'send a mail' in text or 'type a mail' in text:
            speak('Please enter your email address:')
            print("Email ID:")
            s_email = get_audio()
            s_email = s_email.split()
            s_email = "".join(s_email)
            print(s_email)
            speak('Please enter your password:')
            print("Password:")
            p = get_audio()
            p = p.split()
            p = "".join(p)
            print(p)
            speak("Please enter the recipient's email address")
            print("Recipient's email:")
            r_email = get_audio()
            r_email = r_email.split()
            r_email = "".join(r_email)
            print(r_email)
            speak('Please say the subject of the email.')
            print('Subject:')
            sub = get_audio()
            print(sub)
            speak('Please say the content that you want me to write for you.')
            print('Content:')
            content = get_audio()
            print(content)
            msg = EmailMessage()
            msg['Subject'] = sub
            msg["From"] = s_email
            msg["To"] = r_email
            msg.set_content(content)

            with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
                smtp.login(s_email, p)
                smtp.send_message(msg)
            speak('Your email has been sent!')

# mathematical calculations
        elif 'fact' in text or 'factorial' in text:
            r = [int(i) for i in text.split() if i.isdigit()]
            res = int(r[0])
            print(str(res))
            factorial(res)

        elif 'mod' in text or 'modulo' in text or 'modulus' in text:
            s = text
            find_modulo(s)

        elif 'square root' in text:
            s = text
            find_square(s)

        elif 'cube root' in text:
            s = text
            find_cube(s)

# end words
        elif 'quit' in text or 'exit' in text or 'get lost' in text or 'bye' in text or 'goodbye' in text:
            speak("Okay Bye ! See you next time.")
            exit()

        else:
            speak("Sorry! I didn't get you. Can you repeat ?")

b1 = Button(z, bg="black", fg="white", text="TALK TO ME", bd=4, font=("Comic Sans MS", 14), relief=GROOVE, width=12, command=talk)
b1.place(x=215, y=166)

z = mainloop()
