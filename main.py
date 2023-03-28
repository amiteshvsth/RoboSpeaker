import win32com.client as wincom

if __name__ == '__main__':
    speak = wincom.Dispatch("SAPI.SpVoice")
    while True:
        x = input("Welcome to My Speaker: what do you want me to speak: ")
        speak.Speak(x)
        if x == "q":
            speak.Speak("bye bye")
            break
