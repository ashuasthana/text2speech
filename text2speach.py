#pip install pywin32
import win32com.client as wincl

if __name__ == '__main__':
    print("Welcome to RobotSpeaker 1.1")
    speaker = wincl.Dispatch("SAPI.SpVoice")

    while True:
        x = input("Enter What You want me to speak (or 'q' to quit): ")
        if x.lower() == "q":
            speaker.Speak("You press quit so I have to go...Good bye friend. Have a nice day....")
            break
        speaker.Speak(x)
