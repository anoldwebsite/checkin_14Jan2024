import os
import re


# import winsound
# from replit import audio
def get_unique_sn():
    valid_inputs = []
    while True:
        user_input = input("Please scan/Type serial number (type 'stop' to stop): ")
        if user_input == "stop":
            break
        elif re.match("^(?=.*[a-zA-Z])(?=.*\d)[a-zA-Z\d]{10}$", user_input):
            if user_input not in valid_inputs:
                valid_inputs.append(user_input)
                print("Valid input: ", user_input)
            else:
                print("Duplicate input. Please enter a unique 10-character alphanumeric string.")
                # winsound.PlaySound("SystemExclamation", winsound.SND_ALIAS)
                os.system("afplay /System/Library/Sounds/Basso.aiff")
                # audio.play_file('Alarm01.wav')  # This line is for repl.it
        else:
            print("Invalid input. Please enter a 10-character alphanumeric string.")
            # winsound.PlaySound("SystemExclamation", winsound.SND_ALIAS)
            os.system("afplay /System/Library/Sounds/Basso.aiff")
            # audio.play_file('Alarm01.wav')  # This line is for repl.it
    return valid_inputs