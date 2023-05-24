import pygame
import pandas as pd
import time
import random
from win32com.client import Dispatch

# Load the excel file
df = pd.read_excel('sounds.xlsx')

# Initialize Pygame mixer
pygame.mixer.init()

# Establish connection with Oxysoft
Oxysoft = Dispatch('OxySoft.OxyApplication')

def play_sound(file):
    pygame.mixer.music.load(file)
    pygame.mixer.music.play()

def n_back_task(n, sounds, results):
    # Initialize the list of past stimuli
    past_sounds = [None]*n

    start_time = time.time()
    for sound in sounds:
        play_sound(sound)
        time.sleep(1)  # Wait for 1 second while the sound is playing
        pygame.mixer.music.stop()  # Ensure that the sound stops after 1 second

        # Check if the sound n steps back matches the current sound
        if past_sounds[0] == sound:
            correct_answer = 1
        else:
            correct_answer = 0

        # Update the list of past stimuli
        past_sounds.pop(0)
        past_sounds.append(sound)

        time.sleep(2)  # Wait for 2 seconds before playing the next sound

        # Check if 30 seconds have passed
        if time.time() - start_time >= 30:
            break

        # Save the result
        results.append((n, sound, correct_answer))

# Ask for participant number
participant_number = input("Please enter the participant number: ")

# Create a list of the blocks, each repeated three times
blocks = [1, 1, 1, 3, 3, 3]
random.shuffle(blocks)

# Shuffle the order of the sounds
sounds = list(df['Sound File'])
random.shuffle(sounds)

# List to store the results
results = []

# Break before the first block
print("Taking a 10-second break.")
play_sound("sounds/rest.wav")
time.sleep(random.randint(10, 15))  # Wait for a random amount of time between 10 and 15 seconds before starting the first block
pygame.mixer.music.stop()

# Run the blocks
for block in blocks:
    if block == 1:
        play_sound("sounds/oneback.wav")
        Oxysoft.WriteEvent('O', 'one_back') 
        print("Event marker 'one_back' sent.")
    else:
        play_sound("sounds/threeback.wav")
        Oxysoft.WriteEvent('T', 'three_back')
        print("Event marker 'three_back' sent.")
    time.sleep(1)
    pygame.mixer.music.stop()

    print("Starting ", block, "-back task")
    n_back_task(block, sounds, results)
    print("Taking a break.")
    play_sound("sounds/rest.wav")
    time.sleep(random.randint(10, 15))  # Wait for a random amount of time between 10 and 15 seconds before starting the next block
    pygame.mixer.music.stop()

# Convert the results to a DataFrame and save it to an Excel file
results_df = pd.DataFrame(results, columns=['Block', 'Sound', 'Correct Answer'])
results_df.to_excel('results_P' + participant_number + '_' + time.strftime("%Y%m%d-%H%M%S") + '.xlsx', index=False)
