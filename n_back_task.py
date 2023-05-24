# This code assumes that your Excel file is named sounds.xlsx and has a column labeled "Sound File" which contains the names of the sound files to be played

import pygame
import pandas as pd
import time
import random

# Load the excel file
df = pd.read_excel('sounds.xlsx')

# Initialize Pygame mixer
pygame.mixer.init()

def play_sound(file):
    pygame.mixer.music.load(file)
    pygame.mixer.music.play()

def n_back_task(n, sounds):
    # Initialize the list of past stimuli
    past_sounds = [None]*n

    start_time = time.time()
    for sound in sounds:
        play_sound(sound)
        time.sleep(1)  # Wait for 1 second while the sound is playing
        pygame.mixer.music.stop()  # Ensure that the sound stops after 1 second

        # Update the list of past stimuli
        past_sounds.pop(0)
        past_sounds.append(sound)

        time.sleep(1)  # Wait for 1 second before playing the next sound

        # Check if 30 seconds have passed
        if time.time() - start_time >= 30:
            break

# Create a list of the blocks, each repeated three times
blocks = [1, 1, 1, 3, 3, 3]
random.shuffle(blocks)

# Shuffle the order of the sounds
sounds = list(df['Sound File'])
random.shuffle(sounds)

# Run the blocks
for block in blocks:
    print("Starting ", block, "-back task")
    n_back_task(block, sounds)
    print("Taking a 10-second break.")
    time.sleep(10)  # Wait for 10 seconds before starting the next block
