"""
This program takes values from an imported dataset and loads the data into SIP, emulating keyboard inputs.
"""

import pandas as pd
from pynput.keyboard import Key, Controller
from time import sleep

keyboard = Controller()

sleep(2)

delay = 0.08
sleep(delay)

file = r'F:\Projects\CODE\Python\Sip_Loader\bom.xlsx'
newData = pd.read_excel(file)

# print(newData)

for n in range(0, 3):
    numParte = newData.loc[n, 'Numero Parte']
    for character in numParte:
        keyboard.type(character)
        sleep(delay)
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)
