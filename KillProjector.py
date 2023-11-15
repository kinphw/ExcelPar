import pyautogui
import time
import random

while True:
    pyautogui.sleep(3)

    x = random.randrange(1,500)
    y = random.randrange(1,500)
    pyautogui.moveTo(x,y)
    pyautogui.rightClick()