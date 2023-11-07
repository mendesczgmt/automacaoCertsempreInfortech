

import pyautogui
import time

img = pyautogui.locateCenterOnScreen('Captura.png', confidence=0.9)

pyautogui.click(img.x, img.y)

