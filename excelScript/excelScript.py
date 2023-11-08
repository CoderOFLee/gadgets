import pandas as pd
from pandas import DataFrame, Series
from PIL import Image
import pyautogui
import pyperclip
import time

SCRIPT_NAME = "ticketGrabbing.xlsx"


class CommandType:
    SINGLE_CLICK = 1
    DOUBLE_CLICK = 2
    RIGHT_CLICK = 3
    INPUT = 4
    WAIT = 5
    SCROLL = 6


def get_col(df: DataFrame, col_index: int) -> list:
    col: Series = df.iloc[:, col_index]
    return col.to_list()


def loads_scripts():
    df: DataFrame = pd.read_excel(SCRIPT_NAME)
    commands: Series = get_col(df, 0)
    contents: Series = get_col(df, 1)
    jump_list: Series = get_col(df, 2)
    return commands, contents, jump_list


def get_pos(img: str):
    image = Image.open(img)
    while pyautogui.locateAllOnScreen(image) is None:
        print("watting...")
    return pyautogui.locateCenterOnScreen(image)


commands, contents, jump_list = loads_scripts()
i = 0
while i < len(commands):
    match commands[i]:
        case CommandType.SINGLE_CLICK:
            x, y = get_pos(contents[i])
            pyautogui.click(x, y, interval=0.2, duration=0.2)
        case CommandType.DOUBLE_CLICK:
            x, y = get_pos(contents[i])
            pyautogui.click(x, y, interval=0.2, duration=0.2, clicks=2)
        case CommandType.RIGHT_CLICK:
            x, y = get_pos(contents[i])
            pyautogui.click(x, y, interval=0.2, duration=0.2, button="right")
        case CommandType.INPUT:
            pyperclip.copy(contents[i])
            pyautogui.hotkey("crtl", "v")
        case CommandType.WAIT:
            time.sleep(contents[i])
        case CommandType.SCROLL:
            pyautogui.scroll(contents[i])
    if pd.isna(jump_list[i]):
        i += 1
    else:
        i = int(jump_list[i])
