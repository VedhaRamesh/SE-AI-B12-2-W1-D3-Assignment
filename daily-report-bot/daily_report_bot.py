"""Copy a weather page from Chrome into a new TextEdit document."""

import subprocess
import sys
import time

import pyautogui


WEATHER_URL = "https://wttr.in/SanFrancisco"


def main() -> None:
    if sys.platform != "darwin":
        raise SystemExit("This simple version is for macOS.")

    subprocess.run(["open", "-a", "Google Chrome", WEATHER_URL], check=True)
    time.sleep(8)
    pyautogui.hotkey("command", "a")
    pyautogui.hotkey("command", "c")

    subprocess.run(["open", "-a", "TextEdit"], check=True)
    time.sleep(2)
    pyautogui.hotkey("command", "n")
    time.sleep(1)
    pyautogui.hotkey("command", "v")

    print("Weather page copied into TextEdit.")


if __name__ == "__main__":
    main()