# Copyright (C) 2024 Beksultan Artykbaev - All Rights Reserved

import os

from main import APP_NAME

commands = [
    "pyinstaller",
    "--onefile",
    '--add-data "icons;icons"',
    '--add-data "models;models"',
    f'--icon="{os.path.join("icons", "program_icon.ico")}"',
    f'--name "{APP_NAME}"',
    '--noconsole "main.py"'
]

def main():
    os.system(" ".join(commands))

if __name__ == "__main__":
    main()
