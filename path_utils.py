# Copyright (C) 2024 Beksultan Artykbaev - All Rights Reserved

import os
import sys

def get_execution_folder() -> str:
    if getattr(sys, 'frozen', False):
        # If the script is running as a bundled executable (e.g., PyInstaller)
        return os.path.dirname(sys.executable)
    else:
        # If the script is running as a standalone .py file
        return os.path.dirname(os.path.realpath(sys.argv[0]))

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)
