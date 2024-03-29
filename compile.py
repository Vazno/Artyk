import os

os.system(f'''pyinstaller --onefile --add-data "icons;icons" --add-data "models;models" --icon="{os.path.join("icons", "program_icon.ico")}" --name "D2 Research Tool" --noconsole "main.py"''')