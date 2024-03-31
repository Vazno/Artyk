import os

commands = [
    "pyinstaller",
    "--onefile",
    '--add-data "icons;icons"',
    '--add-data "models;models"',
    f'--icon="{os.path.join("icons", "program_icon.ico")}"',
    '--name "D2 Research Tool"',
    '--noconsole "main.py"'
]

def main():
    os.system(" ".join(commands))

if __name__ == "__main__":
    main()
