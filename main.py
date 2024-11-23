import os
import random
import ctypes
import requests
import subprocess
from pathlib import Path
from PySide6.QtWidgets import (QApplication, QMainWindow, QVBoxLayout, QWidget, QComboBox, QPushButton, QLabel,
                               QTextEdit, QHBoxLayout)
from PySide6.QtCore import QTimer, QDateTime
from PySide6.QtGui import QIcon

# Constants
CURRENT_DIR = Path.cwd()
WALLPAPER_DIR = CURRENT_DIR / "Wallpapers"
BING_API_URL = "https://www.bing.com/HPImageArchive.aspx?format=js&idx=0&n=1&mkt=en-US"

# Ensure wallpaper directory exists
WALLPAPER_DIR.mkdir(exist_ok=True)

def get_running_path(relative_path):
    if '_internal' in os.listdir():
        return os.path.join('_internal', relative_path)
    else:
        return relative_path


def set_permanent_wallpaper(wallpaper_path):
    # Set the wallpaper using SystemParametersInfoW (temporarily)
    ctypes.windll.user32.SystemParametersInfoW(20, 0, str(wallpaper_path), 0)

    # PowerShell command to set the wallpaper path in the registry
    ps_command = f'''
    $wallpaperPath = "{wallpaper_path}"
    $regKey = "HKCU:\\Control Panel\\Desktop"
    Set-ItemProperty -Path $regKey -Name Wallpaper -Value $wallpaperPath
    Set-ItemProperty -Path $regKey -Name WallpaperStyle -Value "2"  # 2 = Stretched, 0 = Centered, etc.
    Set-ItemProperty -Path $regKey -Name TileWallpaper -Value "0"    # 0 = No tiling
    '''

    # Execute the PowerShell command to update the registry
    subprocess.run(["powershell", "-Command", ps_command], check=True)

class WallpaperApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("pyDesktopWallpaperRotator v" + open(get_running_path('version.txt')).read())
        self.setGeometry(100, 100, 700, 400)
        self.setWindowIcon(QIcon(get_running_path('icon.ico')))

        # GUI Elements
        self.layout = QVBoxLayout()
        self.refresh_cycle_label = QLabel("Wallpaper Refresh Cycle:")
        self.refresh_cycle_dropdown = QComboBox()
        self.refresh_cycle_dropdown.addItems(["30 minutes", "1 hour", "4 hours", "12 hours"])
        self.refresh_cycle_dropdown.setCurrentIndex(2)  # Default to "4 hours"

        self.download_cycle_label = QLabel("Wallpaper Download Cycle:")
        self.download_cycle_dropdown = QComboBox()
        self.download_cycle_dropdown.addItems(["1 day", "2 days", "3 days", "5 days"])
        self.download_cycle_dropdown.setCurrentIndex(0)  # Default to "1 day"

        self.start_button = QPushButton("ReStart")
        self.start_button.clicked.connect(self.start)
        # Initialize Start button with yellow background and black text, ensuring hovering works as expected
        self.start_button.setStyleSheet("""
            QPushButton {
                background-color: yellow; 
                font-weight: bold; 
                color: black;
            }
            QPushButton:hover {
                background-color: yellow; 
                color: black;
            }
        """)

        # Initialize the new buttons
        self.refresh_button = QPushButton("Refresh wallpaper now")
        self.download_button = QPushButton("Download wallpapers now")

        # Connect the buttons to their respective functions
        self.refresh_button.clicked.connect(self.change_wallpaper)
        self.download_button.clicked.connect(self.download_wallpapers)

        # Horizontal layout for the new buttons
        self.buttons_layout = QHBoxLayout()
        self.buttons_layout.addWidget(self.download_button)
        self.buttons_layout.addWidget(self.refresh_button)

        # Log console
        self.log_console = QTextEdit()
        self.log_console.setReadOnly(True)  # Prevent user edits

        # Add widgets to layout
        self.layout.addWidget(self.refresh_cycle_label)
        self.layout.addWidget(self.refresh_cycle_dropdown)
        self.layout.addWidget(self.download_cycle_label)
        self.layout.addWidget(self.download_cycle_dropdown)
        self.layout.addWidget(self.start_button)
        self.layout.addLayout(self.buttons_layout)
        self.layout.addWidget(self.log_console)

        # Set central widget
        container = QWidget()
        container.setLayout(self.layout)
        self.setCentralWidget(container)

        # Timers
        self.wallpaper_timer = QTimer(self)
        self.download_timer = QTimer(self)

        # start() will be called at the begining
        self.start()

    def log(self, message):
        """Log a message to the console."""
        timestamp = QDateTime.currentDateTime().toString("yyyy-MM-dd HH:mm:ss")
        self.log_console.append(f"[{timestamp}] {message}")

    def start(self):
        # Change Start button to green background with black text, ensuring hovering works as expected
        self.start_button.setStyleSheet("""
                QPushButton {
                    background-color: green; 
                    font-weight: bold; 
                    color: black;
                }
                QPushButton:hover {
                    background-color: green; 
                    color: black;
                }
            """)

        # Starting timers as seconds
        refresh_times = [30 * 60 * 1000, 60 * 60 * 1000, 4 * 60 * 60 * 1000, 12 * 60 * 60 * 1000]
        download_times = [1 * 24 * 60 * 60 * 1000, 2 * 24 * 60 * 60 * 1000, 3 * 24 * 60 * 60 * 1000,
                          5 * 24 * 60 * 60 * 1000]

        refresh_interval = refresh_times[self.refresh_cycle_dropdown.currentIndex()]
        download_interval = download_times[self.download_cycle_dropdown.currentIndex()]

        # Start timers
        self.wallpaper_timer.timeout.connect(self.change_wallpaper)
        self.wallpaper_timer.start(refresh_interval)

        self.download_timer.timeout.connect(self.download_wallpapers)
        self.download_timer.start(download_interval)

        self.log(f"Started with refresh interval {refresh_interval // 60000} minutes "
                 f"and download interval {download_interval // 86400000} days.")

    def change_wallpaper(self):
        wallpapers = list(WALLPAPER_DIR.glob("*.jpg"))
        if wallpapers:
            wallpaper = random.choice(wallpapers)
            # Set the wallpaper and make it permanent
            set_permanent_wallpaper(wallpaper)
            self.log(f"Wallpaper changed to: {wallpaper.name}")
        else:
            self.log("No wallpapers available to set.")

    def download_wallpapers(self):
        self.log("Starting wallpaper download...")
        try:
            # Fetch Bing wallpaper data
            response = requests.get(BING_API_URL)
            response.raise_for_status()
            data = response.json()

            # Extract image URL
            image_url = "https://www.bing.com" + data["images"][0]["url"]
            image_name = data["images"][0]["title"].replace(" ", "_") + ".jpg"
            image_path = WALLPAPER_DIR / image_name

            # Download the image
            with requests.get(image_url, stream=True) as img_response:
                img_response.raise_for_status()
                with open(image_path, "wb") as img_file:
                    for chunk in img_response.iter_content(1024):
                        img_file.write(chunk)

            self.log(f"Downloaded new wallpaper: {image_name}")
        except Exception as e:
            self.log(f"Failed to download wallpaper: {e}")


# Run the app
if __name__ == "__main__":
    app = QApplication([])
    window = WallpaperApp()
    window.show()
    app.exec()
