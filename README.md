# IEEE Paper Download Manager

A modern Python desktop application to batch download IEEE research papers from a CSV file, with a beautiful GUI, progress tracking, and robust error handling.

---

## Table of Contents

- [About the Project](#about-the-project)
- [Features](#features)
- [Screenshots](#screenshots)
- [Installation](#installation)
  - [Windows](#windows)
  - [Linux](#linux)
- [Usage](#usage)
- [CSV Format](#csv-format)
- [Author](#author)
- [License](#license)

---

## About the Project

**IEEE Paper Download Manager** is a cross-platform GUI tool that automates the downloading of IEEE research papers using a CSV file as input. It provides real-time progress, pause/resume/stop controls, and categorizes downloads as successful, failed, or skipped. The application is built with Python and [customtkinter](https://github.com/TomSchimansky/CustomTkinter) for a modern look and feel.

---

## Features

- 📂 **CSV-based Batch Download**: Select your CSV, and the app will process all entries.
- 🖥️ **Modern GUI**: Built with customtkinter for a beautiful, responsive interface.
- 🟢 **Success/Failed/Skipped Tabs**: See which papers were downloaded, failed, or already existed.
- 🔄 **Pause, Resume, Stop**: Control your download process at any time.
- 📊 **Live Counters & Progress Bar**: Track total, success, failed, skipped, and progress.
- 🗂️ **Dynamic Folder Structure**: Output folders are named after your CSV file (special characters replaced).
- 🖱️ **Open/Retry/Manual Download**: Open files directly, retry failed downloads, or open failed links in your browser.
- 🏁 **Cross-Platform**: Works on both Windows and Linux.

---

## Screenshots

### Main Interface

![Main Interface](/Screenshot/1.jpeg)

### Download Progress

![Download Progress](/Screenshot/2.jpeg)
---

## Installation

### Windows

1. **Install Python 3.8+** from [python.org](https://python.org).
2. **Install dependencies**:
pip install customtkinter requests
3. **Run the application**:

python ieee_gui.py

text

**Optional: Create a Standalone Executable**
- Install PyInstaller:
pip install pyinstaller

text
- Build the app:
pyinstaller --onefile --windowed --icon=icon.ico ieee_gui.py

text
- Find the `.exe` in the `dist/` folder and double-click to run.

---

### Linux

1. **Install Python 3.8+** (usually pre-installed).
2. **Install dependencies**:
pip install customtkinter requests

text
3. **Run the application**:
python3 ieee_gui.py

text

**Optional: Create a Standalone Executable**
- Install PyInstaller:
pip install pyinstaller

text
- Build the app:
pyinstaller --onefile --windowed ieee_gui.py

text
- The executable will be in the `dist/` folder.
- For a desktop shortcut, create a `.desktop` file pointing to the executable.

---

## Usage

1. **Open the app** and click **Browse CSV** to select your CSV file.
2. The app will auto-generate an output folder based on your CSV name.
3. Click **Start Download** to begin.
4. Use **Pause**, **Resume**, or **Stop** as needed.
5. Monitor the **Downloaded**, **Failed**, and **Skipped** tabs.
6. Use **Open**, **Retry**, or **Open Link** buttons for each entry.

---

## CSV Format

Your CSV should have at least these columns (case-sensitive):

- `Document Title`
- `PDF Link` (the IEEE wrapper or direct PDF URL)
- `Document Identifier` (used for subfolders)

Example:

| Document Title | PDF Link | Document Identifier |
|----------------|----------|--------------------|
| Example Paper  | https://ieeexplore.ieee.org/stamp/stamp.jsp?arnumber=1234567 | Machine_Learning |

---

## Author

**Niraj Kumar**  
- [GitHub](https://github.com/yourusername)
- [LinkedIn](https://www.linkedin.com/in/yourprofile)
- Email: your.email@example.com

---

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

---

> _Inspired by best practices from [Best-README-Template][5], [DhiWise][6], and [Make a README][7]._  
