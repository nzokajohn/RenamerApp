# 🎬 Media File Renamer Pro

**Version:** 1.3.0 | **OS:** macOS (Highly Optimized)

**Media File Renamer Pro** is a powerful, Python-based desktop application designed to completely automate the tedious process of renaming media files and audio stems for strict Content Management System (CMS) delivery. It enforces strict naming conventions, calculates MD5 checksums, prevents accidental overwrites, and automatically generates organized Excel Delivery Manifests.

---

## ✨ Key Features

### 🗂️ Intelligent Batch Processing
* **Drag-and-Drop:** Drop multiple folders directly into the app for instant batch processing.
* **Auto-Stem Detection:** Automatically assigns content types based on folder names (e.g., folders containing the word "Speech" are automatically assigned to Stem 2).
* **Smart Sorting:** Number duplicate files exactly in the order they were shot or exported by sorting them via **Creation Date**, **Modified Date**, or **Alphabetical** order.
* **Sequence Padding:** Total control over how files are numbered (e.g., `_1`, `_01`, `_001`).

### 🛡️ Enterprise CMS Safety
* **Pre-Flight Checker:** Before touching any files, the app simulates the rename. If a filename exceeds 150 characters, the batch safely aborts.
* **Emoji & Illegal Character Stripper:** Automatically detects and silently vaporizes emojis (🚀, 🔥) and hidden non-ASCII characters that cause CMS databases to crash.
* **Strict Case Formatting:** Force base filenames to `lowercase`, `UPPERCASE`, or `Title Case`.

### 📊 Automated Delivery Manifests (Excel)
Instead of just renaming files, the app generates a `_Rename_Report.xlsx` file containing:
* **Global Bundle Sorting:** Groups all video, speech, and music stems perfectly by bundle.
* **Metadata Extraction (Mac):** Uses native macOS Spotlight (`mdls`) to automatically extract and log the **Duration** of video/audio files.
* **MD5 Checksums:** Generates mathematically unique hashes for file integrity verification.

### ⚙️ Advanced Customization
* **Team Profiles:** Save all your settings, find/replace rules, and custom templates as a "Profile." Use the **Export (⬆️)** and **Import (⬇️)** buttons to share profiles with your team.
* **Dynamic Naming Templates:** Construct custom naming rules using placeholders: `{name}{stem}+lang={lang}&category={cat}&aspect={aspect}&app={app}`
* **Regex Engine:** Utilize Regular Expressions for complex Find & Replace patterns.

### 🔄 1-Click Undo & Auto-Updater
* **Quick Undo:** Made a mistake? Click "Undo Latest Renaming" to safely read the hidden backup logs, revert all filenames from the bottom up, and delete the bad Excel report.
* **Over-The-Air Updates:** Natively checks the GitHub `version.json` file on startup and prompts users to download new updates the second they are published.

---

## 🚀 Installation & Setup

### For End Users (Editors / QA Team)
1. Go to the **[Releases](../../releases)** tab on the right side of this repository.
2. Download the latest `Media File Renamer Pro.zip` file.
3. Extract the app and double-click to run! 
*(Note: You may need to right-click -> Open the first time depending on your Mac's security settings).*

🏗️ Compiling the App (Mac)
If you are modifying the code and want to package a new release for your team:

1. Ensure your virtual environment is active.

2. Update the APP_VERSION variable inside RenamerApp.py.

3. Run the included build script:

chmod +x build_app.command
./build_app.command

4. The standalone .app file will be generated in the dist folder. Zip it and upload it to GitHub Releases!

📡 The Auto-Updater System

This app utilizes a serverless auto-updater. When you push a new release to your team:

1. Upload the new .zip to GitHub Releases.

2. Edit the version.json file in the root of this repository to reflect the new version number and release notes.

3. The next time a team member opens their older version of the app, it will read the JSON file, detect the math discrepancy, and trigger an update popup!

Built with Python, CustomTkinter, and OpenPyXL.

### For Developers (Running from Source)
**Requirements:** Python 3.9+
```bash
# 1. Clone the repository
git clone [https://github.com/nzokajohn/RenamerApp.git](https://github.com/nzokajohn/RenamerApp.git)
cd RenamerApp

# 2. Create and activate a virtual environment
python3 -m venv renamer_env
source renamer_env/bin/activate

# 3. Install required packages
pip install customtkinter tkinterdnd2 openpyxl

# 4. Run the app
python RenamerApp.py


