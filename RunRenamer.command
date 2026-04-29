#!/bin/bash

# Navigate to the folder where this double-clicked file is located (your Desktop)
DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
cd "$DIR"

# Activate the virtual environment
source renamer_env/bin/activate

# Run the Python application
python3 RenamerApp.py

# Deactivate the environment after you close the app window
deactivate

# Close the Terminal window automatically
osascript -e 'tell application "Terminal" to close front window'