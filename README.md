# War Sheep

War Sheep is an attempt to extract data from PDFs/TIFs/PNGs (EPRs, Awards, and SURF) to sanitize all data before placing it in new blank files. This will allow leadership to perform EFDP rack-and-stacking while sanitizing identifiable keywords and masking identities

## Add your files

Files should be placed on the user_files section. It expects a subfolder for each person to be sanitized containing all of their docs.

Run the script to map users and randomize. This will read the folder names at top level of user_files and give them all a random User_0, User_1 etc number. It will also attempt to fill some of their personal information such as SSN if it can be determined from their files. REVIEW THIS TO ENSURE THAT DATA IS CORRECT. IF THIS DATA IS WRONG, RUN ADDITIONAL SCRIPT TO ENSURE DATA IS SANITIZED

# Configuration

Setup involves using venv to include libraries

# Setting up local and virtual environment

Log in to OS and run:

    sudo apt install openjdk-8-jre tesseract-ocr imagemagick -y

    virtualenv venv && source venv/bin/activate && pip install -r requirements.txt
