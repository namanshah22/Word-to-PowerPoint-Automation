# Automatic Content Transfer from MS Word to MS PowerPoint

This script allows you to copy the contents of a Microsoft Word document and paste them automatically into a new Microsoft PowerPoint presentation without any manual intervention. The script uses the `win32com` module in Python to communicate with the Microsoft Office applications.

## Getting Started

### Prerequisites
- Microsoft Office installed (for Word and PowerPoint) and activated
- Python 3.x
- `win32com` module

### Installation

1. Clone the repository to your local machine.
2. Install the required Python modules by running `pip install -r requirements.txt`.
3. Open the `config.ini` file and set the file paths for the Word document and PowerPoint presentation.
4. Run the script by executing `python main.py` in your terminal.

## Usage

1. Set the file paths for the Word document and PowerPoint presentation in the word_path and ppt_path variables, respectively, at the beginning of the script.
2. Run the script in a Python environment or IDE of your choice.
3. The script will automatically open the Word document, copy its contents, create a new PowerPoint presentation, paste the copied content into a new slide, and save the presentation to the specified file path.

