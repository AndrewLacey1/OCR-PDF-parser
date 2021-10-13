# OCR-PDF-parser
This project was designed to automate the process of separating scanned PDFs by each page and renaming them as their serial number.
## Motivation
My employer at the time expressed a desire to automate the process, and assigned me to do it. As this was not personal project, continued maintenance is not currently planned.
## Process
The program begins by prompting the user with a GUI with three input fields:

1-The directory to the folder with the input files

2-The directory to the spreadsheet with known serial numbers

3-The directory to tesseract.exe (Yes, this could have been worked around but do to some technical issues with company computers this was deemed to be easier)

Once the user fills the fields in (all fields have a built in file browser for easier location) and clicks submit the program opens each file one at a time, then runs tesseract on each page to determine the text. Once the text has been determined, an alogrithm is ran to find the serial number contained on the page. The determined serial number is then cross referenced with the spreadsheet to see if it exists, or if there is a very close match. Upon finding a suitable match the page is extracted and saved as its own PDF and name dafter its serial number. If the serial number could not be located it will be indicated in the name.

*NOTE: Due to confidentiality reasons, examples of the files and serial numbers cannot be provided.

## Libraries Used
pdf2image - Used for conversion of pdfs to images for optical character recognition (OCR)
PyPDF2 - Used to load in PDFs

openpyxl - Used to modify an existing excel spreadsheet

difflib - Used to find close matches from the OCR

PySimpleGUI - Used to construct a simple user interface to select various paths

PyInstaller - used to package the program into a format usable by the stakeholders

PIL - Used to open the images such that pytesseract can read them

pytesseract - open source OCR maintained by google used by this program to read serial numbers


## Installation
The tesseract.exe must be downloaded onto the computer at a known location. Download here: https://sourceforge.net/projects/tesseract-ocr-alt/files/tesseract-ocr-3.02.grc.tar.gz/download

Pillow must be downloaded and added to the PATH variable.


## Benefits
Before this process was implemented PDFs had to be manually separated and have their names manually enteres into a database multiple times a year. With this system it reduces the time to complete this task from a few involved hours to a simple click of a button and a run time of about 2 minutes. This runtime is mmostly comprised of running tesseract.exe.

