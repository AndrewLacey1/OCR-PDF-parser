# OCR-PDF-parser
This project was designed to automate the process of separating scanned PDFs by each page and renaming them as their serial number.
## Motivation
My employer at the time expressed a desire to automate the process, and assigned me to do it. As this was not personal project, continued maintenance is not currently planned.
##Process
The program begins by prompting the user with a GUI with two inputs fields, the directory to the spreadsheet containing
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


## Usage


