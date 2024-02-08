'''
Description
---
This script extracts specific pages from PDF files based on information provided in an Excel file.
It utilizes the PyPDF2 and pandas libraries for PDF manipulation and Excel file processing, respectively.
The user is prompted to select a folder containing PDF files and an Excel file with information about the pages to be extracted.
The extracted pages are saved as individual PDF files in the same folder.

Parameter
---
No direct parameters for the script itself.
It relies on user interaction to select the folder and assumes the presence of an Excel file with two columns.

Returns
---
No direct return value. The script performs file operations and extracts pages based on the provided information, saving the results as new PDF files.

Version
---
1.0
'''
import os
import tkinter as tk
from tkinter import filedialog
import PyPDF2
import pandas as pd
import math
import re

def replace_special_characters(title):
    # Define irregular expression pattern
    muster = re.compile(r'[<>:"/\\|?*]')
    
    # Replace all found special characters with spaces
    title = re.sub(muster, ' ', title)
    
    return title


# Function to extract pages from a PDF file
def extract_pages(pdf_path, start_page, end_page):
    with open(pdf_path, 'rb') as pdf_file:
        pdf_reader = PyPDF2.PdfReader(pdf_file)
        pdf_writer = PyPDF2.PdfWriter()

        # Loop through the specified range of pages and add them to the new PDF
        for page_num in range(start_page - 1, end_page):
            pdf_writer.add_page(pdf_reader.pages[page_num])

        return pdf_writer

# Function to process an Excel file and extract pages from corresponding PDF files
def process_excel(pdf_folder, excel_path):
    # Get a list of PDF files in the specified folder
    pdflist = [file for file in os.listdir(pdf_folder) if file.lower().endswith('.pdf')]
    pdf_path = os.path.join(pdf_folder, pdflist[0])

    # Read data from the Excel file
    df = pd.read_excel(excel_path)
    
    # Iterate through rows in the Excel file
    for index, row in df.iterrows():
        title = row[0]
        pages = row[1]

        # Check if page range is specified in the 'Seiten' column
        if "_" in str(pages): 
            start_page, end_page = map(int, row[1].split('_'))
        else:
            if math.isnan(int(pages)):
                continue
            else:
                start_page = int(pages)
                end_page = int(pages)

        # Extract pages and create a new PDF file for each entry in the Excel file
        new_pdf_writer = extract_pages(pdf_path, start_page, end_page)
        new_pdf_path = os.path.join(pdf_folder, f'{replace_special_characters(title)}.pdf')
        with open(new_pdf_path, 'wb') as new_pdf_file:
            new_pdf_writer.write(new_pdf_file)

# Function to open a file dialog and select a folder
def select_folder():
    folder_selected = filedialog.askdirectory()
    
    # Check if a folder is selected
    if folder_selected:
        excellist = [file for file in os.listdir(folder_selected) if file.lower().endswith('.xlsx')]
        excel_path = os.path.join(folder_selected, excellist[0])

        # Check if the Excel file exists in the selected folder
        if os.path.exists(excel_path):
            # Process the Excel file and extract pages from PDFs
            process_excel(folder_selected, excel_path)
            print("Extraktion abgeschlossen.")  # Extraction completed
        else:
            print("Die Excel-Datei wurde im ausgewählten Ordner nicht gefunden.")
            # The file 'Liederliste.xlsx' was not found in the selected folder.
    else:
        print("Kein Ordner ausgewählt.")  # No folder selected.

# Tkinter GUI
root = tk.Tk()
root.withdraw()  # Hide the main window

# Call the function to select a folder and process the files
select_folder()
