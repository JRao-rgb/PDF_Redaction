# -*- coding: utf-8 -*-
"""
Task list:
    - check multiple "words" at once, fault-tolerant string comparison
    - use names as a list instead of "first, last". Check against every element in the list
    - output as a .pdf (redact it by marking up the pdf)
    - loop through multiple files, generating "results output" file each time
"""
#%% import statements and definitions
import os
import numpy as np
from PIL import Image as im 
import pymupdf
from pdf2image import convert_from_path
import pytesseract
import time
import fitz
import re

poppler_path = "C:\\Users\\jraos\\Downloads\\Release-24.08.0-0\\poppler-24.08.0\\Library\\bin"
pytesseract.pytesseract.tesseract_cmd = "C:\\Program Files\\Tesseract-OCR\\tesseract.exe"

file_path = "C:\\Users\\jraos\\Downloads\\Atamian.pdf"
os.chdir("C:\\Users\\jraos\\OneDrive - Stanford\\Documents\\Helping Yurui\\PDF_redaction")
dpi_used = 200

#%% obtain relevant text data (names, AAMCID, etc.)

begin_time = time.time()

# Convert PDF to images
pages = convert_from_path(file_path, dpi_used, poppler_path=poppler_path, first_page =0, last_page=1)  # 300 DPI for better quality
page = pages[0]

# get the text
page_string = pytesseract.image_to_string(page)
page_words = page_string.split()
page_lines = page_string.split('\n')

names = []

for line in page_lines:
    # search for AAMCID
    if "".join(line.split()).lower().find("aamcid") != -1:
        AAMCID = ''
        for char in line:
            if char.isdigit():
                AAMCID += char
    
    line_words = line.split()
    # search for the names
    append_flag = False
    for i in range(len(line_words)):
        if append_flag == True:
            names.append(re.sub(r'[^a-zA-Z0-9]','',str(line_words[i])))
        if re.sub(r'[^a-zA-Z0-9]','',str(line_words[i])).lower() == "name":
            append_flag = True
        
#%% redact the document, first using the raw PDF text
doc = pymupdf.open(file_path)

# open the pdf
for page_num, page in enumerate(doc):
    print(page_num)
    if page_num != 43:
        continue
    
    image_list = page.get_images()

    for image_index, img in enumerate(image_list, start=1): # enumerate the image list
        xref = img[0] # get the XREF of the image
        pix = pymupdf.Pixmap(doc, xref) # create a Pixmap

        if pix.n - pix.alpha > 3: # CMYK: convert to RGB first
            pix = pymupdf.Pixmap(pymupdf.csRGB, pix)

        shape = (pix.height, pix.width, pix.n)
        array_image = np.frombuffer(pix.samples, dtype=np.uint8).reshape(shape)
        image_data = pytesseract.image_to_data(array_image, output_type='dict')
        
        break
    
    break
    
    for name in names:
        instances = page.search_for(name)
        # Redact each instance of "Jane Doe" on the current page
        for inst in instances:
            page.add_redact_annot(inst, fill = [0,0,0])

    # Apply the redactions to the current page
    page.apply_redactions()
    
#%% redact the document, now using the extracted image text

doc.save('redacted_document.pdf')

# Close the document
doc.close()