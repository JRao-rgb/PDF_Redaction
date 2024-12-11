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
import pymupdf
from pdf2image import convert_from_path
import pytesseract
import time
import re
import cv2
import random

pytesseract.pytesseract.tesseract_cmd = 'C:\\Program Files\\Tesseract-OCR\\tesseract.exe'  # Windows
poppler_path = "C:\\Users\\yzhan457\\Downloads\\Release-24.08.0-0\\poppler-24.08.0\\Library\\bin"

file_path = "C:\\Users\\yzhan457\\OneDrive - Johns Hopkins\\2024 Applications\\Conway Brian.pdf"
os.chdir("C:\\Users\\yzhan457\\OneDrive - Johns Hopkins\\CMF Lab\\Machine Learning Resident Applications Redacted\\test_output\\")
input_folder = "C:\\Users\\yzhan457\\OneDrive - Johns Hopkins\\CMF Lab\\Machine Learning Resident Applications Redacted\\test_input\\"
dpi_used = 200

#%%
pdf_files = [f for f in os.listdir(input_folder) if f.endswith('.pdf')]

# List to store the individual name parts
names = []

# Loop over each PDF file
for pdf in pdf_files:
    # Print the original PDF filename
    print(f"Original PDF filename: {pdf}")

    # Clean the filename using regex to remove non-alphanumeric characters (except hyphen and period)
    cleaned_pdf = re.sub(r'[^a-zA-Z0-9\s\-\.]', '', pdf)  # Remove non-alphanumeric characters except space, hyphen, and period
    print(f"Cleaned PDF filename: {cleaned_pdf}")

    # Split the cleaned filename by spaces
    name_parts = cleaned_pdf.replace('.pdf', '').split()

    # Add each part of the name to the names list
    names.extend(name_parts)

# Optionally, print all names added to the list
print("\nAll names extracted:")
for name in names:
    print(name)
#%%
# Define a function to handle quoted names in the filename
def nickname(text):
    # Find any quoted names (e.g., 'Leo', "Leo")
    quoted_names = re.findall(r"'([^']+)'|\"([^\"]+)\"", text)
    # Flatten the list of tuples and filter out empty strings
    return [name for sublist in quoted_names for name in sublist if name]
#%%  
   
for file_name in os.listdir(input_folder):
    if file_name != "Alaniz Leonardo.pdf":
        continue
    file_path = input_folder + file_name
    print("processing ", file_name, "...")
    begin_time = time.time()
    
    # Convert PDF to images
    pages = convert_from_path(file_path, dpi_used, poppler_path=poppler_path, first_page =0, last_page=1)  # 300 DPI for better quality
    page = pages[0]
    
    # get the text
    page_string = pytesseract.image_to_string(page)
    page_words = page_string.split(" ")
    page_lines = page_string.split('\n')
    
    nicked_names = nickname(file_name)
    names.extend(nicked_names)
    print("names: ", names)
    
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
            if re.sub(r'[^a-zA-Z0-9\-]','',str(line_words[i])).lower() == "applicant":
                append_flag = False
            if re.sub(r'[^a-zA-Z0-9\-]','',str(line_words[i])).lower() == "aamc":
                append_flag = False
            if re.sub(r'[^a-zA-Z0-9\-]','',str(line_words[i])).lower() == "aamcid":
                append_flag = False
    
            if append_flag == True:
                names.append(re.sub(r'[^a-zA-Z0-9-\-]','',str(line_words[i])))
            if re.sub(r'[^a-zA-Z0-9\-]','',str(line_words[i])).lower() == "name":
                append_flag = True
    
   
        
    #%%
    if len(names)==0:
        print("image failed to find name, looking at raw text instead")
        doc = pymupdf.open(file_path)
        page = doc[0]
        page_text_raw = page.get_text()
        page_text_raw_by_lines = page_text_raw.split("\n")
        names = page_text_raw_by_lines[0].split()
        
        email = page_text_raw_by_lines[1]
        names.append(email)
        names.append(email.split('@')[0])
        
        for name in names:
            # print(name)
            for name_substring in name.split("-"):
                # print(name_substring)
                names.append(name_substring)  
        print(names)
    
    #%% redact the document, first using the raw PDF text
    doc = pymupdf.open(file_path)
    
    # open the pdf
    for page_num, page in enumerate(doc): 
        # if page_num < 37:
        #     continue
        page.clean_contents()
        
        image_list = page.get_images()
    
        for image_index, img in enumerate(image_list, start=1): # enumerate the image list
            # if image_index != 2:
            #     continue
            
            image_rect = page.get_image_rects(img)
        
            xref = img[0] # get the XREF of the image
            pix = pymupdf.Pixmap(doc, xref) # create a Pixmap
    
            if pix.n - pix.alpha > 3: # CMYK: convert to RGB first
                pix = pymupdf.Pixmap(pymupdf.csRGB, pix)
    
            page_rectangle = pymupdf.Rect(page.rect)
            image_rectangle, correction_matrix = page.get_image_rects(img, transform = True)[0]
    
            image_rectangle = np.array(image_rectangle)
            page_rectangle = np.array(page_rectangle)
            
            i1, j1, i2, j2 = image_rectangle
            x1, y1, x2, y2 = page_rectangle
            
            scale_x = (i2-i1)/pix.width
            scale_y = (j2-j1)/pix.height
            trs_x = i1
            trs_y = j1
            
            shape = (pix.height, pix.width, pix.n)
            array_image = np.frombuffer(pix.samples, dtype=np.uint8).reshape(shape)
            image_data = pytesseract.image_to_data(array_image[:,:,0], output_type='dict')
            
            for i in range(len(image_data['text'])):
                cleaned_word = re.sub(r'[^a-zA-Z0-9]','',str(image_data['text'][i])).lower()
                for name in names:
                    if cleaned_word == name.lower():
                        x1 = image_data['left'][i] * scale_x + trs_x
                        x2 = (image_data['left'][i] + image_data['width'][i]) * scale_x + trs_x
                        y1 = image_data['top'][i] * scale_y + trs_y
                        y2 = (image_data['top'][i] + image_data['height'][i]) * scale_y + trs_y
                        redaction_area = pymupdf.Rect(x1, y1, x2, y2)
                        # redaction_area += pymupdf.Rect(pix.x, pix.y, pix.x, pix.y)
                        # print(image_data['left'][i], image_data['left'][i] + image_data['width'][i], image_data['top'][i], image_data['top'][i] + image_data['height'][i], redaction_area)
                        page.add_redact_annot(redaction_area, fill = [0,0,0])
            
            # Load pre-trained Haar cascade XML file
            face_cascade = cv2.CascadeClassifier(cv2.data.haarcascades + "haarcascade_frontalface_default.xml")
            # convert the image
            if np.shape(array_image)[2] != 1:
                gray = cv2.cvtColor(array_image, cv2.COLOR_BGR2GRAY)
            # Detect faces
            faces = face_cascade.detectMultiScale(gray, scaleFactor=1.1, minNeighbors=5, minSize=(30, 30))
        
            for i in range(len(faces)):
                x1 = faces[i,0] * scale_x + trs_x
                x2 = (faces[i,0] + faces[i,2]) * scale_x + trs_x
                y1 = faces[i,1] * scale_y + trs_y
                y2 = (faces[i,1] + faces[i,3]) * scale_y + trs_y
                redaction_area = pymupdf.Rect(x1, y1, x2, y2)
                # redaction_area += pymupdf.Rect(pix.x, pix.y, pix.x, pix.y)
                page.add_redact_annot(pymupdf.Rect(x1,y1,x2,y2), fill = [0,0,0])
            
        
        for name in names:
            instances = page.search_for(name)
            # Redact each instance of "Jane Doe" on the current page
            for inst in instances:
                page.add_redact_annot(inst, fill = [0,0,0])
    
        # Apply the redactions to the current page
        page.apply_redactions()
        
        # if page_num == 37:
        #     break
    #%% redact the document, now using the extracted image text
    
    document_name = AAMCID + ".pdf"
    
    doc.save(document_name)
    
    # Close the document
    doc.close()
    
    end_time = time.time()
    
    print("written to ", document_name, ". Time elapsed: ", end_time - begin_time)