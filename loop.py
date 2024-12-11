import os
import csv
import time
import numpy as np
import pymupdf
from pdf2image import convert_from_path
import pytesseract
import re
import cv2

pytesseract.pytesseract.tesseract_cmd = 'C:\\Program Files\\Tesseract-OCR\\tesseract.exe'  # Windows
poppler_path = "C:\\Users\\yzhan457\\Downloads\\Release-24.08.0-0\\poppler-24.08.0\\Library\\bin"
input_folder = "C:\\Users\\yzhan457\\OneDrive - Johns Hopkins\\CMF Lab\\Machine Learning Resident Applications Redacted\\test_input"
output_folder = "C:\\Users\\yzhan457\\OneDrive - Johns Hopkins\\CMF Lab\\Machine Learning Resident Applications Redacted\\test_output"

# Set up CSV file for output
csv_file = os.path.join(output_folder, "extracted_data.csv")
csv_columns = ["Filename", "AAMCID", "Names"]

print("Files in input folder:", os.listdir(input_folder))
print(f"Input folder: {input_folder}")
  
def process_pdfs(input_folder):
    # Create a CSV file to store the results
    with open(csv_file, mode='w', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(csv_columns)

        # Loop through each PDF in the folder
        for pdf_file in os.listdir(input_folder):
            if pdf_file.endswith(".pdf"):
                pdf_path = os.path.join(input_folder, pdf_file)
                
                # Extract names and AAMCID from the document
                names, AAMCID = extract_names_and_aamcid(pdf_path)
                
                if names and AAMCID:
                    print(f"Processing {pdf_file}...")
                    # Redact document and save the output
                    output_pdf = redact_document(pdf_path, names, AAMCID)
                    print(f"Redacted PDF saved as {output_pdf}")
                    
                    # Write data to CSV
                    writer.writerow([pdf_file, AAMCID, ', '.join(names)])  
                    
def extract_names_and_aamcid(pdf_file):
    names = []
    AAMCID = ''
    
    # Convert PDF to images
    pages = convert_from_path(pdf_file, 200, poppler_path=poppler_path, first_page=0, last_page=1)
    page = pages[0]

    # Get the text from the image using Tesseract
    page_string = pytesseract.image_to_string(page)
    page_lines = page_string.split('\n')

    for line in page_lines:
        # Search for AAMCID
        if "".join(line.split()).lower().find("aamcid") != -1:
            AAMCID = ''.join(filter(str.isdigit, line))
        
        # Search for names
        append_flag = False  # This flag controls when to start appending names
        names = []
        
        for word in line.split():
            cleaned_word = re.sub(r'[^a-zA-Z0-9]','', word).lower()
        
            # Start appending names after the word "name"
            if cleaned_word == "name":
                append_flag = True
            elif append_flag:
                # Append the word as a potential name, only if it's not a special keyword
                if cleaned_word not in ["applicant", "aamc", "aamcid"]:
                    names.append(re.sub(r'[^a-zA-Z0-9-]','', word))  # Keep alphanumeric characters and hyphens
        
            # Stop appending if we encounter one of the keywords
            if cleaned_word in ["applicant", "aamc", "aamcid"]:
                append_flag = False

    return names, AAMCID

def redact_document(pdf_file, names, AAMCID):
    # Open the document
    doc = pymupdf.open(pdf_file)
    
    # Load pre-trained Haar cascade XML file for face detection
    face_cascade = cv2.CascadeClassifier(cv2.data.haarcascades + "haarcascade_frontalface_default.xml")

    for page_num, page in enumerate(doc):
        page.clean_contents()
        image_list = page.get_images()

        for image_index, img in enumerate(image_list, start=1):
            # Create Pixmap and perform OCR
            image_rect = page.get_image_rects(img)
            xref = img[0]
            pix = pymupdf.Pixmap(doc, xref)

            if pix.n - pix.alpha > 3:  # CMYK: convert to RGB first
                pix = pymupdf.Pixmap(pymupdf.csRGB, pix)

            image_rectangle, correction_matrix = page.get_image_rects(img, transform=True)[0]
            image_rectangle = np.array(image_rectangle)
            page_rectangle = np.array(page.rect)

            scale_x = (image_rectangle[2] - image_rectangle[0]) / pix.width
            scale_y = (image_rectangle[3] - image_rectangle[1]) / pix.height
            trs_x, trs_y = image_rectangle[0], image_rectangle[1]
            
            shape = (pix.height, pix.width, pix.n)
            array_image = np.frombuffer(pix.samples, dtype=np.uint8).reshape(shape)
            
            # Perform OCR to get image data
            image_data = pytesseract.image_to_data(array_image[:,:,0], output_type='dict')
            
            # Face detection - Detect faces in the image
            if np.shape(array_image)[2] != 1:
                gray = cv2.cvtColor(array_image, cv2.COLOR_BGR2GRAY)
            faces = face_cascade.detectMultiScale(gray, scaleFactor=1.1, minNeighbors=5, minSize=(30, 30))
    
            for i in range(len(faces)):
                x1 = faces[i, 0] * scale_x + trs_x
                x2 = (faces[i, 0] + faces[i, 2]) * scale_x + trs_x
                y1 = faces[i, 1] * scale_y + trs_y
                y2 = (faces[i, 1] + faces[i, 3]) * scale_y + trs_y
                redaction_area = pymupdf.Rect(x1, y1, x2, y2)
                page.add_redact_annot(redaction_area, fill=[0, 0, 0])  # Add face redaction
                
            # Redact names detected by OCR
            for i in range(len(image_data['text'])):
                cleaned_word = re.sub(r'[^a-zA-Z0-9]','',str(image_data['text'][i])).lower()
                for name in names:
                    if cleaned_word == name.lower():
                        x1 = image_data['left'][i] * scale_x + trs_x
                        x2 = (image_data['left'][i] + image_data['width'][i]) * scale_x + trs_x
                        y1 = image_data['top'][i] * scale_y + trs_y
                        y2 = (image_data['top'][i] + image_data['height'][i]) * scale_y + trs_y
                        redaction_area = pymupdf.Rect(x1, y1, x2, y2)
                        page.add_redact_annot(redaction_area, fill=[0, 0, 0])  # Add name redaction
        
        # Redact any name in the text content
        for name in names:
            instances = page.search_for(name)
            for inst in instances:
                page.add_redact_annot(inst, fill=[0, 0, 0])

        # Apply redactions
        page.apply_redactions()

    # Save the redacted PDF with the name and AAMCID
    output_pdf = os.path.join(output_folder, f"{AAMCID}_{names[0]}.pdf")
    doc.save(output_pdf)
    doc.close()
    
    return output_pdf


# Start processing the PDFs
process_pdfs(input_folder)
