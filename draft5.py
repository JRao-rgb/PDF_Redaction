# -*- coding: utf-8 -*-
"""
Task list:
    - check multiple "words" at once, fault-tolerant string comparison
    - use names as a list instead of "first, last". Check against every element in the list
    - output as a .pdf (redact it by marking up the pdf)
    - loop through multiple files, generating "results output" file each time
"""
# %% import statements and definitions
import os
import numpy as np
import pymupdf
from pdf2image import convert_from_path
import pytesseract
import time
import re
import cv2
from tqdm import tqdm
from fillpdf import fillpdfs
import xlsxwriter

poppler_path = "C:\\Users\\jraos\\Downloads\\Release-24.08.0-0\\poppler-24.08.0\\Library\\bin"
pytesseract.pytesseract.tesseract_cmd = "C:\\Program Files\\Tesseract-OCR\\tesseract.exe"

input_folder = "C:\\Users\\jraos\OneDrive - Stanford\\Documents\\Helping Yurui\\PDF_redaction\\inputs\\"
input_folder = "C:\\Users\\jraos\OneDrive - Stanford\\Documents\\Helping Yurui\\PDF_redaction\\inputs_single\\"
output_folder = "C:\\Users\\jraos\OneDrive - Stanford\\Documents\\Helping Yurui\\PDF_redaction\\outputs\\"

os.chdir("C:\\Users\\jraos\\OneDrive - Stanford\\Documents\\Helping Yurui\\PDF_redaction")

# convenience parameters
page_begin = 0 # the page that we start at, in case you want to restrict the redaction to only a few pages
page_end = -1  # the page that we end at, in case you want to restrict the redaction to only a few pages
# leave page_end at -1 if you want it to analyze all pages (default)

# "solver" parameters:
dpi_used = 300 # resolution of OCR. The bigger the better, the bigger the slower. Recommend to keep it at 300.
enlarged_bounds = 10 # the number of extra pixels to search for when highlighting specific images
minimum_spacing_needed_before_categories = 50 # in searching for things like name, when do we consider it to end? Units of pixels.
minimum_number_of_redactions_before_lifting_suspicion = 2 # if a page has not enough redactions, consider it to be suspicious.

# export parameters
export_time_stamp = True
export_page_count = True
export_suspicious_pages_count = True
export_redaction_count = True
export_total_names_looked_for = True

pymupdf.TOOLS.set_small_glyph_heights(on=True) # gets rid of text less aggressively than if set to "false"

# %% create "output" file, which contains confidence data, debug data, etc.

all_files_begin_time = time.time()

# deal with potentially duplicate files
performance_summary_file_name = "performance_summary"
copy_nums = []
found_copies = False
for file in os.listdir():
    if file[0:19] != "performance_summary":
        continue
    found_copies = True
    start_id  = file.find("(") + 1
    start_end = file.find(")")
    copy_num  = int(file[start_id:start_end])
    copy_nums.append(copy_num)
    
if found_copies:
    performance_summary_file_name += " (" + str(max(copy_nums)+1) + ").xlsx" 
else:
    performance_summary_file_name += " (0).xlsx"
    
performances_summary = xlsxwriter.Workbook(performance_summary_file_name)
worksheet = performances_summary.add_worksheet()

worksheet.write(0,0, "AAMC ID")
worksheet.write(0,1, "name")
col = 2; row = 1
if export_time_stamp:             worksheet.write(0,col, "Time Stamp (s)");  col += 1
if export_page_count:             worksheet.write(0,col, "Page Count");      col += 1
if export_suspicious_pages_count: worksheet.write(0,col, "Number of Suspicious Pages"); col += 1
if export_redaction_count:        worksheet.write(0,col, "Redaction Count"); col += 1
if export_total_names_looked_for: worksheet.write(0,col, "Names Looked For");

#%% some helpful function definitions

def search_for_sensitive_words(page_data,sensitive_word):
    ID_top_coordinate = []
    ID_bottom_coordinate = []
    found_line = False
    for i in range(len(page_data['text'])):
        if "".join(page_data['text'][i].split()).lower().find(sensitive_word) != -1:
            ID_top_coordinate.append(page_data['top'][i])
            ID_bottom_coordinate.append(
                page_data['top'][i] + page_data['height'][i])
            found_line = True
    return [ID_top_coordinate, ID_bottom_coordinate, found_line]

# %% obtain relevant text data (names, AAMCID, etc.)

for file_num, file_name in enumerate(os.listdir(input_folder)):
    total_name_of_files = len(os.listdir(input_folder))
    print("----------------------------------------------------------")
    print("Analyzing file ", file_num+1, " of ", total_name_of_files, "--", file_name)
    file_path = input_folder + file_name
    
    begin_time = time.time()
    
    # first, flatten all pdfs
    fillpdfs.flatten_pdf(file_path, file_path, as_images=False)
    
    # Convert PDF to images
    pages = convert_from_path(file_path, dpi_used, poppler_path=poppler_path,
                              first_page=0, last_page=1)  # 300 DPI for better quality
    page = pages[0]
    
    # get the text
    page_string = pytesseract.image_to_string(page)
    page_data = pytesseract.image_to_data(
        page, output_type="dict", config="--psm 11")
    page_words = page_string.split(" ")
    page_lines = page_string.split('\n')
    
    names = []
    identifying_info = []
    AAMCID_list = []
    
    # =============================================================================
    # searching for AAMCID with image box technique
    ID_top_coordinate, ID_bottom_coordinate, found_AAMCID_line = \
        search_for_sensitive_words(page_data,"aamc")
    if found_AAMCID_line == True:
        for i in range(len(ID_top_coordinate)):
            sliced_image = np.array(page)[max(ID_top_coordinate[i]-enlarged_bounds, 0):min(ID_bottom_coordinate[i]+enlarged_bounds, np.shape(np.array(page))[0])]
            aamc_containing_string = pytesseract.image_to_string(sliced_image)
            aamc_containing_string = "".join(
                aamc_containing_string.split()).lower()
    
            encountered_numbers = False
            AAMCID = ''
            id_start = aamc_containing_string.find("aamcid")
            for j in range(len(aamc_containing_string[id_start::])):
                if aamc_containing_string[id_start+j].isdigit():
                    AAMCID = aamc_containing_string[id_start+j:id_start+j+8]
                    break
            AAMCID_list.append(AAMCID)
    if all(i == AAMCID_list[0] for i in AAMCID_list):
        print("Image recognition identified AAMCID to be", AAMCID_list[0] + ".")
        AAMCID = AAMCID_list[0]
    else:
        print("Image recognition identified different AAMCIDs. Proceeding with raw text.")
    
    # =============================================================================
    # searching for names with image box technique
    ID_top_coordinate, ID_bottom_coordinate, found_name_line = \
        search_for_sensitive_words(page_data,"name")
    if found_name_line == True:
        for i in range(len(ID_top_coordinate)):
            sliced_image = np.array(page)[max(ID_top_coordinate[i]-enlarged_bounds, 0):min(ID_bottom_coordinate[i]+enlarged_bounds, np.shape(np.array(page))[0]),
                                          :]
            name_containing_data = pytesseract.image_to_data(
                sliced_image, output_type='dict')
            add_the_current_word = False
            for j in range(len(name_containing_data['text'])):
                word = name_containing_data['text'][j]
                if add_the_current_word == True:
                    names.append(re.sub(r'[^a-zA-Z0-9]', '', word).lower())
                    names.append(re.sub(r'[^a-zA-Z0-9-]', '', word).lower())
                    if name_containing_data['left'][min(j+1, len(name_containing_data['text'])-1)] - \
                       name_containing_data['left'][j] - name_containing_data['width'][j] > minimum_spacing_needed_before_categories:
                        add_the_current_word = False
                if re.sub(r'[^a-zA-Z0-9]', '', word).lower() == 'name':
                    add_the_current_word = True
                    # but only add the next word if the next word isn't too far away
                    if name_containing_data['left'][min(j+1, len(name_containing_data['text'])-1)] - \
                       name_containing_data['left'][j] - name_containing_data['width'][j] > minimum_spacing_needed_before_categories:
                        add_the_current_word = False
    
    # ============================================================================
    # searching for emails with image box technique
    ID_top_coordinate, ID_bottom_coordinate, found_email_line = \
        search_for_sensitive_words(page_data,"email")
    if found_email_line == True:
        for i in range(len(ID_top_coordinate)):
            sliced_image = np.array(page)[max(ID_top_coordinate[i]-enlarged_bounds, 0):min(ID_bottom_coordinate[i]+enlarged_bounds, np.shape(np.array(page))[0]),
                                          :]
            email_containing_data = pytesseract.image_to_data(
                sliced_image, output_type='dict')
            add_the_current_word = False
            for j in range(len(email_containing_data['text'])):
                word = email_containing_data['text'][j]
                if add_the_current_word == True:
                    identifying_info.append(re.sub(r'[^a-zA-Z0-9]', '', word).lower())
                    identifying_info.append(re.sub(r'[^a-zA-Z0-9-]', '', word).lower())
                    if email_containing_data['left'][min(j+1, len(email_containing_data['text'])-1)] - \
                       email_containing_data['left'][j] - email_containing_data['width'][j] > minimum_spacing_needed_before_categories:
                        add_the_current_word = False
                if re.sub(r'[^a-zA-Z0-9]', '', word).lower() == 'email':
                    add_the_current_word = True
                    # but only add the next word if the next word isn't too far away
                    if email_containing_data['left'][min(j+1, len(email_containing_data['text'])-1)] - \
                       email_containing_data['left'][j] - email_containing_data['width'][j] > minimum_spacing_needed_before_categories:
                        add_the_current_word = False
    
    if len(names) > 0:
        names = list(set(names))
        print("Image recognition identified elements in the name to be", names)
    else:
        print("Image recognition failed to identify the elements in the name. Proceeding with using raw text instead.")
    
    # %%
    # =============================================================================
    # searching for names using raw text. This is often pretty reliable too, lol.
    if len(names) == 0:
        doc = pymupdf.open(file_path)
        page = doc[0]
        page_text_raw = page.get_text()
        page_text_raw_by_lines = page_text_raw.split("\n")
        names = page_text_raw_by_lines[0].split()
    
        email = page_text_raw_by_lines[1]
        names.append(email)
        names.append(email.split('@')[0])

        print("The following names (or elements of names) were found with raw text.",
              names)
        
    # =============================================================================
    # Break up names with Hyphens, in case that becomes an issue later down the line.
    sub_names = []
    for name in names:
        for name_substring in name.split("-"):
            sub_names.append(name_substring.lower())
    names.extend(sub_names)
    
    if len(names) == 0:
        print("Name identification failed. Using the name of the PDF file instead.")
        
    # =============================================================================
    # searching for names by first using the name of the pdf so we are not moving
    # forward with the analysis with nothing in the names vector
    for word in file_name[0:-4].split():
        names.append(word)
    
    # cleaning up the names list so we have less values to iterate through / search through
    names.extend(identifying_info)
    names = list(set(names))
    while("" in names):
        names.remove("")
    
    # %% redact the document, first using the raw PDF text
    doc = pymupdf.open(file_path)
    
    print("Beginning Redaction.")
    redaction_count = 0
    number_of_suspicious_pages = 0
    
    # open the pdf
    for page_num, page in enumerate(tqdm(doc)):
        if page_num < page_begin:
            continue
        if page_num == page_end:
            break
    
        redaction_count_per_page = 0
        page.clean_contents()
    
        # prepare the page for adding nicknames / middle names
        nicknames = []
        
        # -------------------------------------------------------------------------
        # =========================================================================
        # -------------------------------------------------------------------------
        # NAME REDACTION BY SCOURING IMAGE DATA; ALSO, REDACT FACES
        # first, scour the image data for names and faces:
        # first: if the page doesn't contain any useable text, or there is very little
        # text, then we consider the page to be mainly images. In this case, we render
        # the entire page out as one image and conduct our analysis (which might be faster and better)
    
        # if there are any images, we render out the entire page as a single image, and go through the
        # redaction process.
        image_list = page.get_images()
        if len(image_list) > 0:
            rendered_page = page.get_pixmap(dpi=dpi_used)
            shape = (rendered_page.height, rendered_page.width, rendered_page.n)
            array_image = np.frombuffer(
                rendered_page.samples, dtype=np.uint8).copy().reshape(shape).squeeze()
    
            # correct for image orientation
            osd_result = pytesseract.image_to_osd(array_image, output_type='dict')
            if osd_result['rotate'] == 180:
                osd_result['rotate'] = 0
    
            # must set the rotation to its current rotation + detected misalignment
            page.set_rotation(page.rotation + osd_result['rotate'])
    
            # render out the page again (the rotated version, if it has been rotated)
            if osd_result['rotate'] != 0:
                rendered_page = page.get_pixmap(dpi=dpi_used)
                shape = (rendered_page.height,
                         rendered_page.width, rendered_page.n)
                array_image = np.frombuffer(
                    rendered_page.samples, dtype=np.uint8).copy().reshape(shape).squeeze()
    
            # convert the image in to B & W
            if len(np.shape(array_image)) > 2:
                array_image = cv2.cvtColor(array_image, cv2.COLOR_BGR2GRAY)
    
            # recompute coordinate transforms
            page_rectangle = pymupdf.Rect(page.rect)
            page_rectangle = np.array(page_rectangle)
    
            i1, j1, i2, j2 = page_rectangle
    
            scale_x = (i2-i1)/rendered_page.width
            scale_y = (j2-j1)/rendered_page.height
            trs_x = i1
            trs_y = j1
    
            image_data = pytesseract.image_to_data(
                array_image, output_type='dict', config="--psm 11")
    
            # =====================================================================
            # text data detection for images
            # scour it for single instances of the name
            for i in range(len(image_data['text'])):
                cleaned_word = re.sub(r'[^a-zA-Z0-9]', '',
                                      image_data['text'][i]).lower()
                for name in names:
                    if cleaned_word.find(re.sub(r'[^a-zA-Z0-9]', '', name).lower()) == -1:
                        continue  # if we can't find anything, continue
                    x1 = image_data['left'][i] * scale_x + trs_x
                    x2 = (image_data['left'][i] +
                          image_data['width'][i]) * scale_x + trs_x
                    y1 = image_data['top'][i] * scale_y + trs_y
                    y2 = (image_data['top'][i] +
                          image_data['height'][i]) * scale_y + trs_y
                    redaction_area = pymupdf.Rect(
                        x1, y1, x2, y2) * page.derotation_matrix
                    page.add_redact_annot(redaction_area, fill=[0, 0, 0])
                    redaction_count_per_page += 1
    
                    # detect if there is a word stuck between two "name" words
                    ID_top_coordinate = image_data['top'][i]
                    ID_bottom_coordinate = image_data['top'][i] + image_data['height'][i]
                    words_on_the_same_line = array_image[max(ID_top_coordinate - enlarged_bounds, 0):min(
                        ID_bottom_coordinate + enlarged_bounds, np.shape(array_image)[0]), :]
                    same_line_data = pytesseract.image_to_data(words_on_the_same_line, output_type='dict')
    
                    # now looping through same_line_data and see if we spot any new
                    # "middle names" that can be added to our list
                    for j in range(len(same_line_data['text'])-2):
                        cleaned_word = re.sub(r'[^a-zA-Z0-9]', '', same_line_data['text'][j]).lower()
                        for name in names:
                            if cleaned_word.find(re.sub(r'[^a-zA-Z0-9]', '', name).lower()) == -1:
                                continue  # if the first word isn't a name, continue with checking the next name
                            
                            next_word = same_line_data['text'][j+1]                            
                            
                            # if the word is "Mr", "Ms", "Mrs", or "Dr", skip it
                            if re.sub(r'[^a-zA-Z0-9]', '', next_word).lower() == 'ms' or \
                                re.sub(r'[^a-zA-Z0-9]', '', next_word).lower() == 'mr' or \
                                re.sub(r'[^a-zA-Z0-9]', '', next_word).lower() == 'mrs' or \
                                re.sub(r'[^a-zA-Z0-9]', '', next_word).lower() == 'dr':
                                break  # seeing if this word is a nickname if it's one of the above
    
                            already_a_name = False
                            for name in names:
                                if re.sub(r'[^a-zA-Z0-9]', '', next_word).lower().find(name.lower()) != -1:
                                    already_a_name = True  # skip it if it's already in the names (so something we are already looking for)
                            if already_a_name: break
    
                            two_words_over = same_line_data['text'][j+2]
                            for name in names:
                                # looping through names again to see if the second word
                                # finds anything
                                if re.sub(r'[^a-zA-Z0-9]', '', two_words_over).lower().find(name.lower()) == -1:
                                    continue  # go with the next name if the current one isn't it
    
                                # redact the first instsance of the name
                                x1 = same_line_data['left'][j+1] * scale_x + trs_x
                                x2 = (same_line_data['left'][j+1] + same_line_data['width'][j+1]) * scale_x + trs_x
                                y1 = (same_line_data['top'][j+1] + ID_top_coordinate) * scale_y + trs_y
                                y2 = (same_line_data['top'][j+1] + ID_top_coordinate + same_line_data['height'][j+1]) * scale_y + trs_y
                                redaction_area = pymupdf.Rect(x1, y1, x2, y2) * page.derotation_matrix
                                page.add_redact_annot(redaction_area, fill=[0, 0, 0])
                                redaction_count_per_page += 1
    
                                # redact the second instance of the name
                                x1 = same_line_data['left'][j+2] * scale_x + trs_x
                                x2 = (same_line_data['left'][j+2] + same_line_data['width'][j+2]) * scale_x + trs_x
                                y1 = (same_line_data['top'][j+2] + ID_top_coordinate) * scale_y + trs_y
                                y2 = (same_line_data['top'][j+2] + ID_top_coordinate + same_line_data['height'][j+2]) * scale_y + trs_y
                                redaction_area = pymupdf.Rect(x1, y1, x2, y2) * page.derotation_matrix
                                page.add_redact_annot(redaction_area, fill=[0, 0, 0])
                                redaction_count_per_page += 1
    
                                if len(re.sub(r'[^a-zA-Z0-9]', '', next_word)) <= 1:
                                    continue  # skip appending it if it's only one letter
                                # append the nicknames to the list of things to redact
                                nicknames.append(re.sub(r'[^a-zA-Z0-9]', '', next_word).lower())
                        names.extend(nicknames)
                        names = list(set(names))
                        while("" in names):
                            names.remove("")
                    
    
            # =====================================================================
            # facial data detection for images
            # Load pre-trained Haar cascade XML file
            face_cascade = cv2.CascadeClassifier(cv2.data.haarcascades + "haarcascade_frontalface_default.xml")
    
            # Detect faces
            faces = face_cascade.detectMultiScale(array_image, scaleFactor=1.1, minNeighbors=5, minSize=(30, 30))
    
            for i in range(len(faces)):
                x1 = faces[i, 0] * scale_x + trs_x
                x2 = (faces[i, 0] + faces[i, 2]) * scale_x + trs_x
                y1 = faces[i, 1] * scale_y + trs_y
                y2 = (faces[i, 1] + faces[i, 3]) * scale_y + trs_y
                redaction_area = pymupdf.Rect(x1, y1, x2, y2)
                # redaction_area += pymupdf.Rect(pix.x, pix.y, pix.x, pix.y)
                page.add_redact_annot(pymupdf.Rect(x1, y1, x2, y2), fill=[0, 0, 0])
                redaction_count_per_page += 1
    
        # -------------------------------------------------------------------------
        # =========================================================================
        # -------------------------------------------------------------------------
        # NAME REDACTION USING RAW TEXT DATA THAT'S IN THE PDF ITSELF
        # First, scour the text data for modified versions of the name, and redact those
        # note: this implementation is mildly inefficient, as it performs a new re-
        # daction for every new instance of the name. But seeing as how each name
        # only appears 5-6 times a page max, it doesn't seem that bad. Computers are
        # quite fast, after all, yo.
        # NOTE: WE HAVE SPECIFICALLY GOTTEN RID OF HYPHENS HERE IN THE CASE THAT
        # A NAME CONTAINS HYPHENS IN THE PAGE THAT'S NOT RECOGNIZED BY OUR NAMES
        # LIST, WE APPEND IT TO NAMES
        potentially_modified_names = []
        page_text_raw = page.get_text()
        page_text_raw_by_lines = page_text_raw.split("\n")
        for line in page_text_raw_by_lines:
            for i, word in enumerate(line.split()):
                cleaned_word = re.sub(r'[^a-zA-Z0-9]', '', word).lower()
                for name in names:
                    if cleaned_word.find(re.sub(r'[^a-zA-Z0-9]', '', name).lower()) == -1:
                        continue
                    potentially_modified_names.append(word)
                    instances = page.search_for(word)
                    for inst in instances:
                        page.add_redact_annot(inst, fill=[0, 0, 0])
                        redaction_count_per_page += 1
                names.extend(potentially_modified_names)
                names = list(set(names))
                while("" in names):
                    names.remove("")
    
        # =========================================================================
        # next, scour the text data for initials in the name, and get rid of those
        # too. Also, while we are at it, look for any potential nicknames. We define
        # a nickname as a word that's sandwiched by two "name" words and encased in either
        # brackets () [] {} or quotation marks "" ''
        # the searching for initials is PURELY CONTEXTUAL, while, for nicknames,
        # if we find one, we add it to the list
        page_text_by_words = page_text_raw.split()
        for i, word in enumerate(page_text_by_words):
            cleaned_word = re.sub(r'[^a-zA-Z0-9]', '', word).lower()
            for name in names:
                if cleaned_word.find(re.sub(r'[^a-zA-Z0-9]', '', name).lower()) == -1:
                    continue  # go to the next name if we can't find anything
                # basically, if the current word is a name, and the next word is
                # only a single letter, then we consider that to be an initial
                # of the person, and we redact it too.
                next_word = page_text_by_words[min(i+1, len(page_text_by_words)-1)]
                next_word_contains_1_letters = len(re.sub(r'[^a-zA-Z0-9]', '', next_word)) == 1
                next_word_contains_2_letters = len(re.sub(r'[^a-zA-Z0-9]', '', next_word)) == 2
                last_letter_of_the_next_word_is_a_comma_or_period = any(punctuation in next_word for punctuation in ".,")
                if (next_word_contains_1_letters or next_word_contains_2_letters) \
                        and last_letter_of_the_next_word_is_a_comma_or_period:
                    instances = page.search_for(next_word)
                    for inst in instances:
                        page.add_redact_annot(inst, fill=[0, 0, 0])
                        redaction_count_per_page += 1
                # basically, if the current word is a name, and the word after
                # the word after that is also a name, and the word these two
                # names sandwich is encased in either brackets or quotation marks,
                # we consider this a "nickname" and add it to the redactions list
                two_words_over = page_text_by_words[min(i+2, len(page_text_by_words)-1)]
                cleaned_two_words_over = re.sub(r'[^a-zA-Z0-9]', '', two_words_over).lower()
    
                # uncomment if this takes out too many words because it thinks some random ass word is a nickname
                # contains_brackets = any(punctuation in next_word for punctuation in "()[]{}\"\'“”‘’〈〉《》【】")
    
                # make sure the current word isn't already a name before we spend
                # effort redacting it
                already_a_name = False
                for name in names:
                    if re.sub(r'[^a-zA-Z0-9]', '', next_word).lower().find(name.lower()) != -1:
                        already_a_name = True  # skip it if it's already in the names (so something we are already looking for)
                if already_a_name: break
    
                for name in names:
                    # if cleaned_two_words_over.find(re.sub(r'[^a-zA-Z0-9]','',name).lower()) != -1 and contains_brackets:
                    if cleaned_two_words_over != re.sub(r'[^a-zA-Z0-9]', '', name).lower() or \
                        next_word_contains_1_letters == True or \
                        len(next_word) <= 1:
                        continue  # go to the next name if we don't find  anything
                    if re.sub(r'[^a-zA-Z0-9]', '', next_word).lower() == name:
                        break # we don't need to add this nickname, it's already something
                        # that's being searched for
                    nicknames.append(re.sub(r'[^a-zA-Z0-9]', '', next_word).lower())
                    instances = page.search_for(next_word)
                    for inst in instances:
                        page.add_redact_annot(inst, fill=[0, 0, 0])
                        redaction_count_per_page += 1
                names.extend(nicknames)
                names = list(set(names))
                while("" in names):
                    names.remove("")
             
        #  =========================================================================
        # last, scour the raw text for mentions of the name. This method has the
        # advantage of being able to catch names hyphenated to the next line.
        # we also have the updated names list, so it will catch any nicknames
        # and modified names that it just found, also
        for name in names:
            instances = page.search_for(name)
            # Redact each instance of "Jane Doe" on the current page
            for inst in instances:
                page.add_redact_annot(inst, fill=[0, 0, 0])
                redaction_count_per_page += 1
                    
        # Apply the redactions to the current page
        page.apply_redactions()
    
        # clean up the list
        names.extend(potentially_modified_names)
        names.extend(nicknames)
        names = list(set(names))
        while("" in names):
            names.remove("")
            
        # calculate final redaction count
        redaction_count += redaction_count_per_page
        
        # determine if this page is suspicious
        if redaction_count_per_page < minimum_number_of_redactions_before_lifting_suspicion:
            number_of_suspicious_pages += 1
    
    # %% redact the document, now using the extracted image text
    
    # create the file name and save it
    document_name = output_folder + AAMCID + ".pdf"
    doc.save(document_name)
    doc.close()
    
    # calculate elapsed time
    end_time = time.time()
    all_files_time_elapsed = end_time
    
    print("redacted document written to",  AAMCID +
          ".pdf. Time taken for this file:", end_time - begin_time, "s.")

    worksheet.write(row,0, AAMCID)
    worksheet.write(row,1, file_name[0:-4])
    col = 2
    if export_time_stamp:             
        worksheet.write(row,col, round(all_files_time_elapsed-all_files_begin_time,2)); col += 1
    if export_page_count:             
        worksheet.write(row,col, page_num + 1); col += 1
    if export_suspicious_pages_count:             
        worksheet.write(row,col, number_of_suspicious_pages); col += 1
    if export_redaction_count:        
        worksheet.write(row,col, redaction_count); col += 1
    if export_total_names_looked_for: 
        worksheet.write(row,col, str(names));
    row += 1

# %% export results to operations file
# close the output file
performances_summary.close()

# print final summary statistics
print("=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=")
print("Finished processing all files. Total of", row-1, 
  "files have been processed, average time of",
  round((all_files_time_elapsed-all_files_begin_time)/(row-1),2),"seconds per file.",
  "\nMore performance statistics can be found in '"+performance_summary_file_name+"'.")
print("=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=")
print("Hey bb! Thanks for using my code! MWah <3 MWah <3 MWah <3")