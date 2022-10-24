#!/usr/bin/env python3

# Need to pip install pdfplumber, progress,
# python-docx, deep-translator, openpyxl, 
# xlrd, textract, striprtf
import glob, os.path, sys, pdfplumber
import docx, xlrd, textract
import pandas as pd
from itertools import islice
from multiprocessing import Pool, get_context
from progress.bar import Bar
from striprtf.striprtf import rtf_to_text
from pathlib import Path
from deep_translator import GoogleTranslator
from openpyxl import load_workbook
from PIL import Image
from pytesseract import image_to_string
import os, docx2txt
from pdf2image import convert_from_path

# also must run 'sudo apt-get install antiword' to install antiword for textract
# install tesseract and the Russian language traindata
# on linux, this is done with `sudo apt install tesseract-ocr` and 
# `sudo apt install tesseract-ocr-rus`

# also install poppler outside of Python: https://pdf2image.readthedocs.io/en/latest/installation.html

# Used in main and updateBar, but not child processes
bar = None

TESSERACT_PATH = "./usr/bin/tesseract"

# Takes a path to a single pdf, saves contents as
# "foo.pdf.txt" in the current directory
def processPDF(pdf):
    basename = os.path.basename(pdf)
    #try:
    with pdfplumber.open(pdf) as pdfdoc:
        with open(basename + ".txt", "w+") as txt:
            for page in pdfdoc.pages:
                text = page.extract_text()
                #print("text",text)
                if text != None:
                        transl_text = GoogleTranslator(source="ru", target="en").translate(text=text)
                        txt.write(transl_text + "\n")

    # use arbitrary, small cutoff for byte size of txt file
    # to determine whether PDF was read successfully and decide
    # to try OCR
    if os.path.getsize(basename + ".txt") < 20:
        print("Processing PDF with OCR..")
        jpg_files = PDFtoJPG(pdf)
        processJPG(jpg_files, basename)
        #return
    #except:
    #return ("! - Error parsing '%s', skipping..." % basename)


def PDFtoJPG(pdf):
    """
    Convert PDFs to JPG files for processing with OCR
    """
    basename = os.path.basename(pdf)
    pdf_imgs = convert_from_path(pdf)
    jpg_files = []
    for i in range(len(pdf_imgs)):
        filename = pdf+str(i)+'.jpg'
        pdf_imgs[i].save(filename, "JPEG")
        jpg_files.append(filename)
    return jpg_files


def processJPG(jpgs, basename=None):
    """
    Process either a single jpg or list of jpg images
        and store the results as a single text file
    """
    if basename:
        output_file = basename + ".txt"
        # delete files if they were created from PDF for OCR
        del_jpgs = True
    else:
        output_file = os.path.basename(jpgs) + ".txt"
        print("output_file",output_file)
        del_jpgs = False
    with open(output_file, "w+") as f:
        if isinstance(jpgs, str):
            jpgs = [jpgs]
        for img_file in jpgs:
            text = str(((image_to_string(Image.open(img_file), lang="rus"))))
            transl_text = GoogleTranslator(source="ru", target="en").translate(text=text)
            if del_jpgs:
                os.remove(img_file)
            f.write(transl_text)  

# Takes a path to a single .docx, saves contents as
# "foo.docx.txt" in the current directory
def processDOCX(docx_file):
    basename = os.path.basename(docx_file)
    try:
        document = docx.Document(docx_file)
        with open(basename + ".txt", "w+") as txt:
            for para in document.paragraphs:
                text = para.text
                #print("text",text)
                if( text != None ):
                    transl_text = GoogleTranslator(source="ru", target="en").translate(text=text)
                    txt.write(transl_text + "\n")
        return
    except:
        return ("! - Error parsing '%s', skipping..." % basename)

# Takes a path to a single .doc, saves contents as
# "foo.doc.txt" in the current directory
def processDOC(doc_file):
    basename = os.path.basename(doc_file)
    try:
        text = textract.process(doc_file, language="rus").decode("utf-8")
        #with open(doc_file, "rb") as doc_orig:
        if( text != None ):
            with open(basename + ".txt", "w+") as txt:
                transl_text = GoogleTranslator(source="ru", target="en").translate(text=text)
                txt.write(transl_text + "\n")
        return
    except:
        return ("! - Error parsing '%s', skipping..." % basename)

def get_doc_text(file):
    #print("file",file)
    filepath = "../../Downloads/Roskomnadzor/"
#    try:
    if file.endswith('.docx'):
        text = docx2txt.process(file)
        return text
    elif file.endswith('.doc'):
       # converting .doc to .docx
        doc_file = filepath + file
        docx_file = file + 'x'
        if not os.path.exists(docx_file):
            #print("using antiword to write",doc_file, "to",docx_file)
            os.system('antiword -m "8859-5.txt" ' + doc_file + ' > ' + docx_file)
            #print("trying to read docx file")
            with open(docx_file) as f:
                print("reading docx file",docx_file)
                text = f.read()
                if( text != None ):
                    with open(file + ".txt", "w+") as txt:
                        transl_text = GoogleTranslator(source="ru", target="en").translate(text=text)
                        txt.write(transl_text + "\n")
            os.remove(docx_file) #docx_file was just to read, so deleting
        else:
            # already a file with same name as doc exists having docx extension, 
            # which means it is a different file, so we cant read it
            print('Info : file with same name of doc exists having docx extension, so we cant read it')
            text = ''
    return
#    except:
#        return ("! - Error parsing '%s', skipping..." % basename)

# Takes a path to a single .rtf, saves contents as
# "foo.rtf.txt" in the current directory
def processRTF(rtf):
    basename = os.path.basename(rtf)
    try:
        rtf_path = Path.cwd() / rtf
        with rtf_path.open() as source:
            text = rtf_to_text(source.read())
            #print(text)    
            if( text != None ):
                with open(basename + ".txt", "w+") as txt:
                    transl_text = GoogleTranslator(source="ru", target="en").translate(text=text)
                    txt.write(transl_text + "\n")
        return
    except:
        return ("! - Error parsing '%s', skipping..." % basename)


# Takes a path to a single .xls, saves contents as
# "foo_en.xls_X.csv" (where X is the sheet number) in the current directory
def processXLS(xls):
    basename = os.path.basename(xls)
    try:
        wb = xlrd.open_workbook(xls)
        for i in range(wb.nsheets):
            new_csv = basename + "_" + str(i) + ".csv"
            #print("new csv",new_csv)
            wb_sh = wb.sheet_by_index(i)
            rows = wb_sh.get_rows()
            transl_rows = []
            transl_headers = []
            i = 0
            for row in rows:
                #print("row", row)
                row = [str.value for str in row]
                #print("row stripped",row)
                transl_row = GoogleTranslator(source="ru", target="en").translate_batch(row)
                #print(transl_row)
                if i == 0:
                    transl_headers = transl_row
                    i += 1
                else:
                    transl_rows.append(transl_row)
            transl_df = pd.DataFrame(transl_rows, columns=transl_headers)
            #print(transl_df.head())
            #print(transl_df.shape)
            transl_df.to_csv(new_csv)
        return
        
    except:
        return ("! - Error parsing '%s', skipping..." % basename)


def updateBar(error=None):
    bar.next()
    if( error != None ):
        sys.stderr.write(error + "\n")


# Loop over all relevant files in a folder, ingest them in background processes
if __name__ == '__main__':
    #file_location = "C:/Users/ARK Silverlining/Downloads/Roskomnadzor"
    file_location = "/mnt/c/Users/natra/Documents/Research/Gary/sample_russian_docs"
    # pdfs
    """
    pdfs = glob.glob(file_location + "/*pdf")
    bar = Bar("Extracting text from PDFs", max=len(pdfs))
    pool = get_context("spawn").Pool()
    for pdf in pdfs:
        pool.apply_async(processPDF, [pdf], callback=updateBar)
    pool.close()
    pool.join()
    bar.finish()
    """
    # now jpgs
    jpgs = glob.glob(file_location + "/*jpg")
    jpgs.extend(glob.glob(file_location + "/*JPG"))
    bar = Bar("Extracting text from JPGs", max=len(jpgs))
    pool = get_context("spawn").Pool()
    for jpg in jpgs:
        pool.apply_async(processJPG, [jpg], callback=updateBar)
    pool.close()
    pool.join()
    bar.finish()

    # now pngs
    pngs = glob.glob(file_location + "/*png")
    pngs.extend(glob.glob(file_location + "/*PNG"))
    bar = Bar("Extracting text from PNGs", max=len(pngs))
    pool = get_context("spawn").Pool()
    for png in pngs:
        # use same processJPG for PNGs and all other image formats
        pool.apply_async(processJPG, [png], callback=updateBar)
    pool.close()
    pool.join()
    bar.finish()

    # now tiff
    tiffs = glob.glob(file_location + "/*tif*")
    tiffs.extend(glob.glob(file_location + "/*TIF*"))
    bar = Bar("Extracting text from TIFFs", max=len(pngs))
    pool = get_context("spawn").Pool()
    for tiff in tiffs:
        # use same processJPG for TIFFS and all other image formats
        pool.apply_async(processJPG, [tiff], callback=updateBar)
    pool.close()
    pool.join()
    bar.finish()
    """
    # now docx
    docxs = glob.glob(file_location + "/*docx")
    bar = Bar("Extracting text from DOCX", max=len(docxs))
    pool = get_context("spawn").Pool()
    for docx_file in docxs:
        pool.apply_async(processDOCX, [docx_file], callback=updateBar)
    pool.close()
    pool.join()
    bar.finish()
    
    # now doc
    docs = glob.glob(file_location + "/*doc")
    docs = [os.path.basename(x) for x in docs]
    bar = Bar("Extracting text from DOC", max=len(docs))
    pool = get_context("spawn").Pool()
    for doc_file in docs[0:4]:
        #pool.apply_async(processDOC, [doc_file], callback=updateBar)
        pool.apply_async(get_doc_text, [doc_file], callback=updateBar)
    pool.close()
    pool.join()
    bar.finish()
    
    # now rtf
    rtfs = glob.glob(file_location + "/*rtf")
    bar = Bar("Extracting text from RTF", max=len(rtfs))
    pool = get_context("spawn").Pool()
    for rtf in rtfs:
        pool.apply_async(processRTF, [rtf], callback=updateBar)
    pool.close()
    pool.join()
    bar.finish()
    
    # now xls
    xlss = glob.glob(file_location + "/*xls")
    bar = Bar("Extracting text from XLS", max=len(xlss))
    pool = get_context("spawn").Pool()
    for xls in xlss:
        pool.apply_async(processXLS, [xls], callback=updateBar)
    pool.close()
    pool.join()
    bar.finish()
    """


# References
# https://stackoverflow.com/questions/25228106/how-to-extract-text-from-an-existing-docx-file-using-python-docx
# https://stackoverflow.com/questions/36001482/read-doc-file-with-python