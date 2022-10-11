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
# also must run 'sudo apt-get install antiword' to install antiword for textract


# Used in main and updateBar, but not child processes
bar = None

# Takes a path to a single pdf, saves contents as
# "foo.pdf.txt" in the current directory
def processPDF(pdf):
    basename = os.path.basename(pdf)
    try:
        with pdfplumber.open(pdf) as pdfdoc:
            with open(basename + ".txt", "w+") as txt:
                for page in pdfdoc.pages:
                    text = page.extract_text()
                    #print("text",text)
                    if( text != None ):
                        transl_text = GoogleTranslator(source="ru", target="en").translate(text=text)
                        txt.write(transl_text + "\n")
        return
    except:
        return ("! - Error parsing '%s', skipping..." % basename)


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
    
    # pdfs
    pdfs = glob.glob("/mnt/c/Users/natra/Documents/Research/Gary/sample_russian_docs/*pdf")
    bar = Bar("Extracting text from PDFs", max=len(pdfs))
    pool = get_context("spawn").Pool()
    for pdf in pdfs:
        pool.apply_async(processPDF, [pdf], callback=updateBar)
    pool.close()
    pool.join()
    bar.finish()

    # now docx
    docxs = glob.glob("/mnt/c/Users/natra/Documents/Research/Gary/sample_russian_docs/*docx")
    bar = Bar("Extracting text from DOCX", max=len(docxs))
    pool = get_context("spawn").Pool()
    for docx_file in docxs:
        pool.apply_async(processDOCX, [docx_file], callback=updateBar)
    pool.close()
    pool.join()
    bar.finish()

    # now doc
    docs = glob.glob("/mnt/c/Users/natra/Documents/Research/Gary/sample_russian_docs/*doc")
    bar = Bar("Extracting text from DOC", max=len(docs))
    pool = get_context("spawn").Pool()
    for doc_file in docs:
        pool.apply_async(processDOC, [doc_file], callback=updateBar)
    pool.close()
    pool.join()
    bar.finish()

    # now rtf
    rtfs = glob.glob("/mnt/c/Users/natra/Documents/Research/Gary/sample_russian_docs/*rtf")
    bar = Bar("Extracting text from RTF", max=len(rtfs))
    pool = get_context("spawn").Pool()
    for rtf in rtfs:
        pool.apply_async(processRTF, [rtf], callback=updateBar)
    pool.close()
    pool.join()
    bar.finish()
    
    # now xls
    xlss = glob.glob("/mnt/c/Users/natra/Documents/Research/Gary/sample_russian_docs/*xls")
    bar = Bar("Extracting text from XLS", max=len(xlss))
    pool = get_context("spawn").Pool()
    for xls in xlss:
        pool.apply_async(processXLS, [xls], callback=updateBar)
    pool.close()
    pool.join()
    bar.finish()




# References
# https://stackoverflow.com/questions/25228106/how-to-extract-text-from-an-existing-docx-file-using-python-docx
# https://stackoverflow.com/questions/36001482/read-doc-file-with-python