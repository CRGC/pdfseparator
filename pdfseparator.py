#!/usr/bin/env python
import sys
import os
from os import listdir
from os.path import isfile, join
import tempfile
import re
import pdfrw
import csv
import datetime
from pdfrw import PdfReader, PdfWriter
import operator

from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.pdfdevice import PDFDevice, TagExtractor
from pdfminer.pdfpage import PDFPage
from pdfminer.converter import XMLConverter, HTMLConverter, TextConverter
from pdfminer.cmapdb import CMapDB
from pdfminer.layout import LAParams
from pdfminer.image import ImageWriter
from Tkinter import Tk
from tkFileDialog import askdirectory, askopenfilename

from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from cStringIO import StringIO

def convert_pdf_to_txt(path, PageNumber):
    rsrcmgr = PDFResourceManager()
    retstr = StringIO()
    codec = 'utf-8'
    laparams = LAParams()
    device = TextConverter(rsrcmgr, retstr, codec=codec, laparams=laparams)
    fp = file(path, 'rb')
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    password = ""
    maxpages = 0
    caching = True
    pagenos=[PageNumber]
    for page in PDFPage.get_pages(fp, pagenos, maxpages=maxpages,password=password,caching=caching, check_extractable=True):
        interpreter.process_page(page)

    text = retstr.getvalue()

    fp.close()
    device.close()
    retstr.close()
    return text


def list_files(path):
    # returns a list of names (with extension, without full path) of all PDF / pdf files
    # in folder path
    files = []
    for name in os.listdir(path):
        if os.path.isfile(os.path.join(path, name)):
            if (".pdf" in name) or (".PDF" in name):
               files.append(name)
    return files

rxcountpages = re.compile(r"/Type\s*/Page([^s]|$)", re.MULTILINE|re.DOTALL)

def count_pages(filename):
    data = file(filename,"rb").read()
    return len(rxcountpages.findall(data))

# Initial Folder - change to match your implementation.
InitialFolder = "C:\\temp"


# main
def main(argv):
    print "\nPDF Report Separator\nThis program takes a folder full of PDF files, \nand separates them by facility.\n"

    now = datetime.datetime.now()
    YearMonth = "%04d-%02d" % (now.year,(now.month-1))
    #----------------------------------------------------------------------------------------------
    # This is the place to prompt for a file name preface (followed by a dash)
    #----------------------------------------------------------------------------------------------
    FilePreface = raw_input('Provide a file name preface. <ENTER> will use "' + YearMonth + '-Preliminary": ')
    if FilePreface == '':
       FilePreface = YearMonth + "-Preliminary"

    FilePreface = FilePreface + '-'

    print "\nPlease choose the folder that contains ONLY THE PDF FILES YOU WISH TO IMPORT. \nNO OTHER FILES."
    #Prompt the user for a folder that contains the input files
    Tk().withdraw()  # Hide the window that appeared
    ImportFolder = askdirectory(title='PDF Input Files',initialdir=InitialFolder)
    #print ImportFolder

    print "Please choose the file that contains a list of the facility names, \nwhich match those in the PDF document"
    InputFacilitiesList = askopenfilename(title='Facilities List')

    #Look through the input directory, and create a list of PDF files we will be using for import.
    InputFilenameList = list_files(ImportFolder)

    #Delete the "Missing-" files from the folder. The program created the files in the first place and we do not want to
    #  process them along with the real PDF files
    for InputFile in InputFilenameList:
        if InputFile.find("Missing-") != -1:
           #print "Found: " + InputFile
           os.remove(ImportFolder + "\\" + InputFile)

    #Re-issue the command since we may have deleted files from the directory with the PDF files in it
    InputFilenameList = list_files(ImportFolder)

    #Create a dictionary of missing pages.
    MissingPagesDict = {}
    for item in InputFilenameList:
        MissingPagesDict[item] = []

    #Read in the list of facilities from a text file
    ListOfFacilities = [] # List of Facilities. Used when searching for the facility in the PDF. Somewhat of a duplicate.
    FacilitiesDict = {} #Dictionary of files / facilities. The key is File~~Facility. The value is a list of
                        #pages that correspond to that facility.
    OutputFiles = {} # Dictionary of Facility to Output File.
    csv_file = csv.DictReader(open(InputFacilitiesList, 'rb'), delimiter=',')
    for line in csv_file:
        for InputFile in InputFilenameList:
           #Create one entry per facility / file.
           #The dictionary keys will look like this: File~~Facility
           DictKey = InputFile + '~~' + line['facility'].strip()
           FacilitiesDict[DictKey] = []
           OutputFiles[line['facility'].strip()] = line['outputfile'].strip()
           ListOfFacilities.append(line['facility'].strip())



    pagenos = set()
    MAXPAGES = 0
    codec = 'utf-8'
    pageno = 1
    scale = 1
    caching = True
    debug = 0
    PDFDocument.debug = debug
    PDFParser.debug = debug
    CMapDB.debug = debug
    PDFResourceManager.debug = debug
    PDFPageInterpreter.debug = debug
    PDFDevice.debug = debug
    #
    rsrcmgr = PDFResourceManager(caching=caching)
    outfp = tempfile.SpooledTemporaryFile(max_size=10000000,mode='rw')
    device = TextConverter(rsrcmgr, outfp, codec=codec, laparams=None, imagewriter=None)

    #Read the input files,

    for fname in InputFilenameList:
        PDFFullPath = ImportFolder + '/' + fname
        fp = file(PDFFullPath, 'rb')
        PageCount = len(rxcountpages.findall(fp.read())) # Count the # of pages in the input file
        #-------------------------------------------------
        #PageCount = PageCount - 1  #Compensating for a zero based array?
        #-------------------------------------------------
        print 'Procesing %s : %s pages total' % (fname, str(PageCount))

        # Iterate through each page in the PDF
        for PageIter in range(0, PageCount, 1):
            print "Processing page " + str(PageIter+1)
            #Make a list of all the pages that could be "missing" so we can remove the "found" pages from it.
            MissingPagesDict[fname].append(PageIter)

            #Grab the text from the page.
            CurrentPageText = convert_pdf_to_txt(PDFFullPath, PageIter)

            #Hold the list of facilities found in each page
            FacilitiesFoundInPage = {}
            #Key is the facility name, value is the first location in the text.
            FirstFacilityFound = ''

            #Look for each facility name in the page.
            for FacilityIter in ListOfFacilities:
                LocationFound = CurrentPageText.find(FacilityIter)
                if LocationFound >= 0: #We found this facility:
                   FacilitiesFoundInPage[FacilityIter] = LocationFound
                   FirstFacilityFound = FacilityIter #Used when only one facility name was found on the page ... gets overwritten in the next statement if there was more than one facility on the page.
                   #print "We found facility " + FacilityIter + "\n"
            if len(FacilitiesFoundInPage) > 1:
               #Grab the facility with the earliest mention on the page - which is likely the header.
               FirstFacilityFound = min(FacilitiesFoundInPage.iteritems(), key=operator.itemgetter(1))[0]


            if FirstFacilityFound <> '': #If we have found a facility, notate it.
                #Make our key to notate the file and page number where we located this facility
                FileFacilityKey = fname + '~~' + FirstFacilityFound
                #Add this page to the FacilitiesDict (which is really File~~Facility dict)
                FacilitiesDict[FileFacilityKey].append(PageIter)
                #
                MissingPagesDict[fname].remove(PageIter)

        fp.close()

    device.close()
    outfp.close()
    #Write each facility's files to their output file.

    #Iterate through the facilities.
    for facility in OutputFiles:
        outdata = PdfWriter() #Initialize our output writing for a new file.
        outfile = OutputFiles[facility] + FilePreface + facility + ".pdf" #Match the output file to the full filename specified in the CSV.
        for fname in InputFilenameList: # Iterate through all the files in the input folder
            FullPathFile = ImportFolder + '/' + fname
            InputPages = PdfReader(FullPathFile).pages #InputPages is a PdfReader link to the input file
            for pagenum in FacilitiesDict[fname + '~~' + facility]: #Go through each page for that facility and add it to the output file.
                    outdata.addpage(InputPages[pagenum])

        #We are done with this facility. Write the file.
        outdata.write(outfile)

    #Write the pages (for which we did not have a facility listed in the CSV) to the Missing- PDFs (in the input folder).

    for InputFile in InputFilenameList:
        print "Processing missing for " + InputFile
        outdata = PdfWriter() #Initialize our output writing for a new file.
        outfile = ImportFolder + '/' + "Missing-" + InputFile #Match the output file to the full filename specified in the CSV.
        FullPathFile = ImportFolder + '/' + InputFile
        InputPages = PdfReader(FullPathFile).pages
        if len(MissingPagesDict[InputFile]) > 0:
            for pagenum in MissingPagesDict[InputFile]:
              print "Adding Missing pagenum: " + str(pagenum)
              outdata.addpage(InputPages[pagenum])
            #We are done with this input file. Write the missing facilities file.
            outdata.write(outfile)

    return

if __name__ == '__main__': sys.exit(main(sys.argv))
