# pdfseparator
PDF Separator for reports - written in Python

The pdfseparator program reads PDF files in an input folder, and collates the pages into a single PDF file containing only the pages with matching text.  

Input: CSV file containing 2 columns: "facility" and "outputfile". The "facility" column is the text to be searched for on the page, and the "outputfile" is the output folder that the file should be placed in (containing the collated pages).

The program looks through each PDF file in the input folder, searching for the text in the "facility" column of the csv file. If there are multiple matches on the same page, it uses the match closest to the top of the page.  Then, it creates a PDF output file for each row in the CSV file that had a match.  Finally, if there were any pages in the input file that did not match any entries in the CSV file, it exports those pages with a "missing-" prefix to the source folder.

Prerequisites:
-ActivePython 2.7
-pdfrw library https://code.google.com/p/pdfrw/
-pdfminer library https://pypi.python.org/pypi/pdfminer/

You will want to change the InitialFolder variable to point to the file location where your input PDF files are typically stored.

Also included is a zip file with the most current py2exe execution.  This was done on a 64-bit Windows 7 PC.
