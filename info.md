I would like you to create for me a python program that processes essay feedback from two documents and puts it into a new form in a different format. 

The source documents:
1. The pdf file has a grade, overview, subgrades for 3 subsections with quotes and comments, and other data below this i do not want to use
2. the docx file has a grading report in a 4 row 3 column table

The combined file should docx and format should be exactly like this: Attractive and common sans serif font, 12 points for text, 10 points for tables, and 14 points for headings. All text is simple paragraph form edge to edge 8.5x11 paper with standard margins. No fancy formatting (this is important). It should look like someone typed it not machine generated. This is the format I want with blanks for data to be added later manually and [brackets] for information I am providing to you, but not to be printed:

[start of document]
Name: _______________ 
Essay: [This is the name of the essay, taken from the file name. It is the text between the underscores. For example "Taylor_ Light pollution_review" would have the essay name "Light Pollution" with key words capitalized. No need to preserve trailing and ending spaces.]
Date: [date format 2/17/2026]
[CR]
[CR]
Evidence and Commentary [heading]: 3/4 [this is the section grade]
Overview text
[CR]
Quote 1 [in Italics]
Feedback 1 [standard text]
[CR]
Quote 2 [in Italics]
Feedback 1 [standard text]
[CR]
[and so on for the three headings, grades, quotes, and feedback]
[that ends the data from the pdf file]
[CR]
[CR]
Table from docx file
[CR]
[CR]
ESSAY [Header]
Essay [This section is a precise text of the section labeled content review goes. You must keep it exact, because it is the actual written essay. Keep mispellings, errors, etc. without modification.
[end of document]

The program should be designed to batch load with a button multiple pdfs at once, adding the filenames to a listbox. A second button allows the user to batch load multiple docx files at once and add the filenames to a second listbox. A preview window shows any filename clicked on. There is a batch process button that combines the data into a new file with the file format explained below, saves them each to a new folder on the desktop named "report_data" where the date is in the format "17_Feb_2025" (append folder with number rather than overwriting), and puts the new filenames in a third textbox. These files also show a preview in the same central window when clicked. The filename includes the last name of the student (taken from the first word of the source document before the underscore. For instance, if the source document is "Taylor_ Light pollution_review", the name is Taylor and the filename is "Taylor_report.docx"



