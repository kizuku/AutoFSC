# AutoFSC

 -- Tool to automate generation of FSC documentation, written in AutoIT.

Created with AutoIT by Zachary Zhao (zachary.m.zhao@gmail.com)

In order to use the tool, open AutoFSC.exe and follow the instructions below.

===BEFORE USE ===============================
Before using AutoFSC, you as the user must do a few things:
	1. Remove the security level on all relevant PDFs.
	2. Fill out the protocol as necessary and convert it from a Word doc to a PDF.
	3. Extract the header image from the protocol in JPG form.
	4. Ensure you have the proper FS, in PDF form.

===DURING USE ===============================
The extraction step requires the FS, in PDF form, and the header image, in JPG form. The header image filename must end in "header", case-insensitive.

To run the extraction step, choose the necessary files, change the page size if needed, and enter the page numbers to extract. After 
everything is set, press the "Begin extraction" button.

The combination step requires the Protocol in PDF form, already filled out, and the FS test pages, in PDF form, from the extraction step. 
The protocol filename must end in "protocol", and the test pages filename must end in "test-pages", case-insensitive.

To run the combination step, choose the necessary files, and press the "Begin combination" button.

When printing the final document, make sure to shrink oversized pages.

