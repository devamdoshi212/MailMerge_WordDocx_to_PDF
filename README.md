![alt text](instruction_with_screenshots\1690620623Deer16b.svg)

# MailMerge_WordDocx_to_PDF

After Mail Merge , Document only create in .docx format in MSWord . So this can help you to covert that MailMerge Document into .pdf format

# Step 1

Open MS Word and right click on Menubar and open customize the Ribbon..

![alt text](instruction_with_screenshots/step1.png)

# Step 2

Checked Developer box and Save Settings

![alt text](instruction_with_screenshots/step2.png)

# Step 3

Open Developer Menu in MS Word and Open Visual Basic

![alt text](instruction_with_screenshots/step3.png)

# Step 4

Open Insert and Open Module

![alt text](instruction_with_screenshots/step4.png)

# Step 5

Add code.vb in this module (just copy and paste)

![alt text](instruction_with_screenshots/step5.png)

# Step 6

Just save module settings and open new MS word file

![alt text](instruction_with_screenshots/step6.png)
![alt text](instruction_with_screenshots/step6_2.png)

# Step 7

Create a excel file for MailMerge, add column as per your needs but add complusory 4 column in that excel file (DocFolderPath , DocFileName , PdfFolderPath , PdfFileName).
In above 4 columns , you should add DocFolderPath where your .docx file will be saved ,also should add PdfFolderPath where your .pdf file will be saved.
In DocFileName and PdfFileName , set name as per your needs.
Save the excel file(demo.xlsx)

![alt text](instruction_with_screenshots/step7.png)

# Step 8

Open Mailings and select Select Recipients and choose Use an Existing List...

![alt text](instruction_with_screenshots/step8.png)
![alt text](instruction_with_screenshots/step9.png)

# Step 9

After that , insert datafield as per your need.

![alt text](instruction_with_screenshots/step10.png)

Use datafield like this...
e.g.

![alt text](instruction_with_screenshots/step11.png)

# Step 10

Open Developer Tab and Select Macros

![alt text](instruction_with_screenshots/step12.png)

# Step 11

Hit Run Button !!

![alt text](instruction_with_screenshots/step13.png)

And finally that file created in your desire folder

![alt text](instruction_with_screenshots/step14.png)
