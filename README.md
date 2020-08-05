# excel2062-python
Fill out DA2062 using a formatted excel document

Install python3
Install pip: https://www.liquidweb.com/kb/install-pip-windows/

Use these commands:
pip3 install python-docx
pip3 install openpyxl

Use these naming conventions, or change the code where these values appear:
excel sheet name="property.xlsx"
word doc name="da2062.docx"

Make sure the excel sheet and word doc are in the same folder.
Make sure you have closed any open instance of the out file, if one already exists.
Avoid using spaces in cells with long lines of text, as this creates new lines which
will mess up the document formatting. Simply use underlines instead of spaces.

In command line:
python3 2062.py

This will save over any previous version of the out file, so make sure to
copy and move the out file to a safe place each time.
