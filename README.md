# test_open_xml
#OpenOfficeWpfApp
A desktop application based on .NET platform.
Program opens word document 'test.docx' and operates with XML data inside this file. 
All changes are applied in the same file, so modified 'test.docx' is a result of executing this program.
What it does:
 - adds a row "{FirstName} {LastName}, {City}" after first paragraph, where params are my personal info;
 - all blue symbols are changed to green color;
 - red words are underlined;
 - on sheet #2 adds table with the following structure - http://1.bp.blogspot.com/-5RDua4sju1M/TWNzHKtbVAI/AAAAAAAAADI/Dacxd9hVqtw/s1600/timeTable.gif

It uses Open XML SDK 
test.docx is in "data" folder
Istall Open XML SDK 2.5 for Microsoft Office first: https://www.microsoft.com/en-us/download/details.aspx?id=30425
