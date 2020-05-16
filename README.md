## Python Script development to parse and extract test case info from DOCX file

### Pre-requisite
Python 3.x
PyCharm

### Challenge Overview
1. Take the attached HTML file and scan through all the flow blocks.
2. Pick up all the names from the block which has the image or bust of a man. This name becomes the name of the test case (t-code)
3. Search for that name in the attached word document and when found, check if that has an associated table with it.
4. Pick that table and append it in a new spreadsheet giving it the same name as the two documents.
5. Continue with the same process till all the images are covered from the HTML document (HTML document is attached) and discovered in the Word Document (Word document is attached) and moved into the spreadsheet (Sample spreadsheet is attached). 
6. The spreadsheet should have the first column as the name of the block under t-code.
7. Once this set of document is successful and tested for completion, it will be tested for 5 more sets. These sets will be provided at a later stage.
