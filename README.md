# DOCX-Invoice-Generator
This Python script processes .docx files in a folder named articles, counts the number of words in each file (including text inside tables), calculates the cost based on a given rate per word, and generates a PDF invoice with all the details.
________________________________________
Folder Structure
bash
Copy code
project-directory/
│
├── articles/               # Place all .docx files here
│   ├── file1.docx
│   ├── file2.docx
│   └── ...
│
├── invoice_generator.py    # The main Python script
├── README.md               # This readme file
└── invoice.pdf             # Generated output (user-defined name)
________________________________________
Requirements
Before running the script, install the required libraries using pip:
nginx
Copy code
pip install python-docx fpdf
________________________________________
How to Use
1.	Make sure you have a folder named articles in the same directory as the script.
2.	Place all your .docx files inside the articles folder.
3.	Run the script:
nginx
Copy code
python invoice_generator.py
4.	The script will ask for:
o	The rate per word (e.g., 1.4)
o	The name of the output PDF file
________________________________________
Output
The generated PDF invoice includes:
•	Serial number
•	File name (without the .docx extension)
•	Word count
•	Rate per word
•	Total price per file
At the end of the table:
•	Total words across all files
•	Total cost
________________________________________
Word Count Logic
The word count includes:
•	Text from all paragraphs
•	Text from all table cells
________________________________________
Example
Input files:
•	file1.docx: 500 words
•	file2.docx: 800 words
Rate: 1.4
Output PDF will show:
yaml
Copy code
Sr. | Description | Qty | Rate | Price
1   | file1       | 500 | 1.4  | 700.0
2   | file2       | 800 | 1.4  | 1120.0

Total Words: 1300
Total Rs: 1820

