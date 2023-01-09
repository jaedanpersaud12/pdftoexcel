# pdftoexcel
This repository contains a Python script that uses the OpenAI GPT-3 language model to parse data from a PDF file and create a new row in an Excel spreadsheet

#### To use this code, you will need to install the following libraries:

```
pip install openpyxl
pip install openai
pip install PyPDF2
pip install python-decouple
```

This is an early version of a script that reads in a PDF file, extracts the text from the first page, and uses OpenAI's Completion API to generate a JSON object from the extracted text. The JSON object is then used to populate a new row in an Excel sheet. In the future, the script will be able to process multiple documents and intelligently populate relevant spreadsheets based on the document type.

To use the Completion API, you will need to set your OpenAI organization and API key:

```
openai.organization = "YOUR_ORGANIZATION"
openai.api_key = "YOUR_API_KEY"
```

The name of the PDF file and the Excel file can be modified in the code. The PDF file should be in the same directory as the code file. The Excel file should also be in the same directory and should be an existing workbook with at least one sheet.

When the code is run, it will print the JSON object to the console and add the data from the JSON object to a new row in the Excel sheet. The workbook will then be saved.
