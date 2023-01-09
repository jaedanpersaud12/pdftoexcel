# Import necessary libraries for working with JSON, Excel, OpenAI, and PDF files
import json
import openpyxl
import openai
from PyPDF2 import PdfReader
import os
from decouple import config

# Set the OpenAI organization and API key
openai.organization = "org-sFxfkEzXsJ2IliNjE8FaPZqL"
openai.api_key = config("OPENAI_API")

# Read in the PDF file and extract the text from the first page
reader = PdfReader("example2.pdf")
number_of_pages = len(reader.pages)
page = reader.pages[0]
text = page.extract_text()

# Use OpenAI's Completion API to generate a JSON from the extracted text
answer = openai.Completion.create(
    model="text-davinci-003",
    prompt=f'Please generate a JSON from the following raw text: {text} The JSON should have the following '
           'structure: { "name": "Name", "address": "Address", "invoiceDate": "InvoiceDate", "dueDate": "DueDate", '
           '"customer": { "name": "CustomerName", "email": "CustomerEmail", "address": "CustomerAddress" }, '
           '"description": "Description", "items": [ { "description": "ItemDescription", "hours": "Hours", '
           '"days": "Days", "unitPrice": "UnitPrice", "total": "Total" } ], "subtotal": "Subtotal", "total": "Total", '
           '"balanceDue": "BalanceDue", "paymentTerms": "PaymentTerms" }',
    max_tokens=500,
    temperature=0
)

# Load the generated JSON as a Python object
parsed_json = json.loads(answer["choices"][0]["text"])

# Print the JSON object to the console
print(parsed_json)

# Load the Excel workbook and get the active sheet
workbook = openpyxl.load_workbook('invoice.xlsx')
sheet = workbook.active

# Create a new row in the sheet using data from the JSON object
row = [parsed_json["name"], parsed_json["address"], parsed_json["invoiceDate"], parsed_json["dueDate"],
       parsed_json["customer"]["name"], parsed_json["customer"]["email"], parsed_json["customer"]["address"],
       parsed_json["description"], parsed_json["subtotal"], parsed_json["total"], parsed_json["balanceDue"],
       parsed_json["paymentTerms"]]
for item in parsed_json["items"]:
    current_items = [item["description"], item["hours"], item["days"], item["unitPrice"], item["total"]]
    for i in current_items:
        row.append(i)

# Append the new row to the sheet
sheet.append(row)

# Save the workbook
workbook.save('invoice.xlsx')
