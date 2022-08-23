import PyPDF2
import os
import re
import openpyxl
import datetime

excelsheet = openpyxl.load_workbook('Invoice_Information.xlsx')
sheet = excelsheet['Sheet1']

# Extracting data from the multiple pdf files
for file_name in os.listdir('invoices'):
    print(file_name)
    load_pdf = open(r'C:\\Users\\KOLOTSANE\\PycharmProjects\\DataExtraction\\invoices\\' + file_name, 'rb')
    read_pdf = PyPDF2.PdfFileReader(load_pdf, strict=False)
    page_count = read_pdf.getNumPages()
    first_page = read_pdf.getPage(0)
    page_content = first_page.extractText()

    try:

        # print(page_content)
        # Finding the arrays of required data
        total = re.findall(r'(?<!\S)(?:(?:cad|[$]|usd|R|M|P) ?[\d,.]+|[\d.,]+(?:cad|[$]|usd))(?!\S)', page_content)
        total1 = total[len(total)-1]
        print(total1)

        invoice_date = re.findall(r'(?:\d{1,2}[-/th|st|nd|rd\s]*)?(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)?[a-z\s,.]*(?:\d{1,2}[-/th|st|nd|rd)\s,]*)+(?:\d{2,4})+', page_content)
        # sorted(invoice_date, key=lambda x: datetime.datetime.strptime(x, '%d-|/| %m-|/| %Y'))
        for i in invoice_date:
            match = re.search("[0-9]{2}\/[0-9]{2}\/[0-9]{4}", i)
            match1 = re.search("[0-9]{2}\-[0-9]{2}\-[0-9]{4}", i)
            match2 = re.search(r'(?:[\s]?\d{1,2}[-/th|st|nd|rd\s]*)?(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)?[a-z\s,./]*(?:\d{1,2}[-/th|st|nd|rd)\s,]*)?(?:\d{2,4})', i)
            if match:
                date = match.group()
                print(date)
            elif match1:
                date = match1.group()
                print(date)
            elif match2:
                date = match2.group()
                print(date)

        # print(invoice_date)

        email = re.findall(r'([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9._-]+)', page_content)
        email1 = email[0]
        print(email1)

        invoice_no = re.search(r'[A-Z]{2,4}[0-9]{4,12}', page_content).group()
        print(invoice_no)

        # Storing data in excel file

        # last_row_number = sheet.max_row
        # print(last_row_number)
        #
        # sheet.cell(column=1, row=last_row_number + 1).value = invoice_no
        # sheet.cell(column=2, row=last_row_number + 1).value = invoice_date
        # sheet.cell(column=3, row=last_row_number + 1).value = invoice_date
        # sheet.cell(column=4, row=last_row_number + 1).value = total1
        # sheet.cell(column=5, row=last_row_number + 1).value = total1
        #
        # # saving a file
        # excelsheet.save('Invoice_Information.xlsx')
    except:
        print("Process failed")
        print('\n')

