import PyPDF2
import os
import re
import openpyxl

excelsheet = openpyxl.load_workbook('Invoice_Information.xlsx')
sheet = excelsheet['Sheet1']

for file_name in os.listdir('invoices'):
    print(file_name)
    load_pdf = open(r'C:\\Users\\KOLOTSANE\\PycharmProjects\\DataExtraction\\invoices\\' + file_name, 'rb')
    read_pdf = PyPDF2.PdfFileReader(load_pdf, strict=False)
    page_count = read_pdf.getNumPages()
    first_page = read_pdf.getPage(0)
    page_content = first_page.extractText()
    # print(page_content)

    try:
        # company_name = re.search(r'Company(.*)', page_content).group().split("Company")[1]
        # print(company_name)
        # address = invoice_no = re.search(r'Address(.*)', page_content).group().split("Address")[1]
        # print(address)
        # email = re.search(r'([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9._-]+)', page_content).group()
        # print(email)
        invoice_no = re.search(r'Invoice Number(.*)', page_content).group().split("Invoice Number")[1]
        print(invoice_no)
        invoice_date = re.search(r'Invoice Date (.*)', page_content).group().split("Invoice Date")[1]
        print(invoice_date)
        due_date = re.search(r'Due Date (.*)', page_content).group().split("Due Date")[1]
        print(due_date)
        sub_total = re.search(r'Sub Total (.*)', page_content).group().split("Sub Total")[1]
        print(sub_total)
        total = re.search(r'Total Due (.*)', page_content).group().split("Total Due")[1]
        print(total)

        last_row_number = sheet.max_row
        print(last_row_number)

        sheet.cell(column=1, row=last_row_number + 1).value = invoice_no
        sheet.cell(column=2, row=last_row_number + 1).value = invoice_date
        sheet.cell(column=3, row=last_row_number + 1).value = due_date
        sheet.cell(column=4, row=last_row_number + 1).value = sub_total
        sheet.cell(column=5, row=last_row_number + 1).value = total

        # saving a file
        excelsheet.save('Invoice_Information.xlsx')
    except:
        print("Process failed")

