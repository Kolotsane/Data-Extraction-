import PyPDF2
import os

dir_path = 'C://Users//KOLOTSANE//PycharmProjects//DataExtraction//invoices//'
os.chdir(dir_path)
file_list = os.listdir()

for file_Name in file_list:

    f = open(dir_path + file_Name, 'rb')
    reader = PyPDF2.PdfFileReader(f)
    file_contents = reader.getPage(0).extract_text().split("\n")

    invoice_no = ""
    date = ""

    for i in range(len(file_contents)):
        if file_contents[i].find('INVOICE #') != -1:
            invoice_no = file_contents[i].split(" ")[1]
        if file_contents[i].find('DATE') != -1:
            date = file_contents[i].split(":")[1]
        if file_contents[i].find('Invoice Number') != -1:
            invoice_no = file_contents[i].split(" ")[1]
        if file_contents[i].find('Invoice Date') != -1:
            date = file_contents[i].split(" ")[1]

    print(invoice_no + ',' + date)
    f.close()
