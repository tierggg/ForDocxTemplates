from tkinter.filedialog import askopenfile
from PyPDF2 import PdfFileReader
from shutil import copy2
for k in range(1, 993):
    pdf_document = 'D:/1/'+str(k)+'.pdf'
    with open(pdf_document, "rb") as filehandle:
        pdf = PdfFileReader(filehandle)

        info = pdf.getDocumentInfo()
        pages = pdf.getNumPages()
        for i in range(pages):
            page = pdf.getPage(i)
            #print(page.extractText())
            sss = page.extractText().split()
            print(sss[0], k)
            first = 'D:\\1\\'+str(k)+'.pdf'
            second = 'D:\\1\\OUT\\'+str(sss[0])+'_'+str(k)+'.pdf'
            copy2(first, second)