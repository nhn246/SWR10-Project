'''
Authors: chatGPT, Scot Rosenquist, Nate Nattin

functions in this script were translated from AutoIT to python by chatGPT

from AutoIT script: "S:\PERMITTING\ENG\SWR 10\SWR 10\(4) make PDF.au3"

date: March 6, 2023
'''


import os
import subprocess
import time
import win32com.client
import shutil
import random
import psutil



def getcol(e, workbookname, coltitle):
    found = False
    lastcol = e.Workbooks(workbookname.Name).Worksheets(1).Cells.SpecialCells(11).Column
    # print (lastcol)
    for c in range(1, lastcol + 1):
        # print (c)
        if e.Workbooks(workbookname.Name).Worksheets(1).Cells(1, c).Text == coltitle:
            found = True
            break
    if not found:
        print(f"failed to find column {coltitle}")
        return None
    return c



def print_word_document_to_PDF_creator(worddocumentfilepath, PDFfiledirectory, PDFdocumenttitle):
    # blockinput(1)
    subprocess.run(["clip.exe"], input=b"", check=True) # clears clipboard
    # opt("trayicondebug", 1)
    # opt("expandenvstrings", 1)
    # blockinput(1)
    # mousemove(500, 500, 1)
    # tooltip("Mouse is disabled while PDF is being created.  Please wait.")
    w = None
    try:
        w = win32com.client.DispatchEx("Word.Application")
    except:
        w = win32com.client.Dispatch("Word.Application")
    w.Visible = 1
    w.Documents.Open(worddocumentfilepath)
    temp = worddocumentfilepath.split("\\")
    worddocfilename = temp[-1]
    print ("Print Status", "Printing document to PDF Creator.  Please wait...", 60)
    input_file = open(os.path.join(os.path.dirname(os.path.realpath(__file__)), "PDFCreator settings.reg"), 'r')
    tempfilename = str(random.randint(0,100000)) + ".reg"
    tempfilepath = os.path.join(os.path.dirname(os.path.realpath(__file__)), tempfilename)
    output_file = open(tempfilepath, 'w')
    while True:
        line = input_file.readline()
        if not line:
            break
        if "AutosaveFilename" in line:
            output_file.write('AutosaveFilename="{}"\n'.format(PDFdocumenttitle))
        elif "AutosaveDirectory" in line:
            PDFfiledirectory = PDFfiledirectory.replace("\\", "\\\\")
            output_file.write('AutosaveDirectory="{}"\n'.format(PDFfiledirectory))
        else:
            output_file.write(line)
    input_file.close()
    output_file.close()
    # filesetattrib(@scriptdir & "\PDFCreator settings.reg", "-R")
    # filedelete(@scriptdir & "\PDFCreator settings.reg")
    # filecopy($tempfilepath, @scriptdir & "\PDFCreator settings.reg")
    subprocess.run(["taskkill.exe", "/f", "/im", "PDFCreator.exe"], check=True)
    subprocess.run(["reg", "import", tempfilename], cwd=os.path.dirname(os.path.realpath(__file__)))
    os.remove(tempfilepath)
    originalprinter = w.ActivePrinter
    w.ActivePrinter = "PDFCreator"
    w.ActiveDocument.PrintOut()
    w.ActivePrinter = originalprinter
    print ("", "", 60)
    while not (os.path.exists(PDFfiledirectory + "\\" + PDFdocumenttitle + ".pdf") or os.path.exists(PDFfiledirectory + "\\" + PDFdocumenttitle)):
        time.sleep(1)
    c = 1
    while "PDFCreator.exe" in [p.name() for p in psutil.process_iter()]:
        time.sleep(1)
        c += 1
        if c > 20:
            break
    # blockinput(0)
    # tooltip("")
    subprocess.run(["taskkill.exe", "/f", "/im", "PDFCreator.exe"], check=True)
    subprocess.run(["reg", "import", "PDFCreator settings.reg"], cwd=os.path.dirname(os.path.realpath(__file__)), creationflags=subprocess.CREATE_NO_WINDOW)
    filepath = os.path.join(PDFfiledirectory, PDFdocumenttitle)
    filepath = filepath.replace("\\", "\\\\")
    filepath = filepath.replace("\\\\", "\\")
    subprocess.run(["clip.exe"], input=filepath.encode('utf-8'), check=True)
    w.ActiveDocument.Close(0)
    # w.Quit()




def main():

    host_directory = os.path.dirname(os.path.abspath(__file__)) + "\\"
    excel = win32com.client.Dispatch("Excel.Application")
    workbook = excel.ActiveWorkbook
    # first_row = excel.Selection.Row
    # last_row = excel.ActiveSheet.Cells.SpecialCells(11).Row

    for r in range(excel.Selection.Row, excel.Selection.Row+len(excel.Selection.Rows)):

        excel.Workbooks(workbook.Name).Worksheets(1).Cells(r, 1).Activate

        attn = excel.Workbooks(workbook.Name).Worksheets(1).Cells(r, getcol(excel, workbook, "ATTN")).Text
        api = excel.Workbooks(workbook.Name).Worksheets(1).Cells(r, getcol(excel, workbook, "API#")).Text
        opname = excel.Workbooks(workbook.Name).Worksheets(1).Cells(r, getcol(excel, workbook, "Operator")).Text
        opnumber = excel.Workbooks(workbook.Name).Worksheets(1).Cells(r, getcol(excel, workbook, "operator no.")).Text
        docket = excel.Workbooks(workbook.Name).Worksheets(1).Cells(r, getcol(excel, workbook, "Docket #")).Text
        template = excel.Workbooks(workbook.Name).Worksheets(1).Cells(r, getcol(excel, workbook, "type of letter")).Text
        district = excel.Workbooks(workbook.Name).Worksheets(1).Cells(r, getcol(excel, workbook, "District")).Text

        doc_file_path = excel.Workbooks(workbook.Name).Worksheets(1).Cells(r, getcol(excel, workbook, "disposition letter file name")).Text
        doc_file_path = os.path.join(r'S:\PERMITTING\ENG\SWR 10\SWR 10\letters', doc_file_path)
        temp_doc_path = os.path.expandvars("%userprofile%\\Desktop\\PDF Creator\\") + api + ".doc"
        temp_pdf_path = os.path.expandvars("%userprofile%\\Desktop\\PDF Creator\\") + api + ".pdf"
        try:
            os.mkdir(os.path.join(os.environ['userprofile'], 'desktop', 'PDF Creator'))
        except:
            pass
        shutil.copy2(doc_file_path, temp_doc_path)

        word_obj = win32com.client.Dispatch("Word.Application")
        word_obj.Visible = True
        doc_obj = word_obj.Documents.Open(temp_doc_path)

        doc_obj.ExportAsFixedFormat(temp_pdf_path, 17)

        temp = doc_file_path.split("\\")
        word_doc_file_name = temp[-1]

        pdf_file_directory = "S:\\PERMITTING\\ENG\\SWR 10 approval letters\\district " + district
        pdf_document_title = os.path.splitext(word_doc_file_name)[0] + ".pdf"

        while not os.path.exists(os.path.expandvars("%userprofile%\\Desktop\\PDF Creator\\") + api + ".pdf"):
            time.sleep(0.1)

        while "PDFCreator.exe" in (p.name() for p in psutil.process_iter()):
            time.sleep(0.1)

        time.sleep(2)
        shutil.copy2(temp_pdf_path, os.path.join(pdf_file_directory, pdf_document_title))
        shutil.copy2(temp_pdf_path, "S:\\PERMITTING\\ENG\\SWR 10 approval letters\\letters to email\\" + attn + " " + pdf_document_title)

        doc_obj.Close(0)



if __name__ == "__main__":
    main()
