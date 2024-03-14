import gradio as gr
import cv2
from paddleocr import PaddleOCR, draw_ocr
import win32com.client as win32
from fpdf import FPDF
import os
import zipfile
import rarfile
import fitz
import PyPDF2

def pdf2img(file_tbp):
    pdf_document = fitz.open(file_tbp)
    input_dir = os.path.dirname(file_tbp)

    for page_num in range(pdf_document.page_count):
        page = pdf_document[page_num]
        image = page.get_pixmap()
        output_img = os.path.join(input_dir, f'page_{page_num + 1}.png')
        image.save(output_img)

    pdf_document.close()

def typeprocess(file_tobeprocess,pdf_output):
    filename = file_tobeprocess.name
    filetype = filename.endswith

    if filetype == 'doc' or filetype == 'docx' :
        word_app = win32.Dispatch('Word.Application')
        doc = word_app.Document.Open(file_tobeprocess)
        doc.SaveAs(pdf_output,FileFormat = 17)
        doc.Close()
        word_app.Quit()

    elif filetype == 'xlsx':
        excel_app = win32.Dispatch('Excel.Application')
        sheet = excel_app.Workbooks.open(file_tobeprocess)
        sheet.ActiveSheet.ExportAsFixedFormat(0,pdf_output)
        sheet.Close()
        excel_app.Quit()

    elif filetype == 'ppt' or filetype == 'pptx':
        ppt_app = win32.Dispatch('PowerPoint.Application')
        present = ppt_app.Presentation.Open(file_tobeprocess)
        present.SaveAs(pdf_output,32)
        present.Close()
        ppt_app.Quit()

    elif filetype == 'txt':
        pdf = FPDF()
        pdf.add_page()
        with open(file_tobeprocess,'r',encoding='utf-8') as file_:
            text = file_.read()
            pdf.set_font("Arial",size=12)
            pdf.multi_cell(0,10,text)
        pdf.output(pdf_output)


    elif filetype == 'zip':
        with zipfile.ZipFile(zipfile,'r') as zip_ref:
            extractpath = zip_ref.extractall(file_tobeprocess.url)
        for root,dirs,files in os.walk(extractpath):
            for file in files :
                typeprocess(file)

    elif filetype == 'rar':
        with rarfile.RarFile(rarfile,'r') as rar_ref:
            extractpath = rar_ref.extractall(file_tobeprocess.url)
        for root,dirs,files in os.walk(extractpath):
            for file in files:
                typeprocess(file)

    pdf2img(pdf_output)

def img_identify(file_tbp):

    pdf_file_name = 'result.pdf'
    program_dir = os.path.dirname(os.path.abspath(__file__))
    pdf_file_path = os.path.join(program_dir,pdf_file_name)
    pdf_document = PyPDF2.PdfFileWriter()
    with open(pdf_file_path,'wb') as pdf_file:
        pdf_document.write(pdf_file)

    paddleocr = PaddleOCR(lang='zh',show_log=False)
    img = typeprocess(file_tbp,pdf_file)
    result = paddleocr.ocr(img)
    alist = [None] * len(result[0])
    for i in range(len(result[0])):
        print(result[0][i][1][0])
        alist[i] = tuple(result[0][i][1][0])

    return alist

iface = gr.Interface(img_identify,gr.File(),gr.Dataframe(),title='SheetTransfer',live=True)
iface.launch()
