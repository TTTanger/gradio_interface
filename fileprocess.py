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
    input_dir = os.path.dirname(file_tbp)   # 获取输入文件地址

    for page_num in range(pdf_document.page_count):
        page = pdf_document[page_num]
        image = page.get_pixmap()
        output_img = os.path.join(input_dir, f'page_{page_num + 1}.png')    #输出图片集到输入文件相同目录，避免地址混乱
        image.save(output_img)

    pdf_document.close()


def typeprocess(file_tbp, pdf_output):
    filename = file_tbp.name
    filetype = filename.endswith    # 获取文件后缀

    if filetype == 'doc' or filetype == 'docx':
        word_app = win32.Dispatch('Word.Application')   # 匹配本地的word程序
        doc = word_app.Document.Open(file_tbp)
        doc.SaveAs(pdf_output, FileFormat=17)      # 17表示pdf格式
        doc.Close()
        word_app.Quit()

    elif filetype == 'xlsx':
        excel_app = win32.Dispatch('Excel.Application')
        sheet = excel_app.Workbooks.open(file_tbp)
        sheet.ActiveSheet.ExportAsFixedFormat(0, pdf_output)     # 0表示pdf格式
        sheet.Close()
        excel_app.Quit()

    elif filetype == 'ppt' or filetype == 'pptx':
        ppt_app = win32.Dispatch('PowerPoint.Application')
        present = ppt_app.Presentation.Open(file_tbp)
        present.SaveAs(pdf_output, 32)       # 32表示pdf格式
        present.Close()
        ppt_app.Quit()

    elif filetype == 'txt':
        pdf = FPDF()
        pdf.add_page()
        with open(file_tbp, 'r', encoding='utf-8') as file_:
            text = file_.read()
            pdf.set_font("Arial", size=12)
            pdf.multi_cell(0, 10, text)
        pdf.output(pdf_output)


    elif filetype == 'zip':
        with zipfile.ZipFile(zipfile,'r') as zip_ref:
            extractpath = zip_ref.extractall(file_tbp.url)      # 解压到输入文件相同目录
        for root, dirs, files in os.walk(extractpath):    # 遍历目录下的内容
            for file in files:
                typeprocess(file)

    elif filetype == 'rar':
        with rarfile.RarFile(rarfile,'r') as rar_ref:
            extractpath = rar_ref.extractall(file_tbp.url)
        for root, dirs, files in os.walk(extractpath):
            for file in files:
                typeprocess(file)

    pdf2img(pdf_output)     # 转换pdf后继续转换为img

def img_identify(file_tbp):

    pdf_file_name = 'result.pdf'    # 创建一个内容为控的pdf文件
    program_dir = os.path.dirname(os.path.abspath(__file__))
    pdf_file_path = os.path.join(program_dir,pdf_file_name)     # 与程序在同一个目录下
    pdf_document = PyPDF2.PdfWriter()
    with open(pdf_file_path, 'wb') as pdf_file:
        pdf_document.write(pdf_file)

    paddleocr = PaddleOCR(lang='ch', show_log=False)
    img = typeprocess(file_tbp, pdf_file)
    result = paddleocr.ocr(img)
    alist = [None] * len(result[0])
    for i in range(len(result[0])):
        print(result[0][i][1][0])
        alist[i] = tuple(result[0][i][1][0])    # 新建一个列表储存result内容防止Dataframe输出识别参数

    return alist


iface = gr.Interface(img_identify, gr.File(), gr.Dataframe(), title='SheetTransfer', live=True)
iface.launch()
