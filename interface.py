import gradio as gr
import cv2
from paddleocr import PaddleOCR, draw_ocr
import win32com.client as win32
import struct
import os


def file_convert(file):

    '''
    # 处理逻辑
    # 读取图像
    # 使用默认模型路径
    paddleocr = PaddleOCR(lang='ch', show_log=False)
    img = cv2.imread(file)  # 打开需要识别的图片
    result = paddleocr.ocr(img)
    alist = [None] * len(result[0])  # 创建一个与result[0]相同长度的空列表
    for i in range(len(result[0])):
        print(result[0][i][1][0])  # 输出识别结果
        alist[i] = tuple(result[0][i][1][0]) # 将识别结果存储到alist中

    return alist
'''

    return file


iface = gr.Interface(file_convert, gr.File(), gr.Dataframe(), title="表格转换器", live=True,)
iface.launch()
