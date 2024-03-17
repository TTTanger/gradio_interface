import gradio as gr
import cv2
from paddleocr import PaddleOCR, draw_ocr
import win32com.client as win32
import struct
import os
import pandas as pd


def file_convert(file):

    # 处理逻辑
    # 读取图像
    # 使用默认模型路径
    paddleocr = PaddleOCR(lang='ch', show_log=False)
    img = cv2.imread(file)  # 打开需要识别的图片
    result = paddleocr.ocr(img)
    alist = []
    for i in range(len(result[0])):
        alist.append(result[0][i][1][0])    # 将识别结果存储到alist中
    print(alist)   # 输出识别结果
    '''
    for line in result:
        print(line)
    '''
    # 将结果转换为DataFrame
    df = pd.DataFrame({'识别结果': alist})

    return df


iface = gr.Interface(file_convert, gr.File(), gr.Dataframe(), title="表格转换器", live=True,)
iface.launch()
