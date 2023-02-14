# coding: utf-8
import pandas as pd
import qrcode
import os
import tkinter as tk
from tkinter import filedialog
from PIL import Image, ImageTk
from PIL import ImageFilter
from PIL import ImageEnhance
from PIL import ImageDraw,ImageFont

import win32print
import win32ui
from PIL import Image, ImageWin

# 打印机标识符
REPRE_SYMBOL = 'EasyCoder'

# 以下三个变量决定生成二维码的像素，即承载文字信息的多少
VERSION=8
BOX_SIZE=10
BORDER=2

# 调试程序用的，True则只打印表格中前两行
FORTEST=False
# SCALE=0.3

# 长款比
height_width_ratio=44.9/34.7

# 生成二维码图片的宽度 单位是像素
target_width=680
target_height=int(target_width*height_width_ratio)
# print('target_width',target_width)
# print('target_height',target_height)

# 图片上方宽度
up_margin=100
# 图片上方的文字大小
up_fontsize=70
# 图片底部文字大小
bottom_fontsize=55


# 找到二维码打印机 （用EasyCoder作为标识符的）
def set_up_default_printer():
    printer_name = win32print.GetDefaultPrinter()
    if REPRE_SYMBOL in printer_name:
        print('default printer',printer_name)
        return
    for m in range(10):
        printers = win32print.EnumPrinters(m)
        choose=None
        for i in printers:
            for j in i:
                if REPRE_SYMBOL in str(j):
                    choose = i
                    win32print.SetDefaultPrinter(choose[2])
                    printer_name = win32print.GetDefaultPrinter()
                    print('default printer',printer_name)
                    return

# 传入图片文件的位置，连接打印机打印图片
def print_img(file_name, SCALE):
    printer_name = win32print.GetDefaultPrinter()
    hDC = win32ui.CreateDC ()
    hDC.CreatePrinterDC (printer_name)
    # if file_name is path
    img=file_name
    bmp = Image.open(img)
    if bmp.size[0] < bmp.size[1]:
        bmp = bmp.rotate(0)
    scale = SCALE

    hDC.StartDoc(img)
    hDC.StartPage()
    dib = ImageWin.Dib(bmp)
    scaled_width, scaled_height = [int(scale * i) for i in bmp.size]
    x1 = 8  # 控制位置 越大越向右 （left padding）
    y1 = 100 #越大越向下 (up padding)
    x2 = x1 + scaled_width
    y2 = y1 + scaled_height
    dib.draw(hDC.GetHandleOutput(), (x1, y1, x2, y2))
    hDC.EndPage()
    hDC.EndDoc()
    hDC.DeleteDC()

# 在二维码上方添加文字
def imgAddFont_up(oldimg,width,height):
    # 创建一张白底，长为二维码上方高度，宽为二维码宽度的图片
    im=Image.new("RGB" ,(width,up_margin),(255,255,255))
    draw = ImageDraw.Draw(im)
    fnt = ImageFont.truetype('C:\\Users\\LD\\Downloads\\auto_qr\\msyh.ttf',up_fontsize)
    msg='上海科技大学' 
    w, h = draw.textsize(msg, font=fnt)
    draw.text(((width-w)/2, 15), msg, fill='black', font=fnt)
    
    # 将新创建的图片和原二维码拼接
    blankLongImg=Image.new('RGBA',(width,up_margin+height))
    blankLongImg.paste(im,(0,0))
    blankLongImg.paste(oldimg,(0,up_margin))
    return blankLongImg

# 在二维码下方添加文字
def imgAddFont_bottom(oldimg,width,height,row1,row2):
    add_height=target_height-height
    im=Image.new("RGB" ,(width,add_height),(255,255,255))
    draw = ImageDraw.Draw(im)
    fnt = ImageFont.truetype('C:\\Users\\LD\\Downloads\\auto_qr\\msyh.ttf',bottom_fontsize)

    # 如果是正常尺寸，设备或配件
    row1=row1[:13]
    w, h = draw.textsize(row1, font=fnt)
    draw.text(((width-w)/2,20), row1, fill="black", font=fnt)

    row2=row2[:11]
    w, h = draw.textsize(row2, font=fnt)
    draw.text(((width-w)/2, 20+bottom_fontsize+5), row2, fill='black', font=fnt)

    msg='ShanghaiTech University'
    w, h = draw.textsize(msg, font=fnt)
    draw.text(((width-w)/2,20+bottom_fontsize*2+10), msg, fill="black", font=fnt)

    # 拼接底端文字图片
    blankLongImg=Image.new('RGBA',(width,target_height))
    blankLongImg.paste(im,(0,height))
    blankLongImg.paste(oldimg,(0,0))
    return blankLongImg

# 在二维码左右两端添加白边
def imgAddFont_side(oldimg,width,height):
    add_width=(target_width-width)//2
    im=Image.new("RGB" ,(add_width,height),(255,255,255))
    draw = ImageDraw.Draw(im)
    
    blankLongImg=Image.new('RGBA',(width+add_width*2,height))
    blankLongImg.paste(im,(0,0))
    blankLongImg.paste(oldimg,(add_width,0))
    blankLongImg.paste(im,(width+add_width,0))
    return blankLongImg

# 在生成的二维码（无文字）周围添加文字和白边
def add_margins(qr_code,row1,row2):
    M,N = qr_code.size
    qr_code=imgAddFont_side(qr_code,M,N)
    M,N = qr_code.size
    qr_code=imgAddFont_up(qr_code,M,N)
    M,N = qr_code.size
    qr_code=imgAddFont_bottom(qr_code,M,N,row1,row2)
    return qr_code
 
 # 上传excel表格，源数据
def upload_file():
    # 表格路径存储到entry1变量中
    entry1.delete(0, 'end')
    inputp = tk.filedialog.askopenfilename()  # askopenfilename 1次上传1个；askopenfilenames1次上传多个
    print('选中文件',inputp.split('/')[-1].split('\\')[-1])
    if inputp[-4:]!='xlsx':
        print('请上传excel表格')
    else:
        entry1.insert(0, inputp)

# 存储二维码图片的路径，当前版本直接打印，可不使用
def save_place():
    save = tk.filedialog.askdirectory()
    # print(save)
    entry2.insert(0, save)

# def gen_qrcode(input_path,save_path,size_type):
def gen_qrcode(custom,row1,row2,save_path='',input_path=''):
        input_path=os.getcwd()
        SCALE=0.55
        Infos=[]
        # 存储二维码图片的路径 默认是Excel表格的路径
        if save_path=='':
            save_path=input_path.split('/')[:-1]
            save_path='/'.join(save_path)
        # for i in range(len(datadf)):
            # 二维码中的信息
        info=custom
        # info=str(datadf['资产编码'][i])+' '+datadf['资产名称'][i]+' '+str(datadf['单价/元'][i])+' '+\
        #         datadf['资产所属单位'][i]+' '+datadf['保管人'][i]+' '+datadf['责任人'][i]+' '+datadf['入库日期'][i]
        # 用python包生成二维码
        qr = qrcode.QRCode(
            version=VERSION,
            error_correction=qrcode.constants.ERROR_CORRECT_L,
            box_size=BOX_SIZE,
            border=BORDER
        )
        qr.add_data(info)
        qr.make(fit=True)
        img = qr.make_image()
        # 计算二维码的宽度（正方形）
        width=(21 + (VERSION - 1) * 4 + BORDER * 2) * BOX_SIZE
        # img=imgAddFont(img,width,str(datadf['资产编码'][i])+'\n'+datadf['资产名称'][i])
        # 在生成的二维码（无文字）周围添加文字和白边
        img=add_margins(img,row1,row2)
        # 保存图片
        # save_file_path=os.path.join(save_path,str(datadf['资产编码'][i])+".png")
        save_file_path=os.path.join(save_path,'tmp'+".png")
        img.save(save_file_path)

        # 打印图片并删除保存的图片文件
        # 注释掉以下两行可以不打印，只查看保存的二维码图片样式
        print_img(save_file_path,SCALE)
        os.remove(save_file_path)
        print('完成')

# infodf=pd.DataFrame({})
if __name__=="__main__":
    # 图形界面的搭建
    root = tk.Tk()
    # 正常尺寸(0)/小尺寸(1)二维码
    size_type=tk.IntVar() 

    frm = tk.Frame(root)
    frm.grid(padx='2', pady='3')

    # 最上方为logo图片
    img_open = Image.open('logo.JPG').resize((150,150))
    img_png = ImageTk.PhotoImage(img_open)
    label_img = tk.Label(frm, image = img_png)
    label_img.grid(row=0, column=0,ipady='10', ipadx='10',columnspan=2)

    # 第一行是文字说明
    label_text = tk.Label(frm, text = '二维码生成工具，请输入自定义的资产标签信息')
    label_text.grid(row=1, column=0,ipady='10', ipadx='10',columnspan=2)

    # 第二行是上传文件的互动按钮以及打印源文件路径的文本框
    # btn = tk.Button(frm, text='上传文件', command=upload_file)
    custom_text = tk.Label(frm, text = '自定义的二维码文字')
    custom_text.grid(row=2, column=0, ipadx='3', ipady='3', padx='10', pady='10')
    entry1 = tk.Entry(frm, width='40')
    entry1.grid(row=2, column=1)
    
    # 第三，四行是二维码下方文字信息
    row1_text = tk.Label(frm, text = '二维码下方第一行文字')
    row1_text.grid(row=3, column=0, ipadx='3', ipady='3', padx='10', pady='10')
    entry2 = tk.Entry(frm, width='40')
    entry2.grid(row=3, column=1)
    entry2.insert(0, entry1.get())

    row2_text = tk.Label(frm, text = '二维码下方第二行文字')
    row2_text.grid(row=4, column=0, ipadx='3', ipady='3', padx='10', pady='10')
    entry3 = tk.Entry(frm, width='40')
    entry3.grid(row=4, column=1)

    # 第四行是开始生成按钮
    # btn3 = tk.Button(frm, text = '开始生成',command=lambda: gen_qrcode(entry1.get(), size_type.get(),entry2.get()))
    btn3 = tk.Button(frm, text = '开始生成',command=lambda: gen_qrcode(entry1.get(), entry2.get(),entry3.get()))
    btn3.grid(row=5, column=0, ipadx='3', ipady='3', padx='10', pady='10',columnspan=2)


    root.mainloop()

    # input_path=os.path.join('..','egate导出的原始数据.xlsx')



