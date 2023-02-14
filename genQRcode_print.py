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

VERSION=8
BOX_SIZE=10
BORDER=2
FORTEST=False

height_width_ratio=44.9/34.7

target_width=680
target_height=int(target_width*height_width_ratio)
# print('target_width',target_width)
# print('target_height',target_height)
up_margin=100
up_fontsize=70
bottom_fontsize=55

def set_up_default_printer():
    printer_name = win32print.GetDefaultPrinter()
    if 'EasyCoder' in printer_name:
        print('default printer',printer_name)
        return
    for m in range(10):
        printers = win32print.EnumPrinters(m)
        choose=None
        for i in printers:
            for j in i:
                if 'EasyCoder' in str(j):
                    choose = i
                    win32print.SetDefaultPrinter(choose[2])
                    printer_name = win32print.GetDefaultPrinter()
                    print('default printer',printer_name)
                    return

def print_img(file_name):
    printer_name = win32print.GetDefaultPrinter()
    hDC = win32ui.CreateDC ()
    hDC.CreatePrinterDC (printer_name)
    # if file_name is path
    img=file_name
    bmp = Image.open(img)
    if bmp.size[0] < bmp.size[1]:
        bmp = bmp.rotate(0)
    scale = 0.55

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

def imgAddFont_up(oldimg,width,height):
    im=Image.new("RGB" ,(width,up_margin),(255,255,255))
    draw = ImageDraw.Draw(im)
    fnt = ImageFont.truetype('C:\\Users\\LD\\Downloads\\auto_qr\\msyh.ttf',up_fontsize)
    msg='上海科技大学' 
    w, h = draw.textsize(msg, font=fnt)
    draw.text(((width-w)/2, 15), msg, fill='black', font=fnt)
    
    blankLongImg=Image.new('RGBA',(width,up_margin+height))
    blankLongImg.paste(im,(0,0))
    blankLongImg.paste(oldimg,(0,up_margin))
    return blankLongImg

def imgAddFont_bottom(oldimg,width,height,info_num,info_name,note):
    add_height=target_height-height
    im=Image.new("RGB" ,(width,add_height),(255,255,255))
    draw = ImageDraw.Draw(im)
    fnt = ImageFont.truetype('C:\\Users\\LD\\Downloads\\auto_qr\\msyh.ttf',bottom_fontsize)

    if isinstance(note,str) and len(note)!=0 and 'Z' in info_num:
        fnt = ImageFont.truetype('C:\\Users\\LD\\Downloads\\auto_qr\\msyh.ttf',48)
        w, h = draw.textsize(info_num, font=fnt)
        draw.text(((width-w)/2,0), info_num, fill="black", font=fnt)

        info_name=info_name[:11]
        w, h = draw.textsize(info_name, font=fnt)
        draw.text(((width-w)/2, 0+bottom_fontsize+5), info_name, fill='black', font=fnt)

        w, h = draw.textsize(note, font=fnt)
        draw.text(((width-w)/2, 0+bottom_fontsize*2+10), note, fill='black', font=fnt)

        msg='ShanghaiTech University'
        w, h = draw.textsize(msg, font=fnt)
        draw.text(((width-w)/2,0+bottom_fontsize*3+15), msg, fill="black", font=fnt)

    else:
        w, h = draw.textsize(info_num, font=fnt)
        draw.text(((width-w)/2,20), info_num, fill="black", font=fnt)

        info_name=info_name[:11]
        w, h = draw.textsize(info_name, font=fnt)
        draw.text(((width-w)/2, 20+bottom_fontsize+5), info_name, fill='black', font=fnt)

        msg='ShanghaiTech University'
        w, h = draw.textsize(msg, font=fnt)
        draw.text(((width-w)/2,20+bottom_fontsize*2+10), msg, fill="black", font=fnt)
    
    blankLongImg=Image.new('RGBA',(width,target_height))
    # print('im',im.size)
    # print('oldimg',oldimg.size)
    # print('blankLongImg',blankLongImg.size)
    blankLongImg.paste(im,(0,height))
    blankLongImg.paste(oldimg,(0,0))
    return blankLongImg

def imgAddFont_side(oldimg,width,height):
    add_width=(target_width-width)//2
    im=Image.new("RGB" ,(add_width,height),(255,255,255))
    draw = ImageDraw.Draw(im)
    
    blankLongImg=Image.new('RGBA',(width+add_width*2,height))
    blankLongImg.paste(im,(0,0))
    blankLongImg.paste(oldimg,(add_width,0))
    blankLongImg.paste(im,(width+add_width,0))
    return blankLongImg

def add_margins(qr_code,datadf,i):
    M,N = qr_code.size
    qr_code=imgAddFont_side(qr_code,M,N)
    M,N = qr_code.size
    qr_code=imgAddFont_up(qr_code,M,N)
    M,N = qr_code.size
    qr_code=imgAddFont_bottom(qr_code,M,N,str(datadf['资产编码'][i]),datadf['资产名称'][i],note=datadf['备注'][i])
    return qr_code
 
def upload_file():
    entry1.delete(0, 'end')
    inputp = tk.filedialog.askopenfilename()  # askopenfilename 1次上传1个；askopenfilenames1次上传多个
    print('选中文件',inputp.split('/')[-1].split('\\')[-1])
    if inputp[-4:]!='xlsx':
        print('请上传excel表格')
    else:
        entry1.insert(0, inputp)
    # print(entry1.get())

def save_place():
    save = tk.filedialog.askdirectory()
    # print(save)
    entry2.insert(0, save)

def gen_qrcode(input_path,save_path):
        # datadf=pd.read_excel(input_path,encoding='utf-8')
        print(input_path)
        # datadf=pd.read_excel(input_path,encoding='utf-8',nrows=2)
        if FORTEST:
            datadf=pd.read_excel(input_path,nrows=2)
        else:
            datadf=pd.read_excel(input_path)
        Infos=[]
        if save_path=='':
            save_path=input_path.split('/')[:-1]
            save_path='/'.join(save_path)
            # print(save_path)
        # Nums=[]
        for i in range(len(datadf)):
            # 二维码中的信息
            all_cols=list(datadf.columns)
            info_list=['资产编码','资产名称','财政资产编号','单价/元','资产所属单位','保管人','责任人','入库日期']
            info = ''
            for col in info_list:
                if col in all_cols:
                    info+=str(datadf[col][i])+' '
            qr = qrcode.QRCode(
        	    version=VERSION,
        	    error_correction=qrcode.constants.ERROR_CORRECT_L,
        	    box_size=BOX_SIZE,
        	    border=BORDER
            )
            qr.add_data(info)
            qr.make(fit=True)
            img = qr.make_image()
            width=(21 + (VERSION - 1) * 4 + BORDER * 2) * BOX_SIZE
            # img=imgAddFont(img,width,str(datadf['资产编码'][i])+'\n'+datadf['资产名称'][i])
            img=add_margins(img,datadf,i)
            save_file_path=os.path.join(save_path,str(datadf['资产编码'][i])+".png")
            img.save(save_file_path)
            print_img(save_file_path)
            os.remove(save_file_path)
        print('完成，共生成并打印',len(datadf),'个二维码')

# infodf=pd.DataFrame({})

root = tk.Tk()

frm = tk.Frame(root)
frm.grid(padx='2', pady='3')
img_open = Image.open('logo.JPG').resize((150, 150))
img_png = ImageTk.PhotoImage(img_open)
label_img = tk.Label(frm, image = img_png)
label_img.grid(row=0, column=0,ipady='10', ipadx='10',columnspan=2)



label_text = tk.Label(frm, text = '二维码生成工具，请将egate导出的资产标签信息（.xlsx文件）上传，\n并选择存放所生成二维码的文件夹（默认为.xlsx文件所在的文件夹）')
label_text.grid(row=1, column=0,ipady='10', ipadx='10',columnspan=2)

btn = tk.Button(frm, text='上传文件', command=upload_file)
btn.grid(row=2, column=0, ipadx='3', ipady='3', padx='10', pady='10')
entry1 = tk.Entry(frm, width='40')
entry1.grid(row=2, column=1)

btn2 = tk.Button(frm, text='存放路径', command=save_place)
btn2.grid(row=3, column=0, ipadx='3', ipady='3', padx='10', pady='10')
entry2 = tk.Entry(frm, width='40')
entry2.grid(row=3, column=1)


btn3 = tk.Button(frm, text = '开始生成',command=lambda: gen_qrcode(entry1.get(), entry2.get()))
btn3.grid(row=4, column=0, ipadx='3', ipady='3', padx='10', pady='10',columnspan=2)


root.mainloop()

# input_path=os.path.join('..','egate导出的原始数据.xlsx')


