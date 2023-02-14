import win32print
import win32ui
from PIL import Image, ImageWin
#
# Constants for GetDeviceCaps
#
#
# HORZRES / VERTRES = printable area
#
# HORZRES = 8
# VERTRES = 10
#
# LOGPIXELS = dots per inch
#
# LOGPIXELSX = 88
# LOGPIXELSY = 90
#
# PHYSICALWIDTH/HEIGHT = total area
#
# PHYSICALWIDTH = 110
# PHYSICALHEIGHT = 111
#
# PHYSICALOFFSETX/Y = left / top margin
#
# PHYSICALOFFSETX = 5
# PHYSICALOFFSETY = 5

# printer_name = win32print.GetDefaultPrinter ()
# print('default',printer_name)
printers = win32print.EnumPrinters(6)
choose=None
for i in printers:
    for j in i:
        if 'EasyCoder' in str(j):
            choose = i
            break
# print(choose)
# print(choose[1])

win32print.SetDefaultPrinter(choose[2])
printer_name = win32print.GetDefaultPrinter ()
print('default printer',printer_name)
file_name = "2017013429-1.png"

hDC = win32ui.CreateDC ()
hDC.CreatePrinterDC (printer_name)


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
# printable_area = hDC.GetDeviceCaps (HORZRES), hDC.GetDeviceCaps (VERTRES)
# printer_size = hDC.GetDeviceCaps (PHYSICALWIDTH), hDC.GetDeviceCaps (PHYSICALHEIGHT)
# printer_margins = hDC.GetDeviceCaps (PHYSICALOFFSETX), hDC.GetDeviceCaps (PHYSICALOFFSETY)

# #
# # Open the image, rotate it if it's wider than
# #  it is high, and work out how much to multiply
# #  each pixel by to get it as big as possible on
# #  the page without distorting.
# #
# bmp = Image.open (file_name)
# if bmp.size[0] > bmp.size[1]:
#   bmp = bmp.rotate (90)

# ratios = [1.0 * printable_area[0] / bmp.size[0], 1.0 * printable_area[1] / bmp.size[1]]
# scale = min (ratios)

# #
# # Start the print job, and draw the bitmap to
# #  the printer device at the scaled size.
# #
# hDC.StartDoc (file_name)
# hDC.StartPage ()

# dib = ImageWin.Dib (bmp)
# scaled_width, scaled_height = [int (scale * i) for i in bmp.size]
# x1 = int ((printer_size[0] - scaled_width) / 2)
# y1 = int ((printer_size[1] - scaled_height) / 2)
# x2 = x1 + scaled_width
# y2 = y1 + scaled_height
# dib.draw (hDC.GetHandleOutput (), (x1, y1, x2, y2))

# hDC.EndPage ()
# hDC.EndDoc ()
# hDC.DeleteDC ()




