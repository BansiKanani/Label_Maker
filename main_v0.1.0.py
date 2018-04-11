import os
from PIL import Image
from PIL import ImageFont
from PIL import ImageDraw
import xlrd
import time
from PyPDF2 import PdfFileMerger, PdfFileReader

def xls_to_list():

    file_list = os.listdir(os.getcwd())
    xlsfile = ([f for f in file_list if 'xlsx' in f])[0]

    #list = [Label_NO, Type, Length, Width, top, left, bottom, right, borring point, Length, Width, Quantity]
    #list = [    0   ,  1  ,   2   ,   3  ,  4 ,  5  ,   6   ,   7   ,     8       ,    9  ,  10   ,   11  ]
    dataList = []

    #Open workbook and select sheet
    workBook = xlrd.open_workbook(xlsfile)
    sheet = workBook.sheet_by_index(0)

    #read all the available data
    for row in range(2, sheet.nrows):

        #Remove .0 from the Excel data
        tempStr = str(int(sheet.cell_value(row, 9)))
        for i in range(2):
            dataList.append([tempStr, (sheet.cell_value(row, 8)).upper(),
                             str(int(sheet.cell_value(row, 10))), str(int(sheet.cell_value(row, 11))),
                             str(sheet.cell_value(row, 15)), str(sheet.cell_value(row, 16)),
                             str(sheet.cell_value(row, 17)), str(sheet.cell_value(row, 18))])
        for i in range(2):
            dataList.append([sheet.cell_value(row, 12), (sheet.cell_value(row, 8)).upper(),
                             str(int(sheet.cell_value(row, 13))), str(int(sheet.cell_value(row, 14))),
                             str(sheet.cell_value(row, 15)), str(sheet.cell_value(row, 16)),
                             str(sheet.cell_value(row, 17)), str(sheet.cell_value(row, 18)), str(sheet.cell_value(row, 22))])

        #Label
        #Remark
        #VR Ply Length
        #Width
        #Label
        #HR Ply Length
        #Width
        #top, left, bottom, right
        #Self - Label
        #length
        #width
        #Quantity


    print(' XLSX data fetched!')
    return dataList

def make_label(lblDataList, font_file, tf, ff):

    print (' Making Labels from data.....')
    allImages = []

    #select fonts
    fontNumber = ImageFont.truetype(font_file, 80)
    fontText = ImageFont.truetype(font_file, 45)
    fontSize = ImageFont.truetype(font_file, 40)

    for lbl in lblDataList:

        tff = ''
        if lbl[1] == 'T':
            if 'A' in lbl[0]:
                tff=ff
            else:
                tff=tf
        elif lbl[1] =='F':
            if 'A' in lbl[0]:
                tff=tf
            else:
                tff=ff

        img = Image.open(tff)
        draw = ImageDraw.Draw(img)

        #Print label data on images
        draw.text((350, 150), lbl[0], (0,0,0), font=fontNumber) #Label number
        draw.text((590, 110), '('+lbl[1]+')', (0,0,0), font=fontText) #Type
        draw.text((290, 280), lbl[2]+' X '+lbl[3], (0,0,0), font=fontText) #size

        draw.text((380, 25), lbl[4], (0,0,0), font=fontSize) #top
        draw.text((35, 195), lbl[5], (0,0,0), font=fontSize) #left
        draw.text((145, 195), lbl[8], (0,0,0), font=fontSize) #boring point
        draw.text((380, 365), lbl[6], (0,0,0), font=fontSize) #bottom
        draw.text((735, 190), lbl[7], (0,0,0), font=fontSize) #right
        allImages.append(img)

    print(' Labels are made!')
    return allImages

def merge_images(allImages):

    print(' Making PDF, Please wait.....')
    (width, height) = allImages[0].size
    #total counter counts all iteration / inPageCounter count images in each page / pages stores number
    totalCounter, inPageCounter, pages = 0, 0, 1

    #Do for all images
    while totalCounter < len(allImages):
        #Create while image to print labels in it.
        mixedImg = Image.new('RGB', (3 * width, 8 * height), (255,255,255))
        #Check if page is full then reset inImgCounter.
        while inPageCounter < 24:
            for row in range(8):
                for col in range(3):
                    if totalCounter < len(allImages):
                        mixedImg.paste(im=allImages[totalCounter], box=(col*width, row*height))
                    totalCounter += 1
                    inPageCounter += 1

        mixedImg.save('Label-{}.pdf'.format(pages), 'PDF', resolution=100.0)
        pages += 1
        inPageCounter = 0
        #Delete old image from memory.
        mixedImg.close()

    time.sleep(5)

def merge_pdfs():

    print(' Merging PDFs to Single file...')
    file_list = os.listdir(os.getcwd())
    pdf_list = ([f for f in file_list if 'pdf' and '-' in f])

    merger = PdfFileMerger()
    for f in pdf_list:
        with open(f, "rb") as fl:
            merger.append((PdfFileReader(fl)))
        os.remove(f)

    merger.write("Labels_Print_This_File.pdf")
    print('\n\n Done!')
    time.sleep(10.0)

font_file = r'source/arial.ttf'
t_image = r'source/t.png'
f_image = r'source/f.png'

dataList = xls_to_list() #return LIST containing xls Data
allImages = make_label(dataList, font_file, t_image, f_image) #return LIST containing Image Objects
merge_images(allImages)
merge_pdfs()