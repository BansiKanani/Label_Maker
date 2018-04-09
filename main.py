from PIL import Image
from PIL import ImageFont
from PIL import ImageDraw
import xlrd
import os
import time

def xls_to_list(xlsfile):

    #list = [Label_NO, Type, Length, Width, top, left, bottom, right]
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
                             str(sheet.cell_value(row, 17)), str(sheet.cell_value(row, 18))])
    print(' XLSX data fetched!')
    return dataList


def make_label(lblDataList, font_file, tf, ff):

    print (' Making Labels from data...')
    allImages = []

    #select fonts
    fontNumber = ImageFont.truetype(font_file, 80)
    fontText = ImageFont.truetype(font_file, 50)
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
        draw.text((260, 140), lbl[0]+'-('+lbl[1]+')', (0,0,0), font=fontNumber) #Label number, Type
        draw.text((260, 260), lbl[2]+' X '+lbl[3], (0,0,0), font=fontText) #size

        draw.text((380, 15), lbl[4], (0,0,0), font=fontSize) #top
        draw.text((30, 200), lbl[5], (0,0,0), font=fontSize) #left
        draw.text((380, 390), lbl[6], (0,0,0), font=fontSize) #bottom
        draw.text((740, 200), lbl[7], (0,0,0), font=fontSize) #right
        allImages.append(img)

    print(' Labels are made!')
    return allImages


def merge_images(allImages):

    print(' Making PDF...')
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

    print('-------DONE--------\n You may now close this window!')
    time.sleep(5)

excel_file = r'LabelData.xlsx'
font_file = r'source/Consolas.ttf'
t_image = r'source/t.png'
f_image = r'source/f.png'

dataList = xls_to_list(excel_file) #return LIST containing xls Data
allImages = make_label(dataList, font_file, t_image, f_image) #return LIST containing Image Objects
merge_images(allImages)