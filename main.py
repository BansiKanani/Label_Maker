from PIL import Image
from PIL import ImageFont
from PIL import ImageDraw
import xlrd
import os
import random

def xls_to_list(xlsfile):

    dataList = [] #list = [lbl_NO, Type, Length, Width]
    workBook = xlrd.open_workbook(xlsfile)
    sheet = workBook.sheet_by_index(0)
    for row in range(2, sheet.nrows):
        dataList.append([sheet.cell_value(row, 9), (sheet.cell_value(row, 8)).upper(), sheet.cell_value(row, 10), sheet.cell_value(row, 11)])
        dataList.append([sheet.cell_value(row, 12), (sheet.cell_value(row, 8)).upper(), sheet.cell_value(row, 13), sheet.cell_value(row, 14)])

    return dataList

def make_label(lblDataList):

    allImages = []
    fontNumber = ImageFont.truetype("font/Consolas.ttf", 80)
    fontText = ImageFont.truetype("font/Consolas.ttf", 50)

    for lbl in lblDataList:

        if type(lbl[0]) == float: lblTemp = str(int(lbl[0]))
        else: lblTemp = str(lbl[0])

        if lbl[1] == 'T': img = Image.open("source/t.png")
        elif lbl[1] =='F': img = Image.open("source/f.png")
        draw = ImageDraw.Draw(img)

        draw.text((260, 140), lblTemp+'-('+lbl[1].upper()+')', (0,0,0), font=fontNumber)
        draw.text((260, 260), str(int(lbl[2]))+' X '+str(int(lbl[3])), (0,0,0), font=fontText)

        #print(lblTemp+'-('+lbl[1].upper()+')')
        #print(str(lbl[2])+'x'+str(lbl[3]))



        #draw.text((400, 150), '-'+lbl[1].upper(), (0,0,0), font=fontText)
        #img.save('temp/{0}-{1}.jpg'.format(i[0], i[1]))
        allImages.append(img)

    return allImages

def merge_images(allImages):

    (width, height) = allImages[0].size
    mixedList = []
    totalCounter, inPageCounter, pages = 0, 0, 1

    while totalCounter < len(allImages): #do for all images
        mixedImg = Image.new('RGB', (3 * width, 8 * height), (255,255,255))
        while inPageCounter < 24:    #check if page is full, if so then add to list.
            for row in range(8):
                for col in range(3):
                    if totalCounter < len(allImages):
                        mixedImg.paste(im=allImages[totalCounter], box=(col*width, row*height))
                    totalCounter += 1
                    inPageCounter += 1
        #mixedImg.save('temp/{}.png'.format(pages))
        mixedImg.save('temp/{}.pdf'.format(pages), 'PDF', resolution=100.0)
        pages += 1
        inPageCounter = 0
        mixedImg.close()

    return mixedList



#lblType = ['t', 'f']
#lblRange = range(1, random.randint(1, 500))
#rand_items = []
#for i in lblRange: rand_items.append([i, lblType[random.randint(0,1)]])

file_location = 'TestData.xlsx'
xlsList = xls_to_list(file_location) #return LIST containing xls Data
allImages = make_label(xlsList) #return LIST containing Image Objects
merge_images(allImages)