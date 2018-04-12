import os
import time
import xlrd
from PIL import Image, ImageFont, ImageDraw, ImageOps
from PyPDF2 import PdfFileMerger, PdfFileReader

def xls_to_list():
    print(' 1) Excel File reading!')
    file_list = os.listdir(os.getcwd())
    xlsfile = ([f for f in file_list if 'xlsx' in f])[0]
    all_data_list, current_row = [], []
    workBook = xlrd.open_workbook(xlsfile)
    sheet = workBook.sheet_by_index(0)
    for row in range(2, sheet.nrows):
        for col in range(0, sheet.ncols):
            current_row.append(sheet.cell_value(row, col))
        all_data_list.append(current_row)
        current_row = []
    print('--> Done!')
    return all_data_list

def make_label(lblDataList, source_files):

    print (' 2) Making Labels.....')
    allImages = []

    # Names of columns and it's index in xlsx file.
    quantity_vr_hr = 4
    remark_vr_hr = 8
    vr_ply_lbl = 9
    vr_ply_length = 10
    vr_ply_width = 11

    #hr_ply_remark =
    hr_ply_lbl = 12
    hr_ply_length = 13
    hr_ply_width = 14

    edgeband_top = 15
    edgeband_left = 16
    edgeband_bottom = 17
    edgeband_righ = 18

    boring_point = 22
    self_lbl = 24
    self_length = 25
    self_width = 26
    self_quantity = 27

    for lbl in lblDataList:
        # IF you found any VR PLY label then.

        template_file = ''
        if lbl[vr_ply_lbl] is not '':
            if lbl[remark_vr_hr] == 'T': template_file = source_files['t_file']
            else: template_file = source_files['f_file']

            for i in range(0, int(lbl[quantity_vr_hr])):
                allImages.append(print_label(int(lbl[vr_ply_lbl]), lbl[remark_vr_hr], template_file, int(lbl[vr_ply_length]),
                                             int(lbl[vr_ply_width]), lbl[edgeband_top], lbl[edgeband_left], lbl[edgeband_bottom],
                                             lbl[edgeband_righ], lbl[boring_point], source_files['font_file']))

        if lbl[hr_ply_lbl] is not '':
            if lbl[remark_vr_hr] == 'T': template_file = source_files['f_file']
            else: template_file = source_files['t_file']

            for i in range(0, int(lbl[quantity_vr_hr])):
                allImages.append(print_label(lbl[hr_ply_lbl], lbl[remark_vr_hr], template_file, int(lbl[hr_ply_length]),
                                             int(lbl[hr_ply_width]), lbl[edgeband_top], lbl[edgeband_left],
                                             lbl[edgeband_bottom],
                                             lbl[edgeband_righ], lbl[boring_point], source_files['font_file']))

    for lbl in lblDataList:
        if lbl[self_lbl] is not '':
            template_file = source_files['s_file']
            for i in range(0, int(lbl[self_quantity])):
                allImages.append(print_label(lbl[self_lbl], '', template_file, int(lbl[self_length]),
                                             int(lbl[self_width]), '', '', '', '', '', source_files['font_file']))


    print('--> Done!')
    return allImages

def print_label(name, remark , templateFile, length, width, top, left, bottom, right, boringPoint, fontFile):

    img = Image.open(templateFile)
    draw = ImageDraw.Draw(img)

    fontBig = ImageFont.truetype(fontFile, 80)
    fontMedium = ImageFont.truetype(fontFile, 45)
    fontSmall = ImageFont.truetype(fontFile, 40)


    # FOR LEFT and RIGHT rotated fonts.
    leftFont = Image.new('L', fontSmall.getsize(str(left)))
    rightFont = Image.new('L', fontSmall.getsize(str(right)))
    draw_l = ImageDraw.Draw(leftFont)
    draw_r = ImageDraw.Draw(rightFont)
    draw_l.text((0, 0), str(left), font=fontSmall, fill=255)
    draw_r.text((0, 0), str(right), font=fontSmall, fill=255)
    lt = leftFont.rotate(90, expand=1)
    rt = leftFont.rotate(270, expand=1)
    v_center = int((img.size[1] - leftFont.size[1])/2)
    img.paste( '#000000', (50, v_center), lt)
    img.paste( '#000000', (735, v_center), rt)



    # Print label data on images
    #draw.text((590, 110), '(' + remark + ')', (0, 0, 0), font=fontMedium)  # Type
    name_center = int((img.size[0]-fontBig.getsize(str(name))[0])/2)
    text_center = int((img.size[0]-fontMedium.getsize(str(str(length)+' X '+str(width)))[0])/2)
    tb_center = int((img.size[0]-fontSmall.getsize(str(top))[0])/2)

    draw.text((name_center, 150), str(name), (0, 0, 0), font=fontBig)  # Label number
    draw.text((610, 120), remark, (0, 0, 0), font=fontMedium)  # Type
    draw.text((text_center, 280), str(length) + ' X ' + str(width), (0, 0, 0), font=fontMedium)  # size
    draw.text((tb_center, 25), str(top), (0, 0, 0), font=fontSmall)  # top
    draw.text((145, 195), str(boringPoint), (0, 0, 0), font=fontSmall)  # boring point
    draw.text((tb_center, 365), str(bottom), (0, 0, 0), font=fontSmall)  # bottom
    return img

def merge_images(allImages):

    print(' 3) Making PDF.....')
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
    print('--> Done!')

def merge_pdfs():

    print(' 4) Merging PDF pages...')
    file_list = os.listdir(os.getcwd())
    pdf_list = ([f for f in file_list if 'pdf' and '-' in f])

    merger = PdfFileMerger()
    for f in pdf_list:
        with open(f, "rb") as fl:
            merger.append((PdfFileReader(fl)))
        os.remove(f)

    merger.write("Labels_Print_This_File.pdf")
    print('--> Done!')


source_files = {'font_file': 'source/arial.ttf', 't_file': 'source/t.png', 'f_file':'source/f.png', 's_file':'source/s.png'}

all_Data_List = xls_to_list() #return LIST containing xls Data
allImages = make_label(all_Data_List, source_files) #return LIST containing Image Objects
merge_images(allImages)
merge_pdfs()
print('\n\n >> Message : You can find the labels PDF file in the same folder! \n >> Message : You can now close this window!')
time.sleep(15)