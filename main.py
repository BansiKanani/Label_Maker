from PIL import Image
from PIL import ImageFont
from PIL import ImageDraw
import join
import random

def make_label(imageList):
    allImages = []
    font = ImageFont.truetype("font/adventpro-regular.ttf", 40)
    for i in imageList:
        if i[1] == 'T': img = Image.open("source/t.jpg")
        elif i[1] == 'F': img = Image.open("source/f.jpg")
        draw = ImageDraw.Draw(img)
        draw.text((110, 40),str(i[0]),(0,0,0),font=font)
        #img.save('temp/{0}-{1}.jpg'.format(i[0], i[1]))
        allImages.append(img)
    return allImages


typ = ['T', 'F']
labelRange = range(1, random.randint(1, 150))
rand_items = []
for i in labelRange: rand_items.append([i, typ[random.randint(0,1)]])


allImages = make_label(rand_items) #return image list
join.merge_images(allImages)