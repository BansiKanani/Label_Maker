from PIL import Image

def merge_images(allImages):
    (width, height) = allImages[0].size
    #mixedImg = Image.new('RGB', (3*width, 8*height), 43)
    #mixedImg.paste((200, 200, 200), [0, 0, mixedImg.size[0], mixedImg.size[1]])

    totalCounter = 0
    inPageCounter = 0
    pages = 1

    while totalCounter < len(allImages): #do for all images
        mixedImg = Image.new('RGB', (3 * width, 8 * height), (255,255,255))
        while inPageCounter < 24:    #check if page is full, if so then add to list.
            for row in range(8):
                for col in range(3):
                    if totalCounter < len(allImages):
                        mixedImg.paste(im=allImages[totalCounter], box=(col*width, row*height))
                    totalCounter += 1
                    inPageCounter += 1
        mixedImg.save('temp/{}.jpg'.format(pages))
        pages += 1
        inPageCounter = 0
        mixedImg.close()
