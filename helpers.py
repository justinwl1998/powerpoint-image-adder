import os
from pptx import Presentation
#from pptx import PlaceholderPicture
from PIL import Image
import re

def add_image(slide, placeholder_id, img_path):
    placeholder = slide.placeholders[placeholder_id]
    im = Image.open(img_path)
    width, height = im.size

    placeholder.height = height
    placeholder.width = width

    placeholder = placeholder.insert_picture(img_path)

    image_ratio = width / height
    placeholder_ratio = placeholder.width / placeholder.height
    ratio_difference = placeholder_ratio - image_ratio

    if ratio_difference > 0:
        difference_on_each_side = ratio_difference / 2
        placeholder.crop_left = -difference_on_each_side
        placeholder.crop_right = -difference_on_each_side
    else:
        difference_on_each_side = -ratio_difference / 2
        placeholder.crop_left = -difference_on_each_side
        placeholder.crop_right = -difference_on_each_side

def populateSlides(images, presentation, progress):
    #debug to remove test pptx
    #if os.path.exists('test.pptx'):
    #    os.remove('test.pptx')
    
    prs = Presentation(presentation)

    index = 0
    imgIndex = 0
    breakOut = False

    while not breakOut:
        if index >= len(prs.slides):
            break
        slide = prs.slides[index]
        for shape in slide.placeholders:
            if type(shape).__name__ == "PlaceholderPicture":
                print("Placeholder is occupied! Finding next available one...")
                continue
            
            if type(shape).__name__ == "PicturePlaceholder":
                if imgIndex >= len(images):
                    breakOut = True
                    break
                add_image(slide, shape.placeholder_format.idx, images[imgIndex])
                imgIndex += 1
        index += 1
        #progress.set((i/len(prs.slides)) * 100)

    progress['value'] = 100
    prs.save(presentation)


    
