import os
import copy
from tkinter import filedialog as fd
from tkinter import messagebox
from tkinter import ttk
from tkinter import *
from pptx import Presentation
from pptx.util import Inches
from PIL import Image

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

def populateSlides(images, presentation):
    print(images)
    print(images[0])
    prs = Presentation(presentation)

    index = 0
    imgIndex = 0

    while index < len(prs.slides) and imgIndex < len(images):
        slide = prs.slides[index]
        for shape in slide.placeholders:
            if 'Picture Placeholder' in shape.name:
                print("attempting to add image: ", images[imgIndex])
                add_image(slide, shape.placeholder_format.idx, images[imgIndex])
                imgIndex += 1
        index += 1
