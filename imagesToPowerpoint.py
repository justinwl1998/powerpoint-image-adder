import os
import copy
from pptx import Presentation
from pptx.util import Inches
from PIL import Image

def add_image(slide, placeholder_id):
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

if os.path.exists('./test.pptx'):
    os.remove('test.pptx')

img_path = "stev.png"
path = "template.pptx"

templatePres = Presentation(path)
prs = Presentation()

template = templatePres.slides[0]

for shape in template.placeholders:
    print('%d %s' % (shape.placeholder_format.idx, shape.name))

    if "Picture Placeholder" in shape.name:
        img = shape.insert_picture(img_path)


    
templatePres.save("test.pptx")
#os.startfile("test.pptx")
#prs.save('test.pptx')

#Todo:

# Audomate adding images to an existing powerpoint presentation

# ask for directory where images to use in presentation
# ask for the powerpoint presentation to add to
# go through each slide after the title and add four images maximum
#  to each slide

# learn how add_picture() can be used

# make a powerpoint slide style where four images can be added to it

