import os
import copy
from threading import Thread
from tkinter import filedialog as fd
from tkinter import messagebox
from tkinter import ttk
from tkinter import *
from helpers import add_image, populateSlides

def threading():
    progress.place(anchor = CENTER, relx=0.5, rely=0.6)
    t1 = Thread(target=populateCallback())
    t1.start()

# Prevent adding of images until both fields are properly filled
def checkInputs():
    if len(image_list) > 0 and template_text.get() != "":
        B3['state'] = NORMAL
    else:
        B3['state'] = DISABLED

def imageSelectCallBack():
    selected = fd.askopenfilename(
        parent=root,
        multiple=True,
        filetypes=[('Image files', ['.jpg', '.png'])])

    fileList = root.splitlist(selected)
    image_list.clear()
    pseudoList = ""
    for f in fileList:
        image_list.append(f)

    num_label.set(str(len(root.splitlist(selected))) + " images")
    
    checkInputs()
    return

def templateSelectCallBack():
    selected = fd.askopenfilename(title="Select template presentation file",
                              filetypes=[('Presentation files', ['.pptx'])])
    template_text.set(selected)
    checkInputs()
    return


def populateCallback():
    B1['state'] = "disabled"
    B2['state'] = "disabled"
    B3['state'] = "disabled"
    resultCode = populateSlides(image_list, template_text.get(), progress, counter_label)
    if resultCode != 4:
        msg_box = messagebox.askquestion("Status", "Pictures successfully added to " + template_text.get() + "\n\nDo you want to open the file?")
    else:
        msg_box = messagebox.showerror("Status", "An error occurred.")
    
    progress.place_forget()
    progress['value'] = 0
    B1['state'] = "normal"
    B2['state'] = "normal"
    B3['state'] = "normal"
    counter_label.set("")
    return

image_list = []

root = Tk()
root.title('Powerpoint Mass Image Adder')
root.resizable(False, False)
root.geometry('500x300')

L1 = Label(root, text="Images directory", justify="left", anchor="w")
L1.grid(column=0, row=3, sticky=W, pady=(15,0), padx=(30,0))

num_label = StringVar()
num_label.set("0 images")
imgCountLabel = Label(root, textvariable=num_label, justify="left", anchor="w")
imgCountLabel.grid(column=0, row=4, sticky=W, padx=(30,0))
B1 = Button(root, text="...", command=imageSelectCallBack, width=8)
B1.grid(column=1, row=4, sticky=W)

#Entry for empty presentation file to populate with images
L2 = Label(root, text="Template presentation", justify="left", anchor="w")
L2.grid(column=0, row=6, sticky=W, pady=(15,0), padx=(30,0))

template_text = StringVar()
E2 = Entry(root, width=55, bd=5, state=DISABLED, textvariable=template_text)
E2.grid(column=0, row=7, sticky=W, padx=(30,0))
B2 = Button(root, text="...", command=templateSelectCallBack, width=8)
B2.grid(column=1, row=7, sticky=W)

# Add counter for images successfully added?

#Progress bar eventually
progress = ttk.Progressbar(root, orient = HORIZONTAL, length = 250, mode="determinate")

counter_label = StringVar()
L3 = Label(root, textvariable=counter_label, justify="left", anchor="w")
L3.place(relx=0.5, rely=0.75, anchor=CENTER)

B3 = Button(root, text="Read", command=threading, state=DISABLED)
B3.place(relx=0.5, rely=0.85, anchor=CENTER, height=32, width=150)

root.mainloop()



