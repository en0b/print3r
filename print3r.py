import time
from PIL import Image, ImageEnhance, ImageTk
from thermalprinter import *
import tkinter as tk
from tkinter import filedialog, messagebox
import threading

WIDTH_PIXELS = 384 #Total width available with the thermal printer
CHUNK_LINES = 20 #When printing, the images are split into chunks.
                 #This is because otherwise random character printing can occur.

globalSourceImage = 0
globalDisplayImage = 0
lbl_image_orig = 0
lbl_image_disp = 0
contrast = 0
brightness = 0
printerThread = 0

def image_slicer_and_scaler(img):
    """slice an image into parts slice_size tall"""
    width, height = img.size
    upper = 0
    left = 0
    slices = (0 + round(img.height / CHUNK_LINES))

    count = 1
    output = []
    for slice in range(slices):
        #if we are at the end, set the lower bound to be the bottom of the image
        if count == slices:
            lower = height
        else:
            lower = int(count * CHUNK_LINES)  

        bbox = (left, upper, width, lower)
        print(left)
        print(upper)
        print(width)
        print(lower)
        working_slice = img.crop(bbox)
        upper += CHUNK_LINES
        #save the slice
        output.append(working_slice)
        count +=1
    return output

def repaintImages():
    global lbl_image_orig
    global lbl_image_disp 
    global globalSourceImage
    global globalDisplayImage   

    #do the corrections here
    copyImg = globalSourceImage
    BriEnhancer = ImageEnhance.Brightness(copyImg)
    copyImg = BriEnhancer.enhance(1+brightness)
    CoEnhancer = ImageEnhance.Contrast(copyImg)
    copyImg = CoEnhancer.enhance(1+contrast)

    globalDisplayImage = copyImg.convert('1')
    try:
        lbl_image_orig.destroy()
        lbl_image_disp.destroy()
    except:
        pass

    tmp = ImageTk.PhotoImage(globalSourceImage)
    lbl_image_orig = tk.Label(image=tmp)
    lbl_image_orig.image = tmp
    lbl_image_orig.place(x=10,y=250)
    tmp = ImageTk.PhotoImage(globalDisplayImage)
    lbl_image_disp = tk.Label(image=tmp)
    lbl_image_disp.image = tmp
    lbl_image_disp.place(x=400,y=250)

def openImage():
    global globalSourceImage
    global globalDisplayImage
    imageFileName = filedialog.askopenfilename()
    try:
        globalSourceImage = Image.open(imageFileName)
        if(globalSourceImage.width > globalSourceImage.height):
            globalSourceImage = globalSourceImage.rotate(90, expand=True)
        factor = WIDTH_PIXELS/globalSourceImage.width
        globalSourceImage = globalSourceImage.resize([round(globalSourceImage.width*factor), round(globalSourceImage.height*factor)])

    except:
        messagebox.showerror("Error", "Can not open file, check if valid image was chosen.")
    repaintImages()

def countBlack (image):
    blacks = 0
    for x in range(image.width):
        for y in range(image.height):
            if image.getpixel((x,y)) < 0.0001:
                blacks += 1
    return blacks

def printThreadFcn(evt):
    while True:
        if evt.wait(1):
            slices = image_slicer_and_scaler(globalDisplayImage)
            with ThermalPrinter(port='COM6', heat_time = 110) as printer:
                #32 characters per line (normal, bold, double height)
                #14 characters per line (double width)
                for slice in slices:
                    printer.image(slice)
                    ratio = countBlack(slice)/(384*20)
                    ratio = 1.3 * max(0, ratio - 0.5)
                    time.sleep(ratio)
                printer.feed(2)
            evt.clear()


def printImage():
    global print_event
    if globalSourceImage == 0:
        messagebox.showerror("Error", "No image loaded. Load using open image button")
    else:
        if print_event.is_set():
            messagebox.showerror("Error", "Another print is running. Please wait until it is completed.")
        else:
            print_event.set()

def incCo():
    global contrast
    contrast = contrast + 0.1
    repaintImages()

def decCo():
    global contrast
    contrast = contrast - 0.1
    repaintImages()

def incBri():
    global brightness
    brightness = brightness + 0.1
    repaintImages()

def decBri():
    global brightness
    brightness = brightness - 0.1
    repaintImages()

def resetCoBri():
    global brightness
    global contrast
    brightness = 0
    contrast = 0
    repaintImages()

window = tk.Tk()
window.configure(background="#3e5757")
window.geometry("800x800")
lbl_title = tk.Label(text="Thermal Printer Tool", fg="#b9bcba", bg="#3e5757", width=40, height=3, font=("Arial", 25))
lbl_title.place(x=0, y=0)

btn_openImage = tk.Button(    text="open image...",    width=13,    height=2,    bg="#5c7474",    fg="#b9bcba",    font=("Arial", 10), command=openImage)
btn_openImage.place(x=10, y=100)

btn_print = tk.Button(    text="print",    width=13,    height=2,    bg="#5c7474",    fg="#b9bcba",    font=("Arial", 10), command=printImage)
btn_print.place(x=130, y=100)

btn_incCo = tk.Button(    text="+C",    width=5,    height=1,    bg="#5c7474",    fg="#b9bcba",    font=("Arial", 10), command = incCo)
btn_incCo.place(x=10, y=150)

btn_decCo = tk.Button(    text="-C",    width=5,    height=1,    bg="#5c7474",    fg="#b9bcba",    font=("Arial", 10), command = decCo)
btn_decCo.place(x=10, y=185)

btn_incBri = tk.Button(    text="+B",    width=5,    height=1,    bg="#5c7474",    fg="#b9bcba",    font=("Arial", 10), command = incBri)
btn_incBri.place(x=70, y=150)

btn_decBri = tk.Button(    text="-B",    width=5,    height=1,    bg="#5c7474",    fg="#b9bcba",    font=("Arial", 10), command = decBri)
btn_decBri.place(x=70, y=185)

btn_reset = tk.Button(    text="reset",    width=8,    height=2,    bg="#5c7474",    fg="#b9bcba",    font=("Arial", 10), command = resetCoBri)
btn_reset.place(x=130, y=160)

print_event = threading.Event()
printerThread = threading.Thread(target=printThreadFcn,args=[print_event], daemon=True).start()

window.mainloop()
