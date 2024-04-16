import time
from PIL import Image, ImageEnhance, ImageTk, ImageGrab
from thermalprinter import *
import tkinter as tk
from tkinter import filedialog, messagebox
import threading

WIDTH_PIXELS = 384  # Total width available with the thermal printer
CHUNK_LINES = 20  # When printing, the images are split into chunks.
MIN_WINDOW_WIDTH = 400

globalSourceImage = None
globalDisplayImage = None
lbl_image_orig = None
lbl_image_disp = None
contrast = 0
brightness = 0
printerThread = None
print_cancel_flag = False  # This is the new flag for canceling the print operation.

# New function to handle images from the clipboard
def handle_clipboard_image():
    global globalSourceImage, globalDisplayImage
    try:
        img = ImageGrab.grabclipboard()
        if img is None:
            messagebox.showerror("Error", "No image in clipboard")
            return
        globalSourceImage = img
        process_image()
        repaintImages()
    except Exception as e:
        messagebox.showerror("Error", f"Failed to process clipboard image: {e}")

def process_image():
    global globalSourceImage, globalDisplayImage
    if globalSourceImage.width > WIDTH_PIXELS:
        factor = WIDTH_PIXELS / globalSourceImage.width
        globalSourceImage = globalSourceImage.resize([round(globalSourceImage.width * factor), round(globalSourceImage.height * factor)], Image.ANTIALIAS)
    globalDisplayImage = globalSourceImage


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

    if globalSourceImage is None:
        return  # If no image is loaded, we don't need to repaint or resize anything.

    # Apply brightness and contrast adjustments
    copyImg = globalSourceImage
    BriEnhancer = ImageEnhance.Brightness(copyImg)
    copyImg = BriEnhancer.enhance(1+brightness)
    CoEnhancer = ImageEnhance.Contrast(copyImg)
    copyImg = CoEnhancer.enhance(1+contrast)

    globalDisplayImage = copyImg.convert('1')

    # Destroy previous labels if they exist
    try:
        lbl_image_orig.destroy()
        lbl_image_disp.destroy()
    except:
        pass

    # Create new labels for the original and adjusted images
    tmp_orig = ImageTk.PhotoImage(globalSourceImage)
    lbl_image_orig = tk.Label(image=tmp_orig)
    lbl_image_orig.image = tmp_orig  # Keep a reference!
    lbl_image_orig.place(x=10, y=250)

    tmp_disp = ImageTk.PhotoImage(globalDisplayImage)
    lbl_image_disp = tk.Label(image=tmp_disp)
    lbl_image_disp.image = tmp_disp  # Keep a reference!
    lbl_image_disp.place(x=10 + globalSourceImage.width + 10, y=250)  # Adjust placement based on the original image width

    # Dynamically adjust the window size based on image dimensions
    window_width = globalSourceImage.width * 2 + 30  # Space for both images and some padding
    window_height = max(globalSourceImage.height, 250) + 250  # Space for controls above and the tallest image
    window.geometry(f"{window_width}x{window_height}")


def countBlack (image):
    blacks = 0
    for x in range(image.width):
        for y in range(image.height):
            if image.getpixel((x,y)) < 0.0001:
                blacks += 1
    return blacks

def printThreadFcn(evt):
    global print_cancel_flag
    while True:
        if evt.wait(1):
            slices = image_slicer_and_scaler(globalDisplayImage)
            with ThermalPrinter(port='COM6', heat_time=110) as printer:
                for slice in slices:
                    # Check the print cancel flag here.
                    if print_cancel_flag:
                        print("Print canceled.")
                        evt.clear()
                        print_cancel_flag = False  # Reset the flag for future use.
                        printer.feed(2)
                        return  # Exit the thread function to stop printing.
                    printer.image(slice)
                    ratio = countBlack(slice) / (384 * 20)
                    ratio = 1.3 * max(0, ratio - 0.5)
                    time.sleep(ratio)
                printer.feed(2)
            evt.clear()

def cancelPrint():
    global print_cancel_flag
    print_cancel_flag = True  # Set the flag to cancel the print operation.

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

def openImage():
    global globalSourceImage
    global globalDisplayImage
    imageFileName = filedialog.askopenfilename()
    try:
        globalSourceImage = Image.open(imageFileName)
        process_image()
    except Exception as e:
        messagebox.showerror("Error", f"Can not open file, check if valid image was chosen. {e}")
    repaintImages()

# Make sure to include the call to handle_clipboard_image() in your GUI event loop.
# For example, bind it to a button or a menu option.

window = tk.Tk()
window.configure(background="#3e5757")
window.geometry(str(MIN_WINDOW_WIDTH) + "x400")  # Start with a smaller window that expands as needed
# Setting the window title
window.title("Thermal print3r Tool")
window.iconbitmap('icon.ico')

# Bind Ctrl+V to handle_clipboard_image
window.bind("<Control-v>", lambda event: handle_clipboard_image())

#lbl_title = tk.Label(text="Thermal Printer Tool", fg="#b9bcba", bg="#3e5757", width=40, height=3, font=("Arial", 25))
#lbl_title.place(x=0, y=0)

btn_openImage = tk.Button(    text="open image...",    width=13,    height=2,    bg="#5c7474",    fg="#b9bcba",    font=("Arial", 10), command=openImage)
btn_openImage.place(x=10, y=10)

btn_print = tk.Button(    text="print",    width=13,    height=2,    bg="#5c7474",    fg="#b9bcba",    font=("Arial", 10), command=printImage)
btn_print.place(x=130, y=10)

btn_cancelPrint = tk.Button(text="Cancel Print", width=13, height=2, bg="#5c7474", fg="#b9bcba", font=("Arial", 10), command=cancelPrint)
btn_cancelPrint.place(x=250, y=10)  # Adjust the placement according to your layout.

btn_incCo = tk.Button(    text="+C",    width=5,    height=1,    bg="#5c7474",    fg="#b9bcba",    font=("Arial", 10), command = incCo)
btn_incCo.place(x=10, y=60)

btn_decCo = tk.Button(    text="-C",    width=5,    height=1,    bg="#5c7474",    fg="#b9bcba",    font=("Arial", 10), command = decCo)
btn_decCo.place(x=10, y=95)

btn_incBri = tk.Button(    text="+B",    width=5,    height=1,    bg="#5c7474",    fg="#b9bcba",    font=("Arial", 10), command = incBri)
btn_incBri.place(x=70, y=60)

btn_decBri = tk.Button(    text="-B",    width=5,    height=1,    bg="#5c7474",    fg="#b9bcba",    font=("Arial", 10), command = decBri)
btn_decBri.place(x=70, y=95)

btn_reset = tk.Button(    text="reset",    width=8,    height=2,    bg="#5c7474",    fg="#b9bcba",    font=("Arial", 10), command = resetCoBri)
btn_reset.place(x=130, y=70)

print_event = threading.Event()
printerThread = threading.Thread(target=printThreadFcn,args=[print_event], daemon=True).start()

window.mainloop()

