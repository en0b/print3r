import time
import threading
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageEnhance, ImageTk, ImageGrab
from thermalprinter import ThermalPrinter
import serial.tools.list_ports

# Use the new resampling filter for Pillow (compatible with Pillow 10+)
RESAMPLE_FILTER = getattr(Image, 'Resampling', Image).LANCZOS

# Constants
WIDTH_PIXELS = 384  # Thermal printer width

class ThermalPrintTool:
    def __init__(self):
        # Image and adjustment state
        self.source_image = None
        self.display_image = None
        self.contrast = 0.0
        self.brightness = 0.0
        self.print_cancel_flag = False
        self.printer_port = None

        # Set up main window with a new dark color scheme
        self.window = tk.Tk()
        self.window.title("Thermal Print3r Tool")
        self.window.iconbitmap('icon.ico')
        self.window.configure(background="#2C3E50")  # Deep blue background

        # Top frame: Printer status and refresh button
        self.top_frame = tk.Frame(self.window, bg="#2C3E50")
        self.top_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=5)
        self.lbl_printer_status = tk.Label(
            self.top_frame,
            text="Printer: Detecting...",
            fg="#ECF0F1",  # light text
            bg="#2C3E50",
            font=("Arial", 12)
        )
        self.lbl_printer_status.pack(side=tk.LEFT)
        btn_refresh = tk.Button(
            self.top_frame,
            text="Refresh Printer",
            command=self.refresh_printer,
            bg="#1ABC9C",  # Teal button
            fg="#FFFFFF",
            font=("Arial", 10)
        )
        btn_refresh.pack(side=tk.RIGHT)

        # Left frame: Controls for file operations, printing, and adjustments
        self.left_frame = tk.Frame(self.window, bg="#34495E", relief=tk.RIDGE, borderwidth=2)
        self.left_frame.pack(side=tk.LEFT, fill=tk.Y, padx=10, pady=5)
        
        btn_open = tk.Button(
            self.left_frame,
            text="Open Image",
            command=self.open_image,
            bg="#1ABC9C",
            fg="#FFFFFF",
            font=("Arial", 10)
        )
        btn_open.pack(pady=5, fill=tk.X)
        
        btn_print = tk.Button(
            self.left_frame,
            text="Print",
            command=self.print_image,
            bg="#1ABC9C",
            fg="#FFFFFF",
            font=("Arial", 10)
        )
        btn_print.pack(pady=5, fill=tk.X)
        
        btn_cancel = tk.Button(
            self.left_frame,
            text="Cancel Print",
            command=self.cancel_print,
            bg="#1ABC9C",
            fg="#FFFFFF",
            font=("Arial", 10)
        )
        btn_cancel.pack(pady=5, fill=tk.X)
        
        btn_rotate = tk.Button(
            self.left_frame,
            text="Rotate Image",
            command=self.rotate_image,
            bg="#1ABC9C",
            fg="#FFFFFF",
            font=("Arial", 10)
        )
        btn_rotate.pack(pady=5, fill=tk.X)
        
        separator = tk.Frame(self.left_frame, height=2, bd=1, relief=tk.SUNKEN, bg="#ECF0F1")
        separator.pack(fill=tk.X, padx=5, pady=5)
        
        # Brightness slider
        lbl_brightness = tk.Label(
            self.left_frame,
            text="Brightness",
            fg="#ECF0F1",
            bg="#34495E",
            font=("Arial", 10)
        )
        lbl_brightness.pack(pady=(5, 0))
        self.scale_brightness = tk.Scale(
            self.left_frame,
            from_=-1.0,
            to=1.0,
            resolution=0.1,
            orient=tk.HORIZONTAL,
            length=150,
            bg="#34495E",
            fg="#ECF0F1",
            command=self.on_brightness_change
        )
        self.scale_brightness.set(0.0)
        self.scale_brightness.pack(pady=5)
        
        # Contrast slider
        lbl_contrast = tk.Label(
            self.left_frame,
            text="Contrast",
            fg="#ECF0F1",
            bg="#34495E",
            font=("Arial", 10)
        )
        lbl_contrast.pack(pady=(5, 0))
        self.scale_contrast = tk.Scale(
            self.left_frame,
            from_=-1.0,
            to=1.0,
            resolution=0.1,
            orient=tk.HORIZONTAL,
            length=150,
            bg="#34495E",
            fg="#ECF0F1",
            command=self.on_contrast_change
        )
        self.scale_contrast.set(0.0)
        self.scale_contrast.pack(pady=5)

        # Right frame: Image display (original and processed)
        self.right_frame = tk.Frame(self.window, bg="#34495E", relief=tk.RIDGE, borderwidth=2)
        self.right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        self.lbl_image_orig = tk.Label(self.right_frame, text="Original Image", bg="#34495E", fg="#ECF0F1")
        self.lbl_image_orig.grid(row=0, column=0, padx=5, pady=5)
        self.lbl_image_disp = tk.Label(self.right_frame, text="Processed Image", bg="#34495E", fg="#ECF0F1")
        self.lbl_image_disp.grid(row=0, column=1, padx=5, pady=5)
        
        # Bind Ctrl+V to paste image from clipboard
        self.window.bind("<Control-v>", lambda event: self.handle_clipboard_image())
        
        # Start printer thread for printing jobs
        self.print_event = threading.Event()
        self.printer_thread = threading.Thread(
            target=self.print_thread_function,
            args=(self.print_event,),
            daemon=True
        )
        self.printer_thread.start()
        
        # Initial printer detection
        self.refresh_printer()
        self.window.mainloop()

    def refresh_printer(self):
        """Detect printer by scanning COM ports."""
        com_port = self.detect_printer()
        if com_port:
            self.lbl_printer_status.config(text=f"Printer: {com_port} - Connected", fg="green")
            self.printer_port = com_port
        else:
            self.lbl_printer_status.config(text="Printer: Not Found", fg="red")
            self.printer_port = None

    def detect_printer(self):
        """Search through COM ports and try to identify the printer."""
        ports = list(serial.tools.list_ports.comports())
        for port in ports:
            try:
                with ThermalPrinter(port=port.device, heat_time=110) as printer:
                    return port.device
            except Exception:
                continue
        return None

    def handle_clipboard_image(self):
        """Grabs an image from the clipboard and processes it."""
        try:
            img = ImageGrab.grabclipboard()
            if img is None:
                messagebox.showerror("Error", "No image in clipboard")
                return
            self.source_image = img
            self.process_image()
            self.repaint_images()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to process clipboard image: {e}")

    def open_image(self):
        """Opens an image file using a file dialog."""
        filename = filedialog.askopenfilename()
        if filename:
            try:
                self.source_image = Image.open(filename)
                self.process_image()
                self.repaint_images()
            except Exception as e:
                messagebox.showerror("Error", f"Cannot open file, check if a valid image was chosen. {e}")

    def process_image(self):
        """Resizes the source image to exactly fill available printer width."""
        factor = WIDTH_PIXELS / self.source_image.width
        new_size = (WIDTH_PIXELS, round(self.source_image.height * factor))
        self.source_image = self.source_image.resize(new_size, RESAMPLE_FILTER)
        self.display_image = self.source_image
        # Ensure the image has a filename attribute for logging in ThermalPrinter.image()
        if not hasattr(self.display_image, 'filename'):
            self.display_image.filename = "processed.bmp"

    def repaint_images(self):
        """Applies brightness/contrast adjustments and updates the image displays."""
        if self.source_image is None:
            return
        adjusted = self.source_image.copy()
        adjusted = ImageEnhance.Brightness(adjusted).enhance(1 + self.brightness)
        adjusted = ImageEnhance.Contrast(adjusted).enhance(1 + self.contrast)
        self.display_image = adjusted.convert('1')
        # Ensure the image has a filename attribute for logging
        if not hasattr(self.display_image, 'filename'):
            self.display_image.filename = "processed.bmp"
        
        tmp_orig = ImageTk.PhotoImage(self.source_image)
        self.lbl_image_orig.config(image=tmp_orig)
        self.lbl_image_orig.image = tmp_orig
        
        tmp_disp = ImageTk.PhotoImage(self.display_image)
        self.lbl_image_disp.config(image=tmp_disp)
        self.lbl_image_disp.image = tmp_disp

    def on_brightness_change(self, value):
        try:
            self.brightness = float(value)
            self.repaint_images()
        except ValueError:
            pass

    def on_contrast_change(self, value):
        try:
            self.contrast = float(value)
            self.repaint_images()
        except ValueError:
            pass

    def rotate_image(self):
        """Rotates the source image by 90 degrees, then scales it to fill the available width."""
        if self.source_image is not None:
            self.source_image = self.source_image.rotate(90, expand=True)
            self.process_image()
            self.repaint_images()

    def print_thread_function(self, event):
        """Handles printing in a separate thread."""
        while True:
            if event.wait(1):
                if self.printer_port is None:
                    print("Printer not connected.")
                    event.clear()
                    continue
                try:
                    with ThermalPrinter(port=self.printer_port, heat_time=110) as printer:
                        # Print the processed image using the built-in image() method.
                        printer.image(self.display_image)
                        # Simulate feeding extra paper by printing 10 empty lines.
                        for _ in range(3):
                            printer.out("")
                except Exception as e:
                    print(f"Printing error: {e}")
                event.clear()

    def print_image(self):
        """Starts the print operation if an image is loaded."""
        if self.source_image is None:
            messagebox.showerror("Error", "No image loaded. Please load an image first.")
        elif self.print_event.is_set():
            messagebox.showerror("Error", "Another print is running. Please wait until it is completed.")
        else:
            self.print_event.set()

    def cancel_print(self):
        """Cancels the ongoing print operation."""
        self.print_cancel_flag = True

if __name__ == "__main__":
    ThermalPrintTool()
