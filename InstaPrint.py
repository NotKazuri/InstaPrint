import tkinter as tk
from tkinter import filedialog, messagebox, Canvas, Button, PhotoImage, ttk
from pathlib import Path
import os
import subprocess
from PyPDF2 import PdfReader, PdfWriter
import win32api
import win32print
import win32con
import serial
import time
from PIL import Image, ImageTk
import psutil
import urllib.parse
import win32com.client
import threading
import socket
import pyqrcode
from flask import Flask, request, render_template_string
import fitz  # PyMuPDF for PDF rendering

# ===========================================================================
# == ASSET & UTILITY FUNCTIONS
# ===========================================================================
# Directory for the PNGs and other assets
OUTPUT_PATH = Path(__file__).parent
ASSETS_PATH = Path(r"C:\InstaPrint Machine\build\assets")  # Adjust as needed

def relative_to_assets(frame: str, path: str) -> str:
    return str(ASSETS_PATH / frame / Path(path))

def center_window(window, width, height):
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x = (screen_width // 2) - (width // 2)
    y = (screen_height // 2) - (height // 2)
    window.geometry(f"{width}x{height}+{x}+{y}")

def set_printer_color(printer_name, color_choice):
    hPrinter = None
    try:
        hPrinter = win32print.OpenPrinter(printer_name)
        properties = win32print.GetPrinter(hPrinter, 2)
        devmode = properties["pDevMode"]
        if hasattr(devmode, "dmFields"):
            devmode.dmFields |= win32con.DM_COLOR
        else:
            print("dmFields attribute not found; proceeding without updating dmFields.")
        if color_choice == "Black and White":
            devmode.dmColor = 1  # Black and White
        else:
            devmode.dmColor = 2  # Colored
        win32print.DocumentProperties(None, hPrinter, printer_name, devmode, devmode, win32print.DM_IN_BUFFER | win32print.DM_OUT_BUFFER)
    except Exception as e:
        print("Error setting printer color:", e)
    finally:
        if hPrinter:
            win32print.ClosePrinter(hPrinter)

# ===========================================================================
# == FLASK SERVER FOR WI‑FI FILE UPLOAD
# ===========================================================================
# The Flask app runs in a separate thread to allow Wi‑Fi uploads.
flask_app = Flask(__name__)

# Folder to store uploaded files via Wi‑Fi (use the same folder as used by the machine)
WIFI_UPLOAD_FOLDER = r"C:\INSTAPRINTMACHINE"
os.makedirs(WIFI_UPLOAD_FOLDER, exist_ok=True)

# Folder for storing static files (such as the QR code image)
WIFI_STATIC_FOLDER = 'wifi_static'
os.makedirs(WIFI_STATIC_FOLDER, exist_ok=True)

def get_local_ip():
    """Return the local network IP address of this machine."""
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(("10.255.255.255", 1))
        ip = s.getsockname()[0]
    except Exception:
        ip = "127.0.0.1"
    finally:
        s.close()
    return ip

def get_scheme_and_ssl_context():
    """Return the URL scheme and SSL context if certificate files exist."""
    ssl_cert = 'cert.pem'
    ssl_key = 'key.pem'
    if os.path.exists(ssl_cert) and os.path.exists(ssl_key):
        print("SSL certificate files found. Running in HTTPS mode.")
        return "https", (ssl_cert, ssl_key)
    else:
        print("SSL certificate files not found. Running in HTTP mode.")
        return "http", None

SCHEME, SSL_CONTEXT = get_scheme_and_ssl_context()

@flask_app.route('/')
def wifi_index():
    local_ip = get_local_ip()
    upload_url = f"{SCHEME}://{local_ip}:5000/upload"
    # Generate the QR code for the upload URL
    qr = pyqrcode.create(upload_url)
    qr_path = os.path.join(WIFI_STATIC_FOLDER, 'qr.png')
    qr.png(qr_path, scale=8)
    html = f'''
    <!doctype html>
    <html>
      <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <title>Send File via Wi-Fi</title>
        <style>
          body {{
            background: linear-gradient(135deg, #1e3c72, #2a5298);
            color: white;
            font-family: Arial, sans-serif;
            text-align: center;
            margin: 0;
            padding: 0;
            display: flex;
            flex-direction: column;
            justify-content: center;
            align-items: center;
            height: 100vh;
            font-size: 18px;
          }}
          h1 {{
            font-size: 2em;
            margin-bottom: 20px;
          }}
          a {{
            color: #ffffff;
            text-decoration: underline;
            font-size: 1.2em;
          }}
          img {{
            max-width: 80%;
            height: auto;
          }}
        </style>
      </head>
      <body>
        <div class="container">
          <h1>Send File via Wi-Fi</h1>
          <p>Scan the QR code below on your mobile device to upload a file:</p>
          <img src="/{qr_path}" alt="QR Code">
          <p>Or click <a href="{upload_url}">{upload_url}</a></p>
        </div>
      </body>
    </html>
    '''
    return render_template_string(html)


@flask_app.route('/upload', methods=['GET', 'POST'])
def wifi_upload():
    if request.method == 'POST':
        if 'file' not in request.files:
            return 'No file part in the request.'
        file = request.files['file']
        if file.filename == '':
            return 'No file selected.'
        filepath = os.path.join(WIFI_UPLOAD_FOLDER, file.filename)
        file.save(filepath)
        # Note: This Flask route is separate from the Tkinter app.
        return f'File "{file.filename}" uploaded successfully to "C:\\INSTAPRINTMACHINE"!'
    return '''
    <!doctype html>
    <html>
      <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <title>Upload File</title>
        <style>
          body {
            background: linear-gradient(135deg, #1e3c72, #2a5298);
            color: white;
            font-family: Arial, sans-serif;
            text-align: center;
            padding-top: 50px;
            font-size: 18px;
          }
          input[type="file"] {
            margin: 20px 0;
            font-size: 1em;
          }
          input[type="submit"] {
            background-color: #ffffff;
            color: #1e3c72;
            border: none;
            padding: 10px 20px;
            font-size: 1.2em;
            cursor: pointer;
          }
        </style>
      </head>
      <body>
        <h1>Upload a File</h1>
        <form method="post" enctype="multipart/form-data">
          <input type="file" name="file"><br>
          <input type="submit" value="Upload">
        </form>
      </body>
    </html>
    '''

def run_flask_server():
    flask_app.run(host='0.0.0.0', port=5000, debug=False, use_reloader=False, ssl_context=SSL_CONTEXT)

# ===========================================================================
# == INSTANT PRINT MACHINE APPLICATION
# ===========================================================================
class InstaPrintApp:
    def __init__(self, root):
        self.root = root
        self.root.title("InstaPrint Machine")
        self.screen_width = self.root.winfo_screenwidth()
        self.screen_height = self.root.winfo_screenheight()
        self.root.geometry(f"{self.screen_width}x{self.screen_height}")
        self.root.configure(bg="#FFFFFF")
        self.root.attributes('-fullscreen', True)
        self.root.resizable(False, False)

        # Design dimensions and scaling
        design_width = 1920
        design_height = 1080
        self.scale_x = self.screen_width / design_width
        self.scale_y = self.screen_height / design_height
        self.scale = lambda x, y: (x * self.scale_x, y * self.scale_y)

        # Slide control:
        # 0 = Intro, 1 = File Sending, 2 = File Confirmation,
        # 3 = Paper Selection, 4 = Summary Table, 5 = Exit Screen
        self.current_slide = 0

        # Application variables
        self.file_path = None
        self.selected_size = None
        self.selected_color = None
        self.selected_copies = None
        self.paper_size_cost = 0
        self.paper_color_cost = 0
        self.image_refs = []
        self.preview_images = []  # To hold PDF preview images

        # For multiple copies (Spinbox)
        self.copies_spinbox = tk.Spinbox(self.root, from_=1, to=5, font=("Tahoma", int(20 * self.scale_y)))
        self.copies_spinbox.place_forget()

        self.paper_size_buttons = {}
        self.paper_color_buttons = {}
        self.copies_buttons = {}

        # USB file selection flag
        self.usb_dialog_opened = False

        # Save reference for the Send File button (Frame 1)
        self.send_file_btn = None

        # Arduino (Coin Slot) variables and serial connection
        self.arduino_balance = 0.0
        self.ser = None
        self.init_serial_connection()
        self.read_serial()

        # Main canvas for drawing slides
        self.canvas = Canvas(root, bg="#FFFFFF", height=self.screen_height, width=self.screen_width, bd=0, highlightthickness=0, relief="ridge")
        self.canvas.pack(fill=tk.BOTH, expand=True)

        # Load logo if available
        self.logo_path = "INSTAPRINT MAIN LOGO.png"
        try:
            self.logo_tk = self.create_logo(self.logo_path, size=int(700 * self.scale_x))
        except Exception as e:
            print(f"Error loading logo: {e}")
            self.logo_tk = None

        # Start the Wi‑Fi upload server (Flask) in a separate thread
        self.start_wifi_server()

        # Draw the initial slide
        self.draw_slide()

        # Bind Escape key to exit fullscreen/close window
        self.root.bind("<Escape>", lambda e: self.root.destroy())

    # ----------------------------
    # Arduino & Serial Functions
    # ----------------------------
    def init_serial_connection(self):
        try:
            self.ser = serial.Serial('COM3', 9600, timeout=1)
            time.sleep(2)
            print("Serial connection established at 9600 baud.")
        except Exception as e:
            messagebox.showerror("Serial Error", f"Failed to open serial port: {e}")
            self.ser = None

    def read_serial(self):
        if self.ser is not None:
            try:
                if self.ser.in_waiting:
                    line = self.ser.readline().decode('utf-8', errors='ignore').strip()
                    print("Serial read:", line)
                    if line.startswith("Total Balance:"):
                        parts = line.split()
                        if len(parts) >= 3:
                            try:
                                self.arduino_balance = float(parts[2])
                                self.update_payment()
                            except ValueError:
                                print("Error parsing coin amount")
            except Exception as e:
                print("Serial read error:", e)
        self.root.after(500, self.read_serial)

    def update_payment(self):
        total_cost = self.calculate_cost()
        print(f"Calculated Total Cost: ₱{total_cost:.2f} (Pages: {self.get_page_count()}, Size Cost: {self.paper_size_cost}, Color Cost: {self.paper_color_cost}, Arduino Balance: ₱{self.arduino_balance:.2f})")
        if self.current_slide == 4:
            self.draw_slide()

    def reset_arduino(self):
        if self.ser is not None:
            try:
                self.ser.reset_input_buffer()
                self.ser.reset_output_buffer()
                self.ser.setDTR(False)
                time.sleep(0.1)
                self.ser.setDTR(True)
                print("Arduino reset via DTR toggle.")
            except Exception as e:
                print("Error resetting Arduino via DTR:", e)
        self.arduino_balance = 0.0

    # ----------------------------
    # File & Print Functions
    # ----------------------------
    def convert_pdf_to_grayscale(self, input_pdf):
        gray_file = os.path.splitext(input_pdf)[0] + "_grayscale.pdf"
        gs_command = [
            "C:/InstaPrint Machine/gs10.05.0/bin/gswin64c.exe",
            "-sDEVICE=pdfwrite",
            "-dCompatibilityLevel=1.4",
            "-dPDFSETTINGS=/prepress",
            "-dColorConversionStrategy=/Gray",
            "-dProcessColorModel=/DeviceGray",
            "-dNOPAUSE",
            "-dBATCH",
            f"-sOutputFile={gray_file}",
            input_pdf
        ]
        try:
            subprocess.run(gs_command, check=True)
            print(f"Converted {input_pdf} to grayscale successfully as {gray_file}.")
            return gray_file
        except subprocess.CalledProcessError as e:
            print("Error converting PDF to grayscale:", e)
            return input_pdf

    def convert_pdf_to_grayscale_with_progress(self, input_pdf):
        # NEW FEATURE: Display a progress bar during conversion
        progress_window = tk.Toplevel(self.root)
        progress_window.title("Converting to Black and White")
        progress_label = tk.Label(progress_window, text="Converting file... Please wait.")
        progress_label.pack(padx=20, pady=10)
        progress_var = tk.DoubleVar(value=0)
        progress_bar = ttk.Progressbar(progress_window, variable=progress_var, maximum=100, length=300)
        progress_bar.pack(padx=20, pady=10)
        center_window(progress_window, 400, 150)

        conversion_done = threading.Event()
        result_container = {}
        def conversion_task():
            new_file = self.convert_pdf_to_grayscale(input_pdf)
            result_container['result'] = new_file
            conversion_done.set()

        thread = threading.Thread(target=conversion_task)
        thread.start()

        def update_progress():
            if conversion_done.is_set():
                progress_var.set(100)
                progress_window.update_idletasks()
                progress_window.after(500, progress_window.destroy)
                self.file_path = result_container.get('result', input_pdf)
            else:
                current = progress_var.get()
                if current < 95:
                    progress_var.set(current + 5)
                progress_window.after(200, update_progress)
        update_progress()
        progress_window.grab_set()

    def select_paper_color(self, color, cost):
        self.selected_color = color
        self.paper_color_cost = cost
        for key, data in self.paper_color_buttons.items():
            btn = data["button"]
            if key == color:
                btn.config(image=data["selected_img"])
            else:
                btn.config(image=data["normal_img"])
        print(f"Selected Paper Color: {color} - Cost: {cost} peso(s)")
        if color == "Black and White" and self.file_path and self.file_path.lower().endswith(".pdf"):
            # Use the new conversion method with a progress bar
            self.convert_pdf_to_grayscale_with_progress(self.file_path)

    def create_logo(self, image_path, size=400):
        img = Image.open(image_path).convert("RGBA").resize((size, size), Image.LANCZOS)
        return ImageTk.PhotoImage(img)

    def get_page_count(self):
        if self.file_path and self.file_path.lower().endswith(".pdf"):
            try:
                with open(self.file_path, "rb") as f:
                    reader = PdfReader(f)
                    return len(reader.pages)
            except Exception as e:
                print("PDF error:", e)
                return 1
        return 1

    def calculate_cost(self):
        page_count = self.get_page_count()
        if self.selected_copies == "Multiple Copies":
            try:
                copies = int(self.copies_spinbox.get())
            except Exception as e:
                print("Error getting spinbox value:", e)
                copies = 1
        elif self.selected_copies == "One Copy":
            copies = 1
        else:
            copies = 1
        total = (self.paper_size_cost + self.paper_color_cost) * page_count * copies
        return total

    def cleanup_folder(self):
        if self.file_path:
            folder = r"C:\Instafiles"  # Change as appropriate
            try:
                for filename in os.listdir(folder):
                    file_path = os.path.join(folder, filename)
                    if os.path.isfile(file_path):
                        os.remove(file_path)
                print(f"Cleaned up folder: {folder}")
            except Exception as e:
                print(f"Error cleaning up folder: {e}")

    def execute_print_job(self):
        if self.selected_size and self.selected_size.lower() == "short":
            printer_name = "EPSON L120 Series"
        elif self.selected_size and self.selected_size.lower() == "long":
            printer_name = "Brother DCP-T426W"
        else:
            printer_name = "EPSON L120 Series"

        set_printer_color(printer_name, self.selected_color)

        try:
            win32print.SetDefaultPrinter(printer_name)

            # Determine number of copies
            if self.selected_copies == "Multiple Copies":
                copies = int(self.copies_spinbox.get())
            else:
                copies = 1

            # If the selected file is an image, convert it to a same-size PDF first
            path_to_print = self.file_path
            if path_to_print and path_to_print.lower().endswith((".png", ".jpg", ".jpeg")):
                path_to_print = self.convert_image_to_pdf(path_to_print)

            # Send the job to the printer
            for _ in range(copies):
                win32api.ShellExecute(0, "print", path_to_print, None, ".", 0)

        except Exception as e:
            messagebox.showerror("Printing Error", f"Failed to print document: {e}")

    def convert_image_to_pdf(self, image_path):
            pdf_path = os.path.splitext(image_path)[0] + ".pdf"
            try:
                img = Image.open(image_path)
                # flatten transparency if present
                if img.mode in ("RGBA", "LA"):
                    bg = Image.new("RGB", img.size, (255, 255, 255))
                    bg.paste(img, mask=img.split()[-1])
                    img = bg
                else:
                    img = img.convert("RGB")
                img.save(pdf_path, "PDF")
                print(f"Converted image {image_path} to PDF {pdf_path}")
                return pdf_path
            except Exception as e:
                print(f"Error converting image to PDF: {e}")
                return image_path

    def handle_print_and_next(self):
        self.execute_print_job()
        self.reset_arduino()
        self.current_slide = 5
        self.draw_slide()

    # ----------------------------
    # Slide Drawing & Navigation
    # ----------------------------
    def draw_slide(self, event=None):
        self.canvas.delete("all")
        self.image_refs = []
        if self.current_slide == 0:
            self.display_intro_slide()
        elif self.current_slide == 1:
            self.display_file_sending_slide()
        elif self.current_slide == 2:
            self.display_file_confirmation_slide()
        elif self.current_slide == 3:
            self.display_paper_selection_slide()
        elif self.current_slide == 4:
            self.display_summary_table()
        elif self.current_slide == 5:
            self.display_goodbye_screen()
        self.canvas.update_idletasks()

    def next_slide(self):
        if self.current_slide == 1 and self.file_path is None:
            messagebox.showwarning("No File Selected", "Please select a file before continuing.")
            return
        if self.current_slide == 3:
            if not self.selected_size or not self.selected_color or not self.selected_copies:
                messagebox.showwarning("Incomplete Selection", "Please select paper size, color, and copies before continuing.")
                return
            self.copies_spinbox.place_forget()
        self.current_slide += 1
        self.draw_slide()

    def previous_slide(self):
        if self.current_slide > 0:
            self.current_slide -= 1
            self.draw_slide()

    def restart_application(self):
        if self.ser:
            try:
                self.ser.reset_input_buffer()
                self.ser.reset_output_buffer()
                self.ser.setDTR(False)
                time.sleep(0.1)
                self.ser.setDTR(True)
                print("Arduino reset in restart_application.")
            except Exception as e:
                print("Error resetting Arduino:", e)
        self.cleanup_folder()
        self.file_path = None
        self.selected_size = None
        self.selected_color = None
        self.selected_copies = None
        self.paper_size_cost = 0
        self.paper_color_cost = 0
        self.copies_spinbox.place_forget()
        self.reset_arduino()
        self.current_slide = 0
        self.draw_slide()

    # ----------------------------
    # Frame 0: Introduction Slide
    def display_intro_slide(self):
        self.canvas.configure(bg="#FFFFFF")
        try:
            img = PhotoImage(file=relative_to_assets("frame0", "frame0_image_1.png"))
            x, y = self.scale(960, 540)
            self.canvas.create_image(x, y, image=img)
            self.image_refs.append(img)
        except Exception as e:
            print(f"Error loading frame0_image_1.png: {e}")
        try:
            img = Image.open(relative_to_assets("frame0", "frame0_image_2.png"))
            new_size = (int(900 * self.scale_x), int(900 * self.scale_y))
            img = img.resize(new_size, Image.LANCZOS)
            img = ImageTk.PhotoImage(img)
            x, y = self.scale(960, 500)
            self.canvas.create_image(x, y, image=img)
            self.image_refs.append(img)
        except Exception as e:
            print(f"Error loading frame0_image_2.png: {e}")
        try:
            btn_img = PhotoImage(file=relative_to_assets("frame0", "frame0_button_1.png"))
            continue_btn = Button(self.root, image=btn_img, borderwidth=0, highlightthickness=0, command=self.next_slide, relief="flat")
            x, y = self.scale(960, 890)
            self.canvas.create_window(x, y, window=continue_btn, width=int(327 * self.scale_x), height=int(82 * self.scale_y))
            self.image_refs.append(btn_img)
        except Exception as e:
            print(f"Error loading frame0_button_1.png: {e}")

    # ----------------------------
    # Frame 1: File Sending Slide
    def display_file_sending_slide(self):
        self.canvas.configure(bg="#FFFFFF")
        images = [
            ("frame1_image_1.png", (960, 540)),
            ("frame1_image_2.png", (960, 540)),
            ("frame1_image_3.png", (960, 180)),
            ("frame1_image_4.png", (960, 180)),
            ("frame1_image_5.png", (1366, 449)),
            ("frame1_image_6.png", (589, 449)),
            ("frame1_image_7.png", (85, 1000)),
            ("frame1_image_8.png", (963, 320)),
            ("frame1_image_9.png", (975, 1025))
        ]
        for img_name, pos in images:
            try:
                img = PhotoImage(file=relative_to_assets("frame1", img_name))
                x, y = self.scale(*pos)
                self.canvas.create_image(x, y, image=img)
                self.image_refs.append(img)
            except Exception as e:
                print(f"Error loading {img_name}: {e}")
        buttons = [
            ("frame1_button_1.png", (590, 677), self.handle_usb_transfer),
            ("frame1_button_2.png", (590, 570), self.handle_bluetooth),
            ("frame1_button_3.png", (960, 830), self.next_slide),
            ("frame1_button_4.png", (1370, 610), self.select_file)
        ]
        for img_name, pos, command in buttons:
            try:
                btn_img = PhotoImage(file=relative_to_assets("frame1", img_name))
                btn = Button(self.root, image=btn_img, borderwidth=0, highlightthickness=0, command=command, relief="flat")
                x, y = self.scale(*pos)
                self.canvas.create_window(x, y, window=btn)
                self.image_refs.append(btn_img)
                if command == self.select_file:
                    self.send_file_btn = btn
            except Exception as e:
                print(f"Error loading {img_name}: {e}")
        try:
            wifi_btn_img = PhotoImage(file=relative_to_assets("frame1", "frame1_button_5.png"))
            wifi_btn = Button(self.root, image=wifi_btn_img, borderwidth=0, highlightthickness=0, command=self.show_wifi_qr_popup, relief="flat")
            x, y = self.scale(590, 780)
            self.canvas.create_window(x, y, window=wifi_btn)
            self.image_refs.append(wifi_btn_img)
        except Exception as e:
            print("Error loading frame1_button_5.png:", e)

    # ----------------------------
    # Wi‑Fi QR Code Popup (Frame 1)
    def show_wifi_qr_popup(self):
        messagebox.showinfo("Wi‑Fi", "Please connect to neu-student")
        local_ip = get_local_ip()
        wifi_url = f"{SCHEME}://{local_ip}:5000/upload"
        try:
            qr = pyqrcode.create(wifi_url)
            qr_path = os.path.join(WIFI_STATIC_FOLDER, 'qr.png')
            qr.png(qr_path, scale=8)
        except Exception as e:
            messagebox.showerror("QR Code Error", f"Error generating QR code: {e}")
            return
        popup = tk.Toplevel(self.root)
        popup.title("Wi‑Fi File Upload")
        popup.configure(bg="#1e3c72")
        popup.geometry("300x400")
        try:
            qr_img = Image.open(qr_path)
            qr_photo = ImageTk.PhotoImage(qr_img)
        except Exception as e:
            messagebox.showerror("Image Error", f"Error loading QR image: {e}")
            popup.destroy()
            return
        lbl = tk.Label(popup, image=qr_photo, bg="#1e3c72")
        lbl.image = qr_photo
        lbl.pack(pady=20)
        url_label = tk.Label(popup, text=wifi_url, font=("Arial", 12), fg="white", bg="#1e3c72")
        url_label.pack(pady=10)
        close_btn = Button(popup, text="Close", font=("Arial", 14), command=popup.destroy, bg="#ffffff", fg="#1e3c72")
        close_btn.pack(pady=20)

    # ----------------------------
    # Frame 2: File Confirmation Slide
    def display_file_confirmation_slide(self):
        self.canvas.configure(bg="#FFFFFF")
        images = [
            ("frame2_image_1.png", (960, 540)),
            ("frame2_image_2.png", (85, 1000)),
            ("frame2_image_3.png", (960, 540)),
            ("frame2_image_4.png", (960, 205)),
            ("frame2_image_5.png", (960, 200))
        ]
        for img_name, pos in images:
            try:
                img = PhotoImage(file=relative_to_assets("frame2", img_name))
                x, y = self.scale(*pos)
                self.canvas.create_image(x, y, image=img)
                self.image_refs.append(img)
            except Exception as e:
                print(f"Error loading {img_name}: {e}")
        if self.file_path and self.file_path.lower().endswith(".pdf"):
            preview_frame = tk.Frame(self.canvas, bg="white")
            preview_width = 300
            preview_height = 450
            preview_canvas = tk.Canvas(preview_frame, bg="white", width=preview_width, height=preview_height)
            v_scrollbar = tk.Scrollbar(preview_frame, orient="vertical", command=preview_canvas.yview)
            preview_canvas.configure(yscrollcommand=v_scrollbar.set)
            v_scrollbar.pack(side="right", fill="y")
            preview_canvas.pack(side="left", fill="both", expand=True)
            inner_frame = tk.Frame(preview_canvas, bg="white")
            preview_canvas.create_window((0, 0), window=inner_frame, anchor="nw", width=preview_width)
            def on_configure(event):
                preview_canvas.configure(scrollregion=preview_canvas.bbox("all"))
            inner_frame.bind("<Configure>", on_configure)
            try:
                doc = fitz.open(self.file_path)
                page_images = []
                for page in doc:
                    pix = page.get_pixmap()
                    img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
                    img.thumbnail((preview_width, preview_height))
                    img_tk = ImageTk.PhotoImage(img)
                    page_label = tk.Label(inner_frame, image=img_tk, bg="white")
                    page_label.image = img_tk
                    page_label.pack(pady=10)
                    page_images.append(img_tk)
                self.preview_images = page_images
            except Exception as e:
                tk.Label(inner_frame, text="Error loading PDF preview", bg="white", fg="black").pack()
            x, y = self.scale(960, 500)
            self.canvas.create_window(x, y, window=preview_frame)
            total_pages = self.get_page_count()
            page_sel_frame = tk.Frame(self.canvas, bg="white")
            tk.Label(page_sel_frame, text="Print Pages From:", bg="white", font=("Tahoma", int(20 * self.scale_y))).pack(side="left", padx=5)
            from_var = tk.IntVar(value=1)
            self.print_from_spinbox = tk.Spinbox(page_sel_frame, from_=1, to=total_pages, width=5, font=("Tahoma", int(20 * self.scale_y)), textvariable=from_var)
            self.print_from_spinbox.pack(side="left", padx=5)
            tk.Label(page_sel_frame, text="To:", bg="white", font=("Tahoma", int(20 * self.scale_y))).pack(side="left", padx=5)
            to_var = tk.IntVar(value=total_pages)
            self.print_to_spinbox = tk.Spinbox(page_sel_frame, from_=1, to=total_pages, width=5, font=("Tahoma", int(20 * self.scale_y)), textvariable=to_var)
            self.print_to_spinbox.pack(side="left", padx=5)
            x_sel, y_sel = self.scale(960, 950)
            self.canvas.create_window(x_sel, y_sel, window=page_sel_frame)
        else:
            x, y = self.scale(960, 500)
            self.canvas.create_text(x, y, text="Preview not available for this file type.", font=("Tahoma", int(40 * self.scale_y), "bold"), fill="black")    
        buttons = [
            ("frame2_button_1.png", (1220, 800), self.next_slide),
            ("frame2_button_2.png", (750, 800), self.previous_slide)
        ]
        for img_name, pos, command in buttons:
            try:
                btn_img = PhotoImage(file=relative_to_assets("frame2", img_name))
                btn = Button(self.root, image=btn_img, borderwidth=0, highlightthickness=0, command=command, relief="flat")
                x, y = self.scale(*pos)
                self.canvas.create_window(x, y, window=btn)
                self.image_refs.append(btn_img)
            except Exception as e:
                print(f"Error loading {img_name}: {e}")

    # ----------------------------
    # Frame 3: Paper Selection Slide
    def get_enlarged_image(self, path, factor):
        try:
            img = Image.open(path)
            width, height = img.size
            new_size = (int(width * factor), int(height * factor))
            img = img.resize(new_size, Image.LANCZOS)
            return ImageTk.PhotoImage(img)
        except Exception as e:
            print(f"Error enlarging image: {e}")
            return None

    def display_paper_selection_slide(self):
        self.canvas.configure(bg="#FFFFFF")
        images = [
            ("frame3_image_1.png", (960, 540)),
            ("frame3_image_2.png", (960, 540)),
            ("frame3_image_3.png", (85, 1000)),
            ("frame3_image_4.png", (1341, 334)),
            ("frame3_image_5.png", (960, 701)),
            ("frame3_image_6.png", (544, 334)),
            ("frame3_image_7.png", (960, 166)),
            ("frame3_image_8.png", (960, 164)),
            ("frame3_image_9.png", (960, 701)),
            ("frame3_image_10.png", (544, 331)),
            ("frame3_image_11.png", (1348, 332))
        ]
        for img_name, pos in images:
            try:
                img = PhotoImage(file=relative_to_assets("frame3", img_name))
                x, y = self.scale(*pos)
                self.canvas.create_image(x, y, image=img)
                self.image_refs.append(img)
            except Exception as e:
                print(f"Error loading {img_name}: {e}")
        self.paper_size_buttons = {}
        paper_size_buttons_data = [
            ("frame3_button_1.png", (1340, 460), "Short", 1),
            ("frame3_button_6.png", (1340, 600), "Long", 1)
        ]
        for img_name, pos, size, cost in paper_size_buttons_data:
            try:
                img_path = relative_to_assets("frame3", img_name)
                normal_img = PhotoImage(file=img_path)
                selected_img = self.get_enlarged_image(img_path, 1.2)
                btn = Button(self.root, image=normal_img, borderwidth=0, highlightthickness=0, command=lambda s=size, c=cost: self.select_paper_size(s, c), relief="flat")
                self.paper_size_buttons[size] = {
                    "button": btn,
                    "normal_img": normal_img,
                    "selected_img": selected_img,
                    "pos": pos
                }
                x, y = self.scale(*pos)
                self.canvas.create_window(x, y, window=btn)
                self.image_refs.append(normal_img)
                self.image_refs.append(selected_img)
            except Exception as e:
                print(f"Error loading {img_name}: {e}")
        self.paper_color_buttons = {}
        paper_color_buttons_data = [
            ("frame3_button_2.png", (550, 460), "Black and White", 1),
            ("frame3_button_3.png", (550, 600), "Colored", 3)
        ]
        for img_name, pos, color, cost in paper_color_buttons_data:
            try:
                img_path = relative_to_assets("frame3", img_name)
                normal_img = PhotoImage(file=img_path)
                selected_img = self.get_enlarged_image(img_path, 1.2)
                btn = Button(self.root, image=normal_img, borderwidth=0, highlightthickness=0, command=lambda col=color, c=cost: self.select_paper_color(col, c), relief="flat")
                self.paper_color_buttons[color] = {
                    "button": btn,
                    "normal_img": normal_img,
                    "selected_img": selected_img,
                    "pos": pos
                }
                x, y = self.scale(*pos)
                self.canvas.create_window(x, y, window=btn)
                self.image_refs.append(normal_img)
                self.image_refs.append(selected_img)
            except Exception as e:
                print(f"Error loading {img_name}: {e}")
        self.copies_buttons = {}
        copies_buttons_data = [
            ("frame3_button_4.png", (775, 815), "One Copy", None),
            ("frame3_button_5.png", (1150, 812), "Multiple Copies", None)
        ]
        for img_name, pos, option, _ in copies_buttons_data:
            try:
                img_path = relative_to_assets("frame3", img_name)
                normal_img = PhotoImage(file=img_path)
                selected_img = self.get_enlarged_image(img_path, 1.2)
                btn = Button(self.root, image=normal_img, borderwidth=0, highlightthickness=0, command=lambda opt=option: self.select_copies(opt), relief="flat")
                self.copies_buttons[option] = {
                    "button": btn,
                    "normal_img": normal_img,
                    "selected_img": selected_img,
                    "pos": pos
                }
                x, y = self.scale(*pos)
                self.canvas.create_window(x, y, window=btn)
                self.image_refs.append(normal_img)
                self.image_refs.append(selected_img)
            except Exception as e:
                print(f"Error loading {img_name}: {e}")
        if self.selected_copies == "Multiple Copies":
            try:
                x, y = self.scale(928, 875)
                self.copies_spinbox.place(x=x, y=y, width=int(80 * self.scale_x), height=int(50 * self.scale_y))
                self.copies_spinbox.lift()
                self.root.update_idletasks()
                self.root.update()
            except Exception as e:
                print(f"Error updating spinbox position: {e}")
        else:
            self.copies_spinbox.place_forget()
        nav_buttons = [
            ("frame3_button_7.png", (1580, 875), self.next_slide),
            ("frame3_button_8.png", (340, 875), self.previous_slide)
        ]
        for img_name, pos, command in nav_buttons:
            try:
                btn_img = PhotoImage(file=relative_to_assets("frame3", img_name))
                btn = Button(self.root, image=btn_img, borderwidth=0, highlightthickness=0, command=command, relief="flat")
                x, y = self.scale(*pos)
                self.canvas.create_window(x, y, window=btn)
                self.image_refs.append(btn_img)
            except Exception as e:
                print(f"Error loading {img_name}: {e}")

    def select_paper_size(self, size, cost):
        self.selected_size = size
        self.paper_size_cost = cost
        for s, data in self.paper_size_buttons.items():
            btn = data["button"]
            if s == size:
                btn.config(image=data["selected_img"])
            else:
                btn.config(image=data["normal_img"])
        print(f"Selected Paper Size: {size} - Cost: {cost} peso(s)")

    def select_copies(self, option):
        self.selected_copies = option
        for opt, data in self.copies_buttons.items():
            btn = data["button"]
            if opt == option:
                btn.config(image=data["selected_img"])
            else:
                btn.config(image=data["normal_img"])
        print(f"Selected Copies: {option}")
        if option == "Multiple Copies":
            x, y = self.scale(928, 875)
            self.copies_spinbox.place(x=x, y=y, width=int(80 * self.scale_x), height=int(50 * self.scale_y))
            self.copies_spinbox.lift()
            self.root.update_idletasks()
            self.root.update()
        else:
            self.copies_spinbox.place_forget()

    # ----------------------------
    # Frame 4: Summary Table Slide
    def display_summary_table(self):
        self.canvas.configure(bg="#FFFFFF")
        total_cost = self.calculate_cost()
        net_cost = max(0, total_cost - self.arduino_balance)

        #net_cost = 0 (put the '#' for debugging)

        try:
            img = PhotoImage(file=relative_to_assets("frame4", "frame4_image_1.png"))
            x, y = self.scale(960, 540)
            self.canvas.create_image(x, y, image=img)
            self.image_refs.append(img)
            img = PhotoImage(file=relative_to_assets("frame4", "frame4_image_3.png"))
            x, y = self.scale(960, 160)
            self.canvas.create_image(x, y, image=img)
            self.image_refs.append(img)
        except Exception as e:
            print(f"Error loading frame4 images: {e}")
        summary_table = ttk.Treeview(self.root, columns=("SETTINGS", "DETAILS"), show="headings", height=5)
        summary_table.heading("SETTINGS", text="SETTINGS", anchor=tk.CENTER)
        summary_table.heading("DETAILS", text="DETAILS", anchor=tk.CENTER)
        summary_table.column("SETTINGS", width=int(400 * self.scale_x), anchor=tk.CENTER)
        summary_table.column("DETAILS", width=int(400 * self.scale_x), anchor=tk.CENTER)
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("Treeview", font=("Tahoma", int(20 * self.scale_y)), rowheight=int(65 * self.scale_y), background="#0d084d", foreground="white")
        style.configure("Treeview.Heading", font=("Tahoma", int(25 * self.scale_y), "bold"), background="#0099ff", foreground="white")
        style.map("Treeview", background=[("selected", "#0d084d")])
        summary_table.insert("", "end", values=("FILE", os.path.basename(self.file_path) if self.file_path else "No file selected"))
        summary_table.insert("", "end", values=("SIZE", self.selected_size if self.selected_size else "Not selected"))
        summary_table.insert("", "end", values=("COLOR", self.selected_color if self.selected_color else "Not selected"))
        if self.selected_copies == "Multiple Copies":
            try:
                copies_value = int(self.copies_spinbox.get())
            except:
                copies_value = "Not selected"
        elif self.selected_copies == "One Copy":
            copies_value = 1
        else:
            copies_value = self.selected_copies if self.selected_copies else "Not selected"
        summary_table.insert("", "end", values=("NUMBER OF COPIES", copies_value))
        summary_table.insert("", "end", values=("TOTAL PAGES", self.get_page_count() * (copies_value if isinstance(copies_value, int) else 1)))
        x, y = self.scale(960, 420)
        self.canvas.create_window(x, y, window=summary_table)
        amount_table = ttk.Treeview(self.root, columns=("Amount",), show="headings", height=1)
        amount_table.heading("Amount", text="AMOUNT TO BE PAID", anchor=tk.CENTER)
        amount_table.column("Amount", width=int(800 * self.scale_x), anchor=tk.CENTER)
        amount_table.insert("", "end", values=(f"₱{net_cost:.2f}",))
        x, y = self.scale(960, 690)
        self.canvas.create_window(x, y, window=amount_table)
       
        #net_cost = 0
       
        if net_cost == 0:
            try:
                btn_img = PhotoImage(file=relative_to_assets("frame4", "frame4_button_2(second).png"))
                print_btn = Button(self.root, image=btn_img, borderwidth=0, highlightthickness=0, command=self.handle_print_and_next, relief="flat")
                x, y = self.scale(960, 861)
                self.canvas.create_window(x, y, window=print_btn, width=int(323 * self.scale_x), height=int(114 * self.scale_y))
                self.image_refs.append(btn_img)
            except Exception as e:
                print(f"Error loading frame4_button_2(second).png: {e}")

    # ----------------------------
    # Frame 5: Exit Screen Slide
    def display_goodbye_screen(self):
        self.canvas.configure(bg="#FFFFFF")
        try:
            img = PhotoImage(file=relative_to_assets("frame5", "frame5_image_1.png"))
            x, y = self.scale(960, 540)
            self.canvas.create_image(x, y, image=img)
            self.image_refs.append(img)
        except Exception as e:
            print(f"Error loading frame5_image_1.png: {e}")
        images = [
            ("frame5_image_2.png", (959, 423)),
            ("frame5_image_3.png", (960, 702)),
            ("frame5_image_4.png", (959, 787))
        ]
        for img_name, pos in images:
            try:
                img = PhotoImage(file=relative_to_assets("frame5", img_name))
                x, y = self.scale(*pos)
                self.canvas.create_image(x, y, image=img)
                self.image_refs.append(img)
            except Exception as e:
                print(f"Error loading {img_name}: {e}")
        try:
            btn_img = PhotoImage(file=relative_to_assets("frame5", "frame5_button_1.png"))
            restart_btn = Button(self.root, image=btn_img, borderwidth=0, highlightthickness=0, command=self.restart_application, relief="flat")
            x, y = self.scale(960, 920)
            self.canvas.create_window(x, y, window=restart_btn, width=int(324 * self.scale_x), height=int(117 * self.scale_y))
            self.image_refs.append(btn_img)
        except Exception as e:
            print(f"Error loading frame5_button_1.png: {e}")

    # ----------------------------
    # File Selection & USB/Bluetooth Transfer
    def select_file(self):
        custom_folder = r"C:\INSTAPRINTMACHINE"  # Adjust as needed
        top = tk.Toplevel(self.root)
        top.title("Select a File")
        center_window(top, 600, 400)
        top.resizable(False, False)
        frame = tk.Frame(top)
        frame.pack(fill="both", expand=True, padx=10, pady=10)
        scrollbar = tk.Scrollbar(frame)
        scrollbar.pack(side="right", fill="y")
        listbox = tk.Listbox(frame, font=("Tahoma", int(16 * self.scale_y)), yscrollcommand=scrollbar.set)
        listbox.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=listbox.yview)
        try:
            allowed_exts = ('.pdf', '.docx', '.doc', '.xlsx', '.pptx', '.png', '.jpg', '.jpeg')
            files = [f for f in os.listdir(custom_folder) if os.path.isfile(os.path.join(custom_folder, f)) and f.lower().endswith(allowed_exts)]
            print("Files found:", files)
            for f in files:
                listbox.insert(tk.END, f)
        except Exception as e:
            print(f"Error accessing folder: {e}")
            top.destroy()
            return
       
                # --- Converts a PNG/JPEG into a same-size PDF ---

        def on_select():
            selected = listbox.curselection()
            if selected:
                file_name = listbox.get(selected[0])
                full_path = os.path.join(custom_folder, file_name)
                self.file_path = full_path
                print(f"File selected: {full_path}")
                top.destroy()
                try:
                    enlarged_img = self.get_enlarged_image(relative_to_assets("frame1", "frame1_button_4.png"), 1.2)
                    self.send_file_btn.config(image=enlarged_img)
                    self.send_file_btn.image = enlarged_img
                except Exception as e:
                    print("Error enlarging Send File button:", e)
            else:
                messagebox.showwarning("No Selection", "Please select a file.")

        def on_delete():
            selected = listbox.curselection()
            if selected:
                file_name = listbox.get(selected[0])
                full_path = os.path.join(custom_folder, file_name)
                answer = messagebox.askyesno("Delete Confirmation", f"Are you sure you want to delete '{file_name}'?")
                if answer:
                    try:
                        os.remove(full_path)
                        messagebox.showinfo("Delete", f"File '{file_name}' deleted successfully.")
                        listbox.delete(selected[0])
                    except Exception as e:
                        messagebox.showerror("Delete Error", f"Error deleting file '{file_name}': {e}")
            else:
                messagebox.showwarning("No Selection", "Please select a file to delete.")

        # Create a frame for the selection and delete buttons
        button_frame = tk.Frame(top)
        button_frame.pack(pady=10)
        select_btn = tk.Button(button_frame, text="Select File", font=("Tahoma", int(16 * self.scale_y)), command=on_select)
        select_btn.pack(side="left", padx=10)
        delete_btn = tk.Button(button_frame, text="Delete", font=("Tahoma", int(16 * self.scale_y)), command=on_delete)
        delete_btn.pack(side="left", padx=10)
        top.grab_set()

    def handle_usb_transfer(self):
        messagebox.showinfo("USB Transfer", "Please insert a USB drive.")
        self.usb_dialog_opened = False
        self.poll_for_usb()

    def poll_for_usb(self):
        usb_drives = []
        for p in psutil.disk_partitions(all=False):
            if 'removable' in p.opts.lower():
                usb_drives.append(p.device)
        if usb_drives:
            drive = usb_drives[0]
            if not drive.endswith("\\"):
                drive += "\\"
            if not os.path.exists(drive):
                print(f"Drive not found or not ready: {drive}")
                self.root.after(1000, self.poll_for_usb)
                return
            self.close_usb_explorer(drive)
            if not self.usb_dialog_opened:
                self.usb_dialog_opened = True
                self.usb_file_selection(drive)
        else:
            self.root.after(1000, self.poll_for_usb)

    def usb_file_selection(self, drive):
        top = tk.Toplevel(self.root)
        top.title("Select a File from USB")
        center_window(top, 600, 400)
        top.resizable(False, False)
        top.protocol("WM_DELETE_WINDOW", top.destroy)
        frame = tk.Frame(top)
        frame.pack(fill="both", expand=True, padx=10, pady=10)
        scrollbar = tk.Scrollbar(frame)
        scrollbar.pack(side="right", fill="y")
        listbox = tk.Listbox(frame, font=("Tahoma", 16), yscrollcommand=scrollbar.set)
        listbox.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=listbox.yview)
        allowed_exts = ('.pdf', '.docx', '.doc', '.xlsx', '.pptx', '.png', '.jpg')
        try:
            files = [f for f in os.listdir(drive) if os.path.isfile(os.path.join(drive, f)) and f.lower().endswith(allowed_exts)]
            print("Files found on USB:", files)
            if files:
                for f in files:
                    listbox.insert(tk.END, f)
        except Exception as e:
            print(f"Error accessing the USB drive: {e}")
            top.destroy()
            return

        def on_select():
            selected = listbox.curselection()
            if selected:
                file_name = listbox.get(selected[0])
                full_path = os.path.join(drive, file_name)
                self.file_path = full_path
                print(f"File selected from USB: {full_path}")
                top.destroy()
                messagebox.showinfo("File Received", "Your file has been successfully sent!")
            else:
                messagebox.showwarning("No Selection", "Please select a file.")

        select_btn = tk.Button(top, text="Select File", font=("Tahoma", 16), command=on_select)
        select_btn.pack(pady=10)
        top.grab_set()

    def close_usb_explorer(self, drive):
        try:
            shell = win32com.client.Dispatch("Shell.Application")
            for window in shell.Windows():
                try:
                    location = window.LocationURL
                    if location and location.startswith("file:/C:\\INSTAPRINTMACHINE"):
                        local_path = urllib.parse.unquote(location[8:]).replace('/', '\\')
                        if os.path.normcase(os.path.abspath(local_path)) == os.path.normcase(os.path.abspath(drive)):
                            window.Quit()
                except Exception:
                    continue
        except Exception as e:
            print("Error closing USB explorer windows:", e)

    def handle_bluetooth(self):
        try:
            subprocess.Popen(["C:\\INSTAPRINT MACHINE\\receive.bat"], shell=True)
        except Exception as e:
            messagebox.showerror("Bluetooth Error", f"Failed to run receive.bat: {str(e)}")

    def select_paper_size(self, size, cost):
        self.selected_size = size
        self.paper_size_cost = cost
        for s, data in self.paper_size_buttons.items():
            btn = data["button"]
            if s == size:
                btn.config(image=data["selected_img"])
            else:
                btn.config(image=data["normal_img"])
        print(f"Selected Paper Size: {size} - Cost: {cost} peso(s)")

    def start_wifi_server(self):
        flask_thread = threading.Thread(target=run_flask_server)
        flask_thread.daemon = True
        flask_thread.start()
        print("Wi‑Fi upload server started on port 5000.")

# ===========================================================================
# == MAIN EXECUTION
# ===========================================================================
if __name__ == "__main__":
    root = tk.Tk()
    app = InstaPrintApp(root)
    root.mainloop()
