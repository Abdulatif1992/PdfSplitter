import tkinter as tk
from tkinter import filedialog
from PIL import Image, ImageTk
import fitz  # PyMuPDF
import io
import os

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from pathlib import Path
import pytesseract

# Set the Tesseract executable path
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
# pytesseract.pytesseract.tesseract_cmd = r'C:\Users\sekine\AppData\Local\Programs\Tesseract-OCR\tesseract.exe'

class ImageSelectorApp:
    def __init__(self, master):
        self.frame1 = tk.Frame(master=master, width=150, height=100, border=1)
        self.frame1.pack(fill=tk.BOTH, side=tk.LEFT, expand=False)

        self.frame2 = tk.Frame(master=master, width=100)
        self.frame2.pack(fill=tk.BOTH, side=tk.LEFT, expand=True)

        self.master = master
        self.master.title("Safarov's PDF Splitter")
        self.canvas = tk.Canvas(self.frame2, width=1200, height=700)        
        self.canvas.pack()
        self.img = None
        self.image_path = None
        self.rect = None
        self.rect_coords = None

        self.btn_load = tk.Button(self.frame1, text="Load Pdf", command=self.load_image)
        self.btn_load.place(x=10, y=10) 

        self.btn_select = tk.Button(self.frame1, text="Select Special Part", command=self.select_special_part)
        self.btn_select.place(x=10, y=60) 

        # Create a Label widget
        self.count_label = tk.Label(self.frame1, text="__")
        self.count_label.place(x=10, y=150)

        self.line_label = tk.Label(self.frame1, text=" / ")
        self.line_label.place(x=35, y=150)

        self.max_label = tk.Label(self.frame1, text="__")
        self.max_label.place(x=50, y=150)

        # Create and place the button
        self.btnSplit = tk.Button(self.frame1, text='Split', width=14, command=self.start_split_pdf)
        self.btnSplit.pack(side=tk.BOTTOM, pady=20, padx=10)

    
    def load_image(self):
        self.file_path = filedialog.askopenfilename(
            filetypes=[("PDF files", "*.pdf")],
            title="Choose a PDF file"
        )

        if self.file_path:
            if os.path.exists(self.file_path):
                pdf_document = fitz.open(self.file_path)                
                self.max_label.config(text=len(pdf_document))

                page = pdf_document.load_page(0)

                # Set a higher resolution for better image quality
                zoom_x = 2.0  # horizontal zoom factor
                zoom_y = 2.0  # vertical zoom factor
                mat = fitz.Matrix(zoom_x, zoom_y)

                # Render page to an image with the specified zoom factor
                pix = page.get_pixmap(matrix=mat)
                
                # Convert the PDF page to an image
                self.img = Image.open(io.BytesIO(pix.pil_tobytes('png')))

                self.img_tk = ImageTk.PhotoImage(self.img)
                self.canvas.create_image(0, 0, anchor=tk.NW, image=self.img_tk)
                
                # img.save('original_image.png')
            else:
                print('file yoq')   
    
    def select_special_part(self):
        if self.img:
            self.canvas.bind("<ButtonPress-1>", self.on_press)
            self.canvas.bind("<B1-Motion>", self.on_drag)
            self.canvas.bind("<ButtonRelease-1>", self.on_release)

    def on_press(self, event):
        self.start_x = event.x
        self.start_y = event.y
        if self.rect:
            self.canvas.delete(self.rect)
        self.rect = self.canvas.create_rectangle(self.start_x, self.start_y, self.start_x, self.start_y, outline='red')

    def on_drag(self, event):
        self.canvas.coords(self.rect, self.start_x, self.start_y, event.x, event.y)

    def on_release(self, event):
        self.end_x = event.x
        self.end_y = event.y
        self.rect_coords = (self.start_x, self.start_y, self.end_x, self.end_y)
        
        #print("Selected region coordinates:", self.rect_coords)

        # Calculate and display location values
        self.left = min(self.start_x, self.end_x)
        self.right = max(self.start_x, self.end_x)
        self.upper = min(self.start_y, self.end_y)
        self.lower = max(self.start_y, self.end_y)
    
    def start_split_pdf(self):
        # create a folder if not exist
        currentFolderPath = os.path.abspath(os.getcwd())
        currentFolderPath = currentFolderPath + '/files'      
        Path(currentFolderPath).mkdir(parents=True, exist_ok=True)  
        
        self.page_number = 0
        pdf_document = fitz.open(self.file_path)    
        self.max_pages = len(pdf_document)
        self.update_label() 

    def extract_text_from_image(self, image):

        cropped_image = image.crop((self.left, self.upper, self.right, self.lower))

        text = pytesseract.image_to_string(cropped_image, config='--psm 6')
        return text.strip()
    
    def update_label(self):
        if self.page_number < self.max_pages:
            self.count_label.config(text=self.page_number + 1)

            pdf_document = fitz.open(self.file_path)
            page = pdf_document.load_page(self.page_number)
            
            # Set a higher resolution for better image quality
            zoom_x = 2.0  # horizontal zoom factor
            zoom_y = 2.0  # vertical zoom factor
            mat = fitz.Matrix(zoom_x, zoom_y)

            # Render page to an image with the specified zoom factor
            pix = page.get_pixmap(matrix=mat)
            
            # Convert the PDF page to an image
            img = Image.open(io.BytesIO(pix.pil_tobytes('png')))

            # Extract the examinee number                    
            examinee_number = self.extract_text_from_image(img)

            #remove all spaces from a string
            examinee_number = examinee_number.replace(' ', '')
                                
            # Save the individual PDF page with the examinee number as the filename
            output_filename = f"files/{examinee_number}.pdf"
            output_path = os.path.join(output_filename)
            new_pdf = fitz.open()
            new_pdf.insert_pdf(pdf_document, from_page=self.page_number, to_page=self.page_number)
            new_pdf.save(output_path)
            new_pdf.close()

            self.page_number += 1
            self.master.after(10, self.update_label)  # Call this method again after 100 ms
            

root = tk.Tk()
app = ImageSelectorApp(root)
root.mainloop() # pyinstaller --onefile --icon=project/ready.ico  project/PdfSplitter.py
