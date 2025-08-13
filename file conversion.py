import tkinter as tk
import PIL
from tkinter import filedialog, messagebox
from PIL import Image, ImageDraw, ImageFont, UnidentifiedImageError
import pdf2image
import docx
import os
import zipfile


# Define Poppler Path (Modify for Windows)
poppler_path = r"C:\Users\AAKASH VELAN.M\Documents\internship\Release-24.08.0-0\poppler-24.08.0\Library\bin"


def choose_file():
    file_paths = filedialog.askopenfilenames(filetypes=[
        ("All Supported Files", "*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.pdf;*.docx;*.txt"),
        ("Image Files", "*.png;*.jpg;*.jpeg;*.bmp;*.gif"),
        ("PDF Files", "*.pdf"),
        ("Word Files", "*.docx"),
        ("Text Files", "*.txt")
    ])
    
    if file_paths:
        file_label.config(text=f"Selected {len(file_paths)} files")
        root.selected_files = file_paths

def optimize_tiff(tiff_path):
    """Reduce TIFF size to ensure it stays under 1MB while maintaining dimensions."""
    max_size = 1024 * 1024  # 1MB limit
    quality = 90  # Start with high quality

    while os.path.getsize(tiff_path) > max_size and quality > 10:
        img = Image.open(tiff_path)
        
        # Resize if necessary (keeping aspect ratio)
        width, height = img.size
        scale_factor = 0.9  # Start with 10% reduction
        while os.path.getsize(tiff_path) > max_size and min(width, height) > 500:
            width = int(width * scale_factor)
            height = int(height * scale_factor)
            img = img.resize((width, height), Image.LANCZOS)
        
        # Save with reduced quality
        img.save(tiff_path, format="TIFF", compression="tiff_lzw")
        quality -= 10  # Reduce quality gradually
        
        while os.path.getsize(tiff_path) > max_size:
             img = Image.open(tiff_path)
             width, height = img.size
             img = img.resize((int(width * 0.9), int(height * 0.9)), Image.LANCZOS)  # Reduce further if needed
             img.save(tiff_path, format="TIFF", compression="tiff_lzw")
    
    print(f"Optimized {tiff_path}: {os.path.getsize(tiff_path) / 1024:.2f} KB")


def convert_to_tiff(input_path):
    file_ext = input_path.split('.')[-1].lower()
    tiff_paths = []
    
    try:
        if file_ext in ["png", "jpg", "jpeg", "bmp", "gif"]:
            img = Image.open(input_path).convert("RGB")  # Ensure it's in RGB mode
            tiff_path = input_path.rsplit('.', 1)[0] + ".tiff"

            # First save attempt
            img.save(tiff_path, format="TIFF", compression="tiff_lzw")
            
            # Optimize after saving
            optimize_tiff(tiff_path)
            tiff_paths.append(tiff_path)

        elif file_ext == "pdf":
            images = pdf2image.convert_from_path(input_path, poppler_path=poppler_path, dpi=100)  # Lower DPI to reduce size
            tiff_path = input_path.rsplit('.', 1)[0] + ".tiff"

            # Convert all pages to grayscale and save as multi-page TIFF
            images = [img.convert("L") for img in images]  
            images[0].save(tiff_path, save_all=True, append_images=images[1:], format="TIFF", compression="tiff_lzw")

            optimize_tiff(tiff_path)
            tiff_paths.append(tiff_path)

        elif file_ext == "docx":
            doc = docx.Document(input_path)
            paragraphs = [para.text for para in doc.paragraphs if para.text.strip()]  # Remove empty paragraphs
            
            if not paragraphs:
                raise ValueError("The document contains no text.")
            
            text = "\n\n".join(paragraphs)
            tiff_path = render_text_as_image(text, input_path)
            optimize_tiff(tiff_path)
            tiff_paths.append(tiff_path)

        elif file_ext == "txt":
            with open(input_path, "r", encoding="utf-8") as f:
                text = f.read()
            tiff_path = render_text_as_image(text, input_path)
            optimize_tiff(tiff_path)
            tiff_paths.append(tiff_path)

        else:
            raise ValueError("Unsupported file format")

        return tiff_paths
    
    except UnidentifiedImageError:
        messagebox.showerror("Error", "Selected file is not a valid image.")
    except Exception as e:
        messagebox.showerror("Error", f"Conversion failed: {e}")


def render_text_as_image(text, input_path):
    """Convert multi-page text to a TIFF image."""
    font_size = 18
    width, height = 800, 1000
    lines_per_page = 40  

    try:
        font = ImageFont.truetype("arial.ttf", font_size)
    except:
        font = ImageFont.load_default()

    lines = text.split("\n")
    pages = [lines[i:i + lines_per_page] for i in range(0, len(lines), lines_per_page)]
    
    images = []
    for page in pages:
        img = Image.new("L", (width, height), "white")  # Grayscale for smaller file size
        draw = ImageDraw.Draw(img)
        y_offset = 50
        for line in page:
            draw.text((50, y_offset), line, font=font, fill="black")
            y_offset += font_size + 5  

        images.append(img)

    tiff_path = input_path.rsplit('.', 1)[0] + ".tiff"
    images[0].save(tiff_path, save_all=True, append_images=images[1:], format="TIFF", compression="tiff_lzw")

    return tiff_path


def zip_tiff_files(tiff_paths):
    """Zip all converted TIFF files."""
    zip_filename = filedialog.asksaveasfilename(defaultextension=".zip", filetypes=[("ZIP Files", "*.zip")])
    if zip_filename:
        try:
            with zipfile.ZipFile(zip_filename, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for tiff_path in tiff_paths:
                    zipf.write(tiff_path, os.path.basename(tiff_path))
            messagebox.showinfo("Success", f"TIFF files successfully zipped: {zip_filename}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to create ZIP file: {e}")

def upload_and_convert():
    """Process selected files and convert them to TIFF."""
    if hasattr(root, 'selected_files'):
        all_tiff_paths = []
        for file in root.selected_files:
            tiff_paths = convert_to_tiff(file)
            if tiff_paths:
                all_tiff_paths.extend(tiff_paths)
        
        if all_tiff_paths:
            zip_tiff_files(all_tiff_paths)
    else:
        messagebox.showwarning("Warning", "Please choose files first!")

# Initialize the main window
root = tk.Tk()
root.title("Document to TIFF Converter")
root.geometry("500x250")

# Create widgets
file_label = tk.Label(root, text="No file selected", wraplength=450)
file_label.pack(pady=10)

choose_button = tk.Button(root, text="Choose Files", command=choose_file)
choose_button.pack(pady=5)

upload_button = tk.Button(root, text="Upload and Convert", command=upload_and_convert)
upload_button.pack(pady=5)

# Run the application
root.mainloop()
