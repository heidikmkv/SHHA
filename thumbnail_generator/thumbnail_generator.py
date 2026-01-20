import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image
import fitz  # PyMuPDF
import re
import os


def generate_thumbnail_image(pdf_path,output_name):
    doc = fitz.open(pdf_path)
    page = doc.load_page(0)  # First page
    pix = page.get_pixmap()
    
    # Convert pixmap to PIL Image
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    
    # OPTION 1: Crop to 600x400 pixels, keeping the top of the page
    #cropped_img = img.crop((0, 0, 600, 400))
    #cropped_img.save(output_path)

    # OPTION 2: Center the whole page image in 600*400px white background
    # Calculate the new width while keeping aspect ratio
    new_height = 400
    aspect_ratio = img.width / img.height
    new_width = int(aspect_ratio * new_height)
    
    # Resize the image
    resized_img = img.resize((new_width, new_height))
    
    # Create a white background (600x400)
    im = Image.new('RGB', (600, 400), (255, 255, 255))
    
    # Center the resized image on the background
    offset = ((600 - new_width) // 2, 0)  # Center horizontally
    im.paste(resized_img, offset)
    im.save(output_name)
    return im

# Dictionary to map month names, abbreviations, and seasons to month numbers
MONTH_MAP = {
    "january": "01", "jan": "01",
    "february": "02", "feb": "02",
    "march": "03", "mar": "03",
    "april": "04", "apr": "04",
    "may": "05",
    "june": "06", "jun": "06",
    "july": "07", "jul": "07",
    "august": "08", "aug": "08",
    "september": "09", "sep": "09", "sept": "09",
    "october": "10", "oct": "10",
    "november": "11", "nov": "11",
    "december": "12", "dec": "12",
    "winter": "01", "spring": "03", "summer": "06", "fall": "09", "autumn": "09",
}

# Regular expression to capture numeric months (e.g., 01, 1, etc.)
MONTH_NUM_REGEX = re.compile(r'(\d{1,2})')

# Function to detect month from filename
def detect_month(filename):
    filename_lower = filename.lower()
    
    # Check for month names or seasons
    for month_name, month_num in MONTH_MAP.items():
        if month_name in filename_lower:
            return month_num
    
    # Check for numeric month (e.g., 01, 1, etc.)
    month_match = MONTH_NUM_REGEX.search(filename_lower)
    if month_match:
        month = month_match.group(1).zfill(2)  # Pad with zero if needed
        if 1 <= int(month) <= 12:  # Ensure it's a valid month
            return month
    
    return None

def detect_year(filename):
    # Regular expression to match a 4-digit year between 1900 and 2099
    match = re.search(r'\b(19|20)\d{2}\b', filename)
    
    # Return the year if found, otherwise return None
    return int(match.group()) if match else None


def browse_file():
    pdf_path = filedialog.askopenfilename(
        filetypes=[("PDF Files", "*.pdf")], title="Select a PDF file"
    )
    if pdf_path:
        try:
            month = detect_month(os.path.basename(pdf_path))
            year = detect_year(os.path.basename(pdf_path))
            if month is not None and year is not None:
                output_name = f"SHHA-GRIT-{year}_{str(month).zfill(2)}.png"
            else:
                output_name = 'thumbnail.png'
            thumbnail = generate_thumbnail_image(pdf_path,output_name)
            messagebox.showinfo("Success", "Thumbnail generated and saved as "+output_name)
            thumbnail.show()  # Opens the generated thumbnail for a quick preview

        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate thumbnail: {e}")

# Set up the main application window
root = tk.Tk()
root.title("GRIT Thumbnail Generator")
root.geometry("600x400")

# Add a button to select the PDF file
browse_button = tk.Button(root, text="Select PDF and Generate Thumbnail", command=browse_file)
browse_button.pack(pady=20)

# Run the GUI loop
root.mainloop()