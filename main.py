import tkinter as tk
import tkinter.ttk as ttk 
from tkinter import filedialog, messagebox
import pandas as pd
import pytesseract
from pdf2image import convert_from_path
from PIL import Image, ImageFilter, ImageOps
from collections import defaultdict
import fitz
import logging
import re
import sys
import os

# POPPLER_PATH = "YOUR PATH GOES HERE" # poppler path goes here in case the other one does not work.

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# Replace your existing path definitions with these:
TESSERACT_PATH = resource_path(os.path.join('tesseract', 'tesseract.exe'))
POPPLER_PATH = resource_path(os.path.join('poppler', 'bin'))

pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH

Image.MAX_IMAGE_PIXELS = None

# New pattern for circle codes: 1-5 alphanumeric characters, case insensitive
CIRCLE_CODE_PATTERN = r'\b([A-Za-z0-9]{1,5})\b'

def preprocess_image(img):
    """Enhanced image preprocessing for better OCR accuracy"""
    try:
        # Convert to grayscale
        img = img.convert('L')

        # Auto-contrast to enhance text visibility
        img = ImageOps.autocontrast(img)

        # Adaptive binarization (with a lower threshold to catch faint text)
        threshold = 150
        img = img.point(lambda x: 0 if x < threshold else 255, '1')

        # Resize to double dimensions to help OCR with small fonts
        img = img.resize((img.width * 2, img.height * 2), Image.LANCZOS)

        # Sharpen text edges
        img = img.filter(ImageFilter.SHARPEN)
        img = img.filter(ImageFilter.EDGE_ENHANCE_MORE)

        return img
    except Exception as e:
        logging.error(f"Image preprocessing failed: {str(e)}")
        return img


def clean_ocr_text(text):
    """Clean OCR output text while preserving essential format"""
    # Standardize common OCR errors in material codes
    replacements = {
        "ﬁ": "fi",
        "ﬂ": "fl",
        "..": ".",
        "cm": "cm",
        "em": "cm",
        "m'": "m",
        "m\"": "m",
        "L =": "L=",
        "L= ": "L=",
        " ,": ",",
        " '": "'"
    }
    for wrong, right in replacements.items():
        text = text.replace(wrong, right)
    # Remove problematic characters but keep needed punctuation
    text = re.sub(r"[|!\"';~_]", "", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text

def preprocess_text(text):
    """Handle line breaks and hyphenation in material codes"""
    # Replace line breaks with space when they occur within material codes
    text = re.sub(r'(\d+Rfi[^\n]*)-\s*\n\s*(\d)', r'\1\2', text)
    # Standardize spaces around key elements
    text = re.sub(r'\s*,\s*L\s*=\s*', ',L=', text)
    text = re.sub(r'\s*/\s*', '/', text)
    text = re.sub(r'\s*cm\s*', 'cm', text)
    return text

def extract_materials(pdf_path, target_circle_codes):
    """Extracts materials from PDF (native + OCR) with duplicates"""
    results = []
    counts = defaultdict(int)

    # Try native PDF extraction first
    try:
        doc = fitz.open(pdf_path)
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            text = page.get_text()
            page_results, page_counts = extract_all_codes(text, target_circle_codes, str(page_num + 1))
            results.extend(page_results)
            for code, count in page_counts.items():
                counts[code] += count
    except Exception as e:
        logging.warning(f"Native PDF extraction failed: {str(e)}")

    # Fallback to OCR if no results
    if not results:
        pages = convert_from_path(pdf_path, dpi=300, poppler_path=POPPLER_PATH)
        for page_num, page_img in enumerate(pages, 1):
            text = pytesseract.image_to_string(page_img, config='--oem 3 --psm 6')
            page_results, page_counts = extract_all_codes(text, target_circle_codes, str(page_num))
            results.extend(page_results)
            for code, count in page_counts.items():
                counts[code] += count

    return results, counts

def extract_all_codes(text, target_circle_codes, current_page="N/A"):
    results = []
    counts = defaultdict(int)
    text = re.sub(r"[’']", "'", text)
    text = re.sub(r"\r", "", text)
    lines = text.split("\n")

    # Material code pattern (with length or with spacing)
    material_pattern = re.compile(
        r"""
        (\d+Rfi(?:\d+(?:/\d+)*)?(?:cm)?[,]?L=\d+\.?\d*m['']?)  # Format 1
        |                                                       # OR
        (\d+Rfi\d+/\d+cm)                                      # Format 2
        """,
        re.VERBOSE | re.IGNORECASE
    )

    # Strict circle code pattern: e.g., T1, T10, R14, B1, UT9, UT7a, or just numbers (1, 10, 14, etc.)
    strict_circle_pattern = re.compile(r"([A-Z]{1,2}\d{1,2}[a-z]?|\d{1,3})")

    i = 0
    while i < len(lines):
        line = lines[i].strip()
        if not line:
            i += 1
            continue
        mat_match = material_pattern.search(line)
        if mat_match:
            material_code = mat_match.group(1) or mat_match.group(2)
            clean_material = re.sub(r"\s*,\s*L\s*=\s*", ",L=", material_code)
            clean_material = re.sub(r"\s*/\s*", "/", clean_material)
            clean_material = re.sub(r"\s*cm\s*", "cm", clean_material)
            clean_material = re.sub(r"\s*Rfi\s*", "Rfi", clean_material)
            clean_material = clean_material.rstrip("'")
            # Look for a circle code to the right of the material code
            right_text = line[mat_match.end():]
            code_match = strict_circle_pattern.search(right_text)
            if code_match:
                circle_code = code_match.group(1)
                results.append({
                    "Circle Code": circle_code,
                    "Material Code": clean_material,
                    "Page": current_page
                })
                counts[circle_code] += 1
            else:
                # Look at the next non-empty line for a circle code
                j = i + 1
                found_code = None
                while j < len(lines):
                    next_line = lines[j].strip()
                    if next_line:
                        code_match_next = strict_circle_pattern.fullmatch(next_line)
                        if code_match_next:
                            found_code = code_match_next.group(1)
                        break
                    j += 1
                if found_code:
                    results.append({
                        "Circle Code": found_code,
                        "Material Code": clean_material,
                        "Page": current_page
                    })
                    counts[found_code] += 1
                else:
                    results.append({
                        "Circle Code": "",
                        "Material Code": clean_material,
                        "Page": current_page
                    })
        i += 1
    return results, counts

def clean_material_code(material_code):
    """More flexible cleaning of material codes"""
    clean = re.sub(r"\s+", "", material_code)
    clean = re.sub(r"(?i)rfi?", "Rfi", clean)
    clean = re.sub(r",\s*L\s*=", ",L=", clean)
    clean = re.sub(r"([.,])(\d*m)", r".\2", clean)

    # Handle cases where cm might be missing
    if "cm" not in clean.lower():
        clean = re.sub(r"(?i)(\d+rfi\d+(?:/\d+)*),", r"\1cm,", clean)

    return clean if validate_material_code(clean) else None


def validate_material_code(material_code):
    """Validates material code format"""
    # More flexible pattern to handle variations
    pattern = r"""
        ^\d+              # Number at start
        Rfi              # Rfi
        \d+              # Number
        /\d+cm           # /number cm
        ,L=\d+\.?\d*m$   # ,L=number.m
    """
    return bool(re.match(pattern, material_code, re.VERBOSE | re.IGNORECASE))



def save_to_excel(data, save_path):
    """Save results to Excel with proper formatting"""
    try:
        df = pd.DataFrame(data)
        if not df.empty:
            # Clean and standardize the data before saving
            df['Material Code'] = df['Material Code'].str.replace(r"\s+", "", regex=True)
            df['Material Code'] = df['Material Code'].str.replace(r",L=", ",L=")
            df = df[['Circle Code', 'Material Code', 'Page']]
            df.to_excel(save_path, index=False)
            logging.info(f"Saved {len(df)} records to {save_path}")
            return True
        return False
    except Exception as e:
        logging.error(f"Error saving to Excel: {str(e)}")
        return False


def browse_pdf():
    """File browser for PDF selection"""
    file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    pdf_entry.delete(0, tk.END)
    pdf_entry.insert(0, file_path)


def extract_all_materials(text):
    """Extract all material codes, regardless of circle code association"""
    # Pattern for both types: with length and with spacing
    material_pattern = re.compile(
        r"""
        (\d+Rfi(?:\d+(?:/\d+)*)?(?:cm)?[,]?L=\d+\.?\d*m['']?)  # Format 1
        |                                                       # OR
        (\d+Rfi\d+/\d+cm)                                      # Format 2
        """,
        re.VERBOSE | re.IGNORECASE
    )
    materials = set()
    for m in material_pattern.finditer(text):
        code = m.group(1) or m.group(2)
        if code:
            clean_material = re.sub(r"\s*,\s*L\s*=\s*", ",L=", code)
            clean_material = re.sub(r"\s*/\s*", "/", clean_material)
            clean_material = re.sub(r"\s*cm\s*", "cm", clean_material)
            clean_material = re.sub(r"\s*Rfi\s*", "Rfi", clean_material)
            clean_material = clean_material.rstrip("'")
            materials.add(clean_material)
    return list(materials)

def process_pdf():
    """Main processing function with improved error handling and dual material counting """
    pdf_path = pdf_entry.get()
    circle_codes_input = circle_entry.get().strip()
    
    if not pdf_path:
        messagebox.showerror("Error", "Please provide a PDF path!")
        return
        
    if not circle_codes_input:
        messagebox.showerror("Error", "Please enter at least one circle code!")
        return

    try:
        # Process circle codes
        circle_codes = [code.strip() for code in circle_codes_input.split(",") if code.strip()]
        circle_codes_upper = [code.upper() for code in circle_codes]
        
        if not circle_codes:
            messagebox.showerror("Error", "Please enter valid circle codes!")
            return

        # First extract all materials from the PDF (all pages)
        doc = fitz.open(pdf_path)
        all_text = "\n".join([page.get_text() for page in doc])
        all_materials = extract_all_materials(all_text)

        # Extract all circle code/material pairs
        all_results, all_counts = extract_all_codes_from_pdf(pdf_path)
        # Only keep results with a circle code and that match the entered circle codes (case-insensitive)
        filtered_results = [
            item for item in all_results
            if item['Circle Code'] and item['Circle Code'].upper() in circle_codes_upper
        ]

        # Clear previous results
        results_text.config(state=tk.NORMAL)
        results_text.delete(1.0, tk.END)

        # Show all materials count (filtered)
        results_text.insert(tk.END, f"ALL MATERIALS WITH CIRCLE CODE COUNT: {len(filtered_results)}\n")
        results_text.insert(tk.END, "-" * 50 + "\n")
        for item in filtered_results:
            results_text.insert(tk.END, f"{item['Circle Code']}: {item['Material Code']} (Page {item['Page']})\n")
        results_text.insert(tk.END, "\n")

        # Show materials with circle codes (case-insensitive count, display as entered)
        if filtered_results:
            results_text.insert(tk.END, "MATERIALS WITH CIRCLE CODES:\n")
            results_text.insert(tk.END, "-" * 50 + "\n")
            for code, code_upper in zip(circle_codes, circle_codes_upper):
                count = sum(1 for item in filtered_results if item['Circle Code'].upper() == code_upper)
                results_text.insert(tk.END, f"{code}: {count} materials found\n")
            results_text.insert(tk.END, "\n" + "=" * 50 + "\n")
            results_text.insert(tk.END, f"TOTAL MATERIALS WITH CIRCLE CODES: {len(filtered_results)}\n")
            save_button.grid(row=5, column=1, pady=10)
        else:
            results_text.insert(tk.END, f"No materials found for the specified circle codes: {', '.join(circle_codes)}")

        results_text.config(state=tk.DISABLED)

    except Exception as e:
        logging.error(f"Processing error: {str(e)}")
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

def extract_all_codes_from_pdf(pdf_path):
    """Extract all circle codes and their material codes from the PDF"""
    try:
        doc = fitz.open(pdf_path)
        results = []
        code_counts = defaultdict(int)

        # First pass: collect all circle codes
        all_text = "\n".join([page.get_text() for page in doc])
        circle_codes = set(re.findall(r'\b([A-Za-z0-9]{1,5})\b', all_text, re.IGNORECASE))
        circle_codes = {code.upper() for code in circle_codes
                       if re.match(r'^[A-Za-z0-9]{1,5}$', code, re.IGNORECASE) and code != '0'}

        if not circle_codes:
            logging.info("No circle codes found in document")
            return [], defaultdict(int)

        # Second pass: extract materials with page numbers
        for page_num, page in enumerate(doc, 1):
            text = page.get_text()
            # Pass empty set to extract_all_codes to get all materials
            page_results, page_counts = extract_all_codes(text, set(), str(page_num))
            results.extend(page_results)
            for code, count in page_counts.items():
                code_counts[code] += count

        if results:
            return results, code_counts

    except Exception as e:
        logging.warning(f"Native PDF extraction failed: {str(e)}")

    # OCR fallback
    pages = convert_from_path(pdf_path, dpi=500, poppler_path=POPPLER_PATH)

    results = []
    code_counts = defaultdict(int)

    for page_num, page_img in enumerate(pages, 1):
        text = pytesseract.image_to_string(page_img, config='--oem 3 --psm 6')
        text = clean_ocr_text(text)
        # Pass empty set to extract_all_codes to get all materials
        page_results, page_counts = extract_all_codes(text, set(), str(page_num))
        results.extend(page_results)
        for code, count in page_counts.items():
            code_counts[code] += count

    return results, code_counts


def process_all_materials():
    """Process PDF to extract all materials (with or without circle codes)"""
    pdf_path = pdf_entry.get()

    if not pdf_path:
        messagebox.showerror("Error", "Please select a PDF file first!")
        return

    try:
        results, counts = extract_all_codes_from_pdf(pdf_path)

        # Clear previous results
        results_text.config(state=tk.NORMAL)
        results_text.delete(1.0, tk.END)

        if results:
            results_text.insert(tk.END, "ALL MATERIALS (WITH OR WITHOUT CIRCLE CODES):\n\n")
            results_text.insert(tk.END, "-" * 50 + "\n")
            for item in results:
                code_display = f"{item['Circle Code']}: " if item['Circle Code'] else ""
                results_text.insert(tk.END, f"{code_display}{item['Material Code']} (Page {item['Page']})\n")
            results_text.insert(tk.END, "\n" + "=" * 50 + "\n")
            results_text.insert(tk.END, f"TOTAL MATERIALS FOUND: {len(results)}\n")
            save_button.grid(row=5, column=1, pady=10)
        else:
            results_text.insert(tk.END, "No materials found")
        results_text.config(state=tk.DISABLED)
    except Exception as e:
        logging.error(f"Processing error: {str(e)}")
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

# Update process_all_codes to only show materials with a circle code
def process_all_codes():
    """Process PDF to extract all circle codes and their materials (only those with a circle code)"""
    pdf_path = pdf_entry.get()

    if not pdf_path:
        messagebox.showerror("Error", "Please select a PDF file first!")
        return

    try:
        results, counts = extract_all_codes_from_pdf(pdf_path)
        # Filter to only those with a circle code
        filtered_results = [item for item in results if item['Circle Code']]
        filtered_counts = defaultdict(int)
        for item in filtered_results:
            filtered_counts[item['Circle Code']] += 1

        # Clear previous results
        results_text.config(state=tk.NORMAL)
        results_text.delete(1.0, tk.END)

        if filtered_results:
            results_text.insert(tk.END, "ALL CIRCLE CODES AND MATERIALS FOUND (ONLY WITH CIRCLE CODE):\n\n")
            results_text.insert(tk.END, "MATERIAL CODE COUNTS:\n")
            results_text.insert(tk.END, "-" * 50 + "\n")
            for code, count in filtered_counts.items():
                results_text.insert(tk.END, f"{code}: {count} materials found\n")
            results_text.insert(tk.END, "\nALL MATERIAL CODES:\n")
            results_text.insert(tk.END, "-" * 50 + "\n")
            for item in filtered_results:
                results_text.insert(tk.END, f"{item['Circle Code']}: {item['Material Code']} (Page {item['Page']})\n")
            results_text.insert(tk.END, "\n" + "=" * 50 + "\n")
            results_text.insert(tk.END, f"TOTAL MATERIALS FOUND: {len(filtered_results)}\n")
            save_button.grid(row=5, column=1, pady=10)
        else:
            results_text.insert(tk.END, "No circle codes and materials found")
        results_text.config(state=tk.DISABLED)
    except Exception as e:
        logging.error(f"Processing error: {str(e)}")
        messagebox.showerror("Error", f"An error occurred: {str(e)}")


def save_results():
    """Save the current results to Excel, exporting all columns present in the current output (with or without circle code)"""
    # Get the text from results_text widget
    content = results_text.get(1.0, tk.END)

    # Check if we have results to save
    if "TOTAL MATERIALS FOUND: 0" in content or "No circle codes" in content or "No materials found" in content:
        messagebox.showwarning("No Results", "There are no results to save")
        return

    save_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")],
        initialfile="Circle_Codes_Results.xlsx"
    )

    if save_path:
        # Parse the results from the text widget
        lines = content.split('\n')
        results = []
        for line in lines:
            # Try to match lines like: CODE: MATERIAL (Page X)
            if ": " in line and "Page" in line:
                parts = line.split(': ', 1)
                circle_code = parts[0].strip()
                rest = parts[1].split('(Page ')
                material_code = rest[0].strip()
                page = rest[1].replace(')', '').strip() if len(rest) > 1 else ''
                # If circle_code is empty or just a dash, treat as blank
                if circle_code == '' or circle_code == '-':
                    circle_code = ''
                results.append({
                    "Circle Code": circle_code,
                    "Material Code": material_code,
                    "Page": page
                })
            # Or lines like: MATERIAL (Page X) (no circle code)
            elif "(Page" in line and ":" not in line:
                material_code = line.split('(Page')[0].strip()
                page = line.split('(Page')[1].replace(')', '').strip() if '(Page' in line else ''
                results.append({
                    "Circle Code": '',
                    "Material Code": material_code,
                    "Page": page
                })
        if results:
            # Dynamically determine columns
            import pandas as pd
            df = pd.DataFrame(results)
            # Only keep columns that are present in the data
            columns = [col for col in ["Circle Code", "Material Code", "Page"] if col in df.columns]
            df = df[columns]
            df.to_excel(save_path, index=False)
            messagebox.showinfo("Saved", f"Results saved to {save_path}")
        else:
            messagebox.showwarning("No Results", "Could not parse results for saving")


# === GUI Setup ===
root = tk.Tk()
root.title("PDF Material Extractor")
root.geometry("1000x600")
root.configure(bg="#f4f4f4")

# Style
style = ttk.Style()
style.theme_use('clam')
style.configure('TLabel', font=('Segoe UI', 11), background="#f4f4f4")
style.configure('TButton', font=('Segoe UI', 11), padding=3, background="lightblue")
style.configure('TEntry', font=('Segoe UI', 11))
style.configure('TFrame', background="#f4f4f4")
style.configure('TLabelframe', background="#f4f4f4", font=('Segoe UI', 12, 'bold'))

# Main frame
main_frame = ttk.Frame(root, padding="20 15 20 15")
main_frame.grid(row=0, column=0, sticky="nsew")
root.grid_rowconfigure(0, weight=1)
root.grid_columnconfigure(0, weight=1)

# Configure grid weights for resizing
root.grid_rowconfigure(4, weight=1)
root.grid_columnconfigure(1, weight=1)

# PDF Selection
pdf_frame = ttk.LabelFrame(main_frame, text="PDF Selection", padding="10 10 10 10")
pdf_frame.grid(row=0, column=0, sticky="ew", pady=10)
pdf_frame.grid_columnconfigure(1, weight=1)
ttk.Label(pdf_frame, text="Select PDF File:").grid(row=0, column=0, padx=5, pady=5, sticky='e')
pdf_entry = ttk.Entry(pdf_frame, width=60)
pdf_entry.grid(row=0, column=1, padx=5, pady=5, sticky='ew')
ttk.Button(pdf_frame, text="Browse", command=browse_pdf).grid(row=0, column=2, padx=5, pady=5)


# Circle Codes Input
circle_frame = ttk.LabelFrame(main_frame, text="Circle Codes", padding="10 10 10 10")
circle_frame.grid(row=1, column=0, sticky="ew", pady=10)
circle_frame.grid_columnconfigure(1, weight=1)
ttk.Label(circle_frame, text="Enter Circle Codes (comma separated):").grid(row=0, column=0, padx=5, pady=5, sticky='e')
circle_entry = ttk.Entry(circle_frame, width=60)
circle_entry.grid(row=0, column=1, padx=5, pady=5, sticky='ew')

# Process Buttons
button_frame = ttk.Frame(main_frame)
button_frame.grid(row=2, column=0, pady=10, sticky="ew")
button_frame.grid_columnconfigure((0, 1, 2), weight=1)

# Extract buttons
extract_material_btn = ttk.Button(button_frame, text="Extract Material Codes", command=process_pdf)
extract_material_btn.grid(row=0, column=0, padx=5, ipadx=5, ipady=3, sticky="ew")
extract_all_codes_btn = ttk.Button(button_frame, text="Extract All Circle Codes", command=process_all_codes)
extract_all_codes_btn.grid(row=0, column=1, padx=5, ipadx=5, ipady=3, sticky="ew")
extract_all_materials_btn = ttk.Button(button_frame, text="Extract All Materials", command=process_all_materials)
extract_all_materials_btn.grid(row=0, column=2, padx=5, ipadx=5, ipady=3, sticky="ew")


# Results Text Area with Scrollbar
results_frame = ttk.LabelFrame(main_frame, text="Results", padding="10 10 10 10")
results_frame.grid(row=3, column=0, padx=10, pady=5, sticky='ew')
results_frame.grid_rowconfigure(0, weight=1)
results_frame.grid_columnconfigure(0, weight=1)

results_text = tk.Text(results_frame, wrap=tk.WORD, font=('Consolas', 11), bg="#f9f9f9", relief=tk.FLAT, borderwidth=2)
results_text.grid(row=0, column=0, sticky="nsew")
scrollbar = ttk.Scrollbar(results_frame, command=results_text.yview)
scrollbar.grid(row=0, column=1, sticky='ns')
results_text['yscrollcommand'] = scrollbar.set

# Save Button (initially hidden)
save_button = ttk.Button(main_frame, text="Save Results to Excel", command=save_results)
save_button.grid(row=4, column=0, pady=(5,15), sticky='w', padx=10)
save_button.grid_remove()

main_frame.grid_rowconfigure(3, weight=1)
main_frame.grid_columnconfigure(0, weight=1)

root.mainloop()