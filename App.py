import sys
import subprocess
import importlib
import io
import os
import tkinter as tk
from tkinter import filedialog

# --- Ensure required packages are installed before importing them ---
def _get_package_version(module):
    return getattr(module, "__version__", getattr(module, "Version", "?"))

def ensure_package(package_import_name, pip_name=None):
    package_to_install = pip_name or package_import_name
    print(f"Checking {package_to_install} ...", end=" ")
    try:
        module = importlib.import_module(package_import_name)
        print(f"OK (v{_get_package_version(module)})")
        return
    except ImportError:
        print("missing")

    print(f"Installing {package_to_install} ...")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", package_to_install])
        module = importlib.import_module(package_import_name)
        print(f"Installed {package_to_install} (v{_get_package_version(module)})")
    except Exception as e:
        print(f"Failed to install {package_to_install}: {e}")
        sys.exit(1)

# Map import names to pip package names where they differ
REQUIRED_PACKAGES = [
    ("pandas", "pandas"),
    ("PyPDF2", "PyPDF2"),
    ("reportlab", "reportlab"),
    ("openpyxl", "openpyxl"),
]

print("Checking required packages...")
for import_name, pip_name in REQUIRED_PACKAGES:
    ensure_package(import_name, pip_name)
print("All required packages are present.\n")

# Safe to import third-party libraries now
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# --- Setup tkinter ---
root = tk.Tk()
root.withdraw()

# --- Ask user for input files ---
names_file = filedialog.askopenfilename(
    title="Select Excel/CSV file with names",
    filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")]
)

if not names_file:
    print("No names file selected. Exiting...")
    exit()

template_pdf = filedialog.askopenfilename(
    title="Select Certificate Template PDF",
    filetypes=[("PDF files", "*.pdf")]
)

if not template_pdf:
    print("No template selected. Exiting...")
    exit()

output_folder = filedialog.askdirectory(title="Select folder to save certificates")

if not output_folder:
    print("No folder selected. Exiting...")
    exit()

# --- Load spreadsheet ---
if names_file.endswith(".xlsx"):
    names = pd.read_excel(names_file)
else:
    names = pd.read_csv(names_file)

# --- Register custom font (must have the .ttf file in same folder) ---
pdfmetrics.registerFont(TTFont("MyFont", "Montserrat-Medium.ttf"))

# --- Generate certificates ---
for index, row in names.iterrows():
    name = str(row['Name']).strip()  # use column "Name"

    # Create overlay with name
    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=letter)
    can.setFont("MyFont", 30)

    # Position (x, y) → adjust for your certificate layout
    can.drawCentredString(420, 300, name)

    can.save()
    packet.seek(0)

    # Merge overlay with certificate
    overlay_pdf = PdfReader(packet)
    template = PdfReader(open(template_pdf, "rb"))
    writer = PdfWriter()

    page = template.pages[0]
    page.merge_page(overlay_pdf.pages[0])
    writer.add_page(page)

    # Save
    output_path = os.path.join(output_folder, f"{name}.pdf")
    with open(output_path, "wb") as out_file:
        writer.write(out_file)

print(f"✅ Certificates generated in: {output_folder}")
