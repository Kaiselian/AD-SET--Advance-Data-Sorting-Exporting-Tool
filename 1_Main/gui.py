# gui.py
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog
import pandas as pd
import ttkbootstrap as tb  # Modern UI Framework
import darkdetect #detect system default active UI
from fpdf import FPDF
import fitz  # PyMuPDF
from PIL import Image, ImageTk
import openpyxl
import pdfplumber  # Extract tables from PDFs
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib.styles import getSampleStyleSheet
import subprocess
import sys
import os

# Detect System Theme (Light/Dark)
def get_system_theme():
    return "darkly" if darkdetect.isDark() else "journal"

# 🔄 Function to Change Theme
def change_theme(selected_theme):
    global root
    root.style.theme_use(selected_theme)  # Apply the new theme instantly

# Initialize GUI
theme = "darkly" if darkdetect.isDark() else "journal"

root = tb.Window(themename=theme)  # Default theme, fixed
root.title("Advanced Data Search & Export Tool 1.13.5.7")
root.geometry("1920x1080")
root.state("zoomed")

class PDFApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Form Filler")
        self.setup_ui()

    def setup_ui(self):
        self.load_button = tk.Button(self.root, text="Load PDF", command=self.load_pdf)
        self.load_button.pack()

    def load_pdf(self):
        file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if file_path:
            load_pdf(file_path)

def run_gui():
    root = tk.Tk()
    app = PDFApp(root)
    root.mainloop()
