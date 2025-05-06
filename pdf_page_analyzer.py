import sys
import os
import pytesseract
import fitz  # PyMuPDF
import enchant
import re
import logging
from PIL import Image
import gc
import csv
from datetime import datetime
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QFileDialog, QMessageBox, QTextEdit, QTableWidget,
    QTableWidgetItem, QHeaderView, QScrollArea, QLabel
)
from PyQt5.QtCore import Qt

# Dynamically set the path to Tesseract binary in the bundled app
if getattr(sys, 'frozen', False):  # Check if running as a PyInstaller bundle
    bundle_dir = sys._MEIPASS
    tesseract_path = os.path.join(bundle_dir, 'tesseract', 'tesseract.exe' if sys.platform == 'win32' else 'tesseract')
    pytesseract.pytesseract.tesseract_cmd = tesseract_path
    os.environ['TESSDATA_PREFIX'] = os.path.join(bundle_dir, 'tessdata')
else:
    pytesseract.pytesseract.tesseract_cmd = 'tesseract'  # Assumes Tesseract is in PATH

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
logger = logging.getLogger(__name__)

def is_blank_page(text, char_threshold=50):
    """Check if page is blank based on character count."""
    return len(text.strip()) < char_threshold

def is_gibberish_page(text, valid_word_threshold=0.1):
    """Check if page contains mostly gibberish text."""
    if not text.strip():
        return False
    words = re.findall(r'\b\w+\b', text.lower())
    if not words:
        return True
    dictionary = enchant.Dict("en_US")
    valid_words = sum(1 for word in words if dictionary.check(word) and len(word) > 1)
    valid_ratio = valid_words / len(words)
    return valid_ratio < valid_word_threshold

def extract_text_with_pymupdf_image(page, page_num):
    """Extract text from a page image using PyMuPDF and OCR."""
    try:
        pix = page.get_pixmap(dpi=300)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.rgb)
        text = pytesseract.image_to_string(img, lang='eng')
        logger.info(f"OCR text length for page {page_num}: {len(text)} characters")
        logger.debug(f"OCR text sample for page {page_num}: {text[:50]}...")
        pix = None
        img.close()
        del img
        gc.collect()
        return text
    except Exception as e:
        logger.error(f"PyMuPDF image OCR failed for page {page_num}: {e}")
        return ""

def analyze_pdf(pdf_path, log_widget):
    """Analyze PDF to count blank, gibberish, and billable pages."""
    blank_pages = 0
    gibberish_pages = 0
    billable_pages = 0
    page_details = []

    try:
        doc = fitz.open(pdf_path)
        total_pages = len(doc)
        log_widget.append(f"Total pages: {total_pages}")
        log_widget.repaint()

        for i in range(total_pages):
            page_num = i + 1
            log_widget.append(f"Processing page {page_num}/{total_pages}...")
            log_widget.repaint()
            try:
                page = doc[i]
                text = page.get_text("text") or ""
                log_widget.append(f"PyMuPDF text length for page {page_num}: {len(text)} characters")
                logger.debug(f"PyMuPDF text sample for page {page_num}: {text[:50]}...")
            except Exception as e:
                logger.warning(f"Text extraction failed for page {page_num}: {e}")
                text = ""

            if not text.strip():
                log_widget.append(f"Attempting OCR for page {page_num}")
                log_widget.repaint()
                text = extract_text_with_pymupdf_image(page, page_num)

            if is_blank_page(text):
                status = "Blank"
                blank_pages += 1
            elif is_gibberish_page(text):
                status = "Gibberish"
                gibberish_pages += 1
            else:
                status = "Billable"
                billable_pages += 1
            log_widget.append(f"Page {page_num}: {status}")
            log_widget.repaint()
            page_details.append((page_num, status, len(text)))

        doc.close()

    except Exception as e:
        logger.error(f"Error processing PDF: {e}")
        log_widget.append(f"Error processing PDF: {e}")
        return None

    return {
        "total_pages": total_pages,
        "blank_pages": blank_pages,
        "gibberish_pages": gibberish_pages,
        "billable_pages": billable_pages,
        "page_details": page_details
    }

class PDFAnalyzerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF Page Analyzer")
        self.setGeometry(100, 100, 800, 600)

        # Main widget and layout
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.layout = QVBoxLayout(self.central_widget)
        self.layout.setSpacing(10)

        # File selection section
        file_widget = QWidget()
        file_layout = QVBoxLayout(file_widget)
        file_layout.setContentsMargins(10, 10, 10, 10)
        file_widget.setStyleSheet("background-color: #e9ecef; border-radius: 5px;")
        self.file_label = QLabel("Selected PDFs: None")
        self.file_label.setStyleSheet("font: bold 14px 'Segoe UI'; color: #2b3e50;")
        file_layout.addWidget(self.file_label)
        button_layout = QHBoxLayout()
        self.browse_button = QPushButton("Browse PDFs")
        self.browse_button.setStyleSheet("background-color: #28a745; color: white; font: bold 12px 'Segoe UI'; padding: 8px; border-radius: 5px;")
        self.browse_button.clicked.connect(self.browse_files)
        self.analyze_button = QPushButton("Analyze PDFs")
        self.analyze_button.setStyleSheet("background-color: #28a745; color: white; font: bold 12px 'Segoe UI'; padding: 8px; border-radius: 5px;")
        self.analyze_button.clicked.connect(self.analyze_pdfs)
        button_layout.addWidget(self.browse_button)
        button_layout.addWidget(self.analyze_button)
        file_layout.addLayout(button_layout)
        self.status_label = QLabel("Ready")
        self.status_label.setStyleSheet("font: italic 12px 'Segoe UI'; color: #2b3e50;")
        file_layout.addWidget(self.status_label)
        self.layout.addWidget(file_widget)

        # Processing log section
        log_widget = QWidget()
        log_layout = QVBoxLayout(log_widget)
        log_layout.setContentsMargins(10, 10, 10, 10)
        log_widget.setStyleSheet("background-color: #e9ecef; border-radius: 5px;")
        log_label = QLabel("Processing Log")
        log_label.setStyleSheet("font: bold 14px 'Segoe UI'; color: #2b3e50;")
        log_layout.addWidget(log_label)
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setStyleSheet("background-color: white; font: 12px 'Segoe UI'; border: 1px solid #ccc; border-radius: 5px;")
        log_layout.addWidget(self.log_text)
        self.layout.addWidget(log_widget)

        # Summary table section
        summary_widget = QWidget()
        summary_layout = QVBoxLayout(summary_widget)
        summary_layout.setContentsMargins(10, 10, 10, 10)
        summary_widget.setStyleSheet("background-color: #e9ecef; border-radius: 5px;")
        summary_label = QLabel("Summary Table")
        summary_label.setStyleSheet("font: bold 14px 'Segoe UI'; color: #2b3e50;")
        summary_layout.addWidget(summary_label)
        self.summary_table = QTableWidget()
        self.summary_table.setColumnCount(5)
        self.summary_table.setHorizontalHeaderLabels(["File", "Total Pages", "Blank Pages", "Gibberish Pages", "Billable Pages"])
        self.summary_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.summary_table.setStyleSheet("background-color: white; font: 12px 'Segoe UI'; alternate-background-color: #f4f6f9;")
        self.summary_table.horizontalHeader().setStyleSheet("QHeaderView::section { background-color: #3b5998; color: white; font: bold 12px 'Segoe UI'; border: 1px solid #2b3e50; padding: 4px; }")
        self.summary_table.setSortingEnabled(True)  # Enable built-in sorting
        self.summary_table.horizontalHeader().setSectionsClickable(True)  # Ensure headers are clickable
        self.summary_table.setEditTriggers(QTableWidget.NoEditTriggers)  # Disable editing
        summary_scroll = QScrollArea()
        summary_scroll.setWidget(self.summary_table)
        summary_scroll.setWidgetResizable(True)
        summary_layout.addWidget(summary_scroll)
        self.layout.addWidget(summary_widget)

        # Results table section
        result_widget = QWidget()
        result_layout = QVBoxLayout(result_widget)
        result_layout.setContentsMargins(10, 10, 10, 10)
        result_widget.setStyleSheet("background-color: #e9ecef; border-radius: 5px;")
        result_label = QLabel("Results Table")
        result_label.setStyleSheet("font: bold 14px 'Segoe UI'; color: #2b3e50;")
        result_layout.addWidget(result_label)
        self.result_table = QTableWidget()
        self.result_table.setColumnCount(4)
        self.result_table.setHorizontalHeaderLabels(["File", "Page Number", "Status", "Text Length"])
        self.result_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.result_table.setStyleSheet("background-color: white; font: 12px 'Segoe UI'; alternate-background-color: #f4f6f9;")
        self.result_table.horizontalHeader().setStyleSheet("QHeaderView::section { background-color: #3b5998; color: white; font: bold 12px 'Segoe UI'; border: 1px solid #2b3e50; padding: 4px; }")
        self.result_table.setSortingEnabled(True)  # Enable built-in sorting
        self.result_table.horizontalHeader().setSectionsClickable(True)  # Ensure headers are clickable
        self.result_table.setEditTriggers(QTableWidget.NoEditTriggers)  # Disable editing
        result_scroll = QScrollArea()
        result_scroll.setWidget(self.result_table)
        result_scroll.setWidgetResizable(True)
        result_layout.addWidget(result_scroll)
        self.layout.addWidget(result_widget)

        self.selected_files = []

    def browse_files(self):
        """Open file dialog to select multiple PDFs."""
        files, _ = QFileDialog.getOpenFileNames(self, "Select PDF Files", "", "PDF Files (*.pdf)")
        if files:
            self.selected_files = files
            self.file_label.setText(f"Selected PDFs: {len(files)} file(s)")
            self.log_text.clear()
            self.log_text.append(f"Selected {len(files)} PDF(s)")
            self.summary_table.setRowCount(0)
            self.result_table.setRowCount(0)
            self.status_label.setText("Ready")

    def analyze_pdfs(self):
        """Analyze all selected PDFs and display results."""
        if not self.selected_files:
            QMessageBox.critical(self, "Error", "Please select at least one PDF file.")
            return

        self.log_text.clear()
        self.summary_table.setRowCount(0)
        self.result_table.setRowCount(0)
        self.status_label.setText("Processing...")

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        csv_file = f"pdf_analysis_{timestamp}.csv"
        csv_data = []

        for idx, pdf_path in enumerate(self.selected_files, 1):
            self.log_text.append(f"\nAnalyzing PDF {idx}/{len(self.selected_files)}: {os.path.basename(pdf_path)}")
            self.status_label.setText(f"Processing file {idx}/{len(self.selected_files)}")

            if not os.path.exists(pdf_path):
                self.log_text.append(f"Error: File {pdf_path} not found.")
                continue

            result = analyze_pdf(pdf_path, self.log_text)
            if result:
                self.log_text.append(f"\nSummary for {os.path.basename(pdf_path)}:")
                self.log_text.append(f"Total Pages: {result['total_pages']}")
                self.log_text.append(f"Blank Pages: {result['blank_pages']}")
                self.log_text.append(f"Gibberish Pages: {result['gibberish_pages']}")
                self.log_text.append(f"Billable Pages: {result['billable_pages']}")

                # Populate Summary Table
                row = self.summary_table.rowCount()
                self.summary_table.insertRow(row)
                self.summary_table.setItem(row, 0, QTableWidgetItem(os.path.basename(pdf_path)))
                self.summary_table.setItem(row, 1, QTableWidgetItem(str(result['total_pages'])))
                self.summary_table.setItem(row, 2, QTableWidgetItem(str(result['blank_pages'])))
                self.summary_table.setItem(row, 3, QTableWidgetItem(str(result['gibberish_pages'])))
                self.summary_table.setItem(row, 4, QTableWidgetItem(str(result['billable_pages'])))

                # Populate Results Table
                for page_num, status, text_length in result['page_details']:
                    row = self.result_table.rowCount()
                    self.result_table.insertRow(row)
                    self.result_table.setItem(row, 0, QTableWidgetItem(os.path.basename(pdf_path)))
                    self.result_table.setItem(row, 1, QTableWidgetItem(str(page_num)))
                    self.result_table.setItem(row, 2, QTableWidgetItem(status))
                    self.result_table.setItem(row, 3, QTableWidgetItem(str(text_length)))

                # Add to CSV
                csv_data.append({
                    "File": os.path.basename(pdf_path),
                    "Total Pages": result['total_pages'],
                    "Blank Pages": result['blank_pages'],
                    "Gibberish Pages": result['gibberish_pages'],
                    "Billable Pages": result['billable_pages']
                })
                for page_num, status, text_length in result['page_details']:
                    csv_data.append({
                        "File": os.path.basename(pdf_path),
                        "Page Number": page_num,
                        "Status": status,
                        "Text Length": text_length
                    })

        # Write CSV
        if csv_data:
            with open(csv_file, 'w', newline='') as f:
                writer = csv.DictWriter(f, fieldnames=["File", "Total Pages", "Blank Pages", "Gibberish Pages", "Billable Pages", "Page Number", "Status", "Text Length"])
                writer.writeheader()
                for row in csv_data:
                    writer.writerow({k: v for k, v in row.items() if k in writer.fieldnames})
            self.log_text.append(f"\nResults saved to {csv_file}")

        self.status_label.setText("Analysis Complete")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PDFAnalyzerApp()
    window.show()
    sys.exit(app.exec_())