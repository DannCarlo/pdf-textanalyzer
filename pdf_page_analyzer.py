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
from PyQt5.QtCore import Qt, QThread, QObject, pyqtSignal

# Dynamically set the path to Tesseract binary in the bundled app
if getattr(sys, 'frozen', False):
    bundle_dir = sys._MEIPASS
    tesseract_path = os.path.join(bundle_dir, 'tesseract', 'tesseract.exe' if sys.platform == 'win32' else 'tesseract')
    pytesseract.pytesseract.tesseract_cmd = tesseract_path
    os.environ['TESSDATA_PREFIX'] = os.path.join(bundle_dir, 'tessdata')
else:
    pytesseract.pytesseract.tesseract_cmd = 'tesseract'

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

def is_likely_blank_pixmap(pix, whiteness_threshold=0.98):
    """Check if a pixmap is mostly white (likely blank) by sampling pixels."""
    if pix.n != 3:  # Ensure RGB format
        pix = pix.convert_to_rgb()
    pixels = pix.samples
    total_pixels = pix.width * pix.height
    white_pixels = sum(1 for i in range(0, len(pixels), 3) if pixels[i] > 240 and pixels[i+1] > 240 and pixels[i+2] > 240)
    return (white_pixels / total_pixels) > whiteness_threshold

def extract_text_with_pymupdf_image(page, page_num, log_callback):
    """Extract text from a page image using PyMuPDF and OCR."""
    try:
        pix = page.get_pixmap(dpi=150)
        if is_likely_blank_pixmap(pix):
            log_callback(f"Page {page_num} appears blank based on image analysis, skipping OCR")
            return ""
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.rgb)
        text = pytesseract.image_to_string(img, lang='eng')
        log_callback(f"OCR text length for page {page_num}: {len(text)} characters")
        logger.debug(f"OCR text sample for page {page_num}: {text[:50]}...")
        pix = None
        img.close()
        del img
        gc.collect()
        return text
    except Exception as e:
        logger.error(f"PyMuPDF image OCR failed for page {page_num}: {e}")
        log_callback(f"PyMuPDF image OCR failed for page {page_num}: {e}")
        return ""

class AnalysisWorker(QObject):
    log_signal = pyqtSignal(str)
    status_signal = pyqtSignal(str)
    result_signal = pyqtSignal(dict)
    finished_signal = pyqtSignal()

    def __init__(self, pdf_files):
        super().__init__()
        self.pdf_files = pdf_files
        self.csv_data = []

    def analyze_pdf(self, pdf_path):
        blank_pages = 0
        gibberish_pages = 0
        billable_pages = 0
        page_details = []

        try:
            doc = fitz.open(pdf_path)
            total_pages = len(doc)
            self.log_signal.emit(f"Total pages: {total_pages}")

            for i in range(total_pages):
                page_num = i + 1
                self.log_signal.emit(f"Processing page {page_num}/{total_pages}...")
                try:
                    page = doc[i]
                    text = page.get_text("text") or ""
                    self.log_signal.emit(f"PyMuPDF text length for page {page_num}: {len(text)} characters")
                except Exception as e:
                    logger.warning(f"Text extraction failed for page {page_num}: {e}")
                    text = ""

                if not text.strip():
                    self.log_signal.emit(f"Attempting OCR for page {page_num}")
                    text = extract_text_with_pymupdf_image(page, page_num, self.log_signal.emit)

                if is_blank_page(text):
                    status = "Blank"
                    blank_pages += 1
                elif is_gibberish_page(text):
                    status = "Gibberish"
                    gibberish_pages += 1
                else:
                    status = "Billable"
                    billable_pages += 1
                self.log_signal.emit(f"Page {page_num}: {status}")
                page_details.append((page_num, status, len(text)))

            doc.close()

            result = {
                "file": os.path.basename(pdf_path),
                "total_pages": total_pages,
                "blank_pages": blank_pages,
                "gibberish_pages": gibberish_pages,
                "billable_pages": billable_pages,
                "page_details": page_details
            }
            self.csv_data.append({
                "File": os.path.basename(pdf_path),
                "Total Pages": total_pages,
                "Blank Pages": blank_pages,
                "Gibberish Pages": gibberish_pages,
                "Billable Pages": billable_pages
            })
            for page_num, status, text_length in page_details:
                self.csv_data.append({
                    "File": os.path.basename(pdf_path),
                    "Page Number": page_num,
                    "Status": status,
                    "Text Length": text_length
                })
            return result

        except Exception as e:
            logger.error(f"Error processing PDF: {e}")
            self.log_signal.emit(f"Error processing PDF: {e}")
            return None

    def run(self):
        for idx, pdf_path in enumerate(self.pdf_files, 1):
            self.log_signal.emit(f"\nAnalyzing PDF {idx}/{len(self.pdf_files)}: {os.path.basename(pdf_path)}")
            self.status_signal.emit(f"Processing file {idx}/{len(self.pdf_files)}")

            if not os.path.exists(pdf_path):
                self.log_signal.emit(f"Error: File {pdf_path} not found.")
                continue

            result = self.analyze_pdf(pdf_path)
            if result:
                self.log_signal.emit(f"\nSummary for {result['file']}:")
                self.log_signal.emit(f"Total Pages: {result['total_pages']}")
                self.log_signal.emit(f"Blank Pages: {result['blank_pages']}")
                self.log_signal.emit(f"Gibberish Pages: {result['gibberish_pages']}")
                self.log_signal.emit(f"Billable Pages: {result['billable_pages']}")
                self.result_signal.emit(result)

        if self.csv_data:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            csv_file = f"pdf_analysis_{timestamp}.csv"
            with open(csv_file, 'w', newline='') as f:
                writer = csv.DictWriter(f, fieldnames=["File", "Total Pages", "Blank Pages", "Gibberish Pages", "Billable Pages", "Page Number", "Status", "Text Length"])
                writer.writeheader()
                for row in self.csv_data:
                    writer.writerow({k: v for k, v in row.items() if k in writer.fieldnames})
            self.log_signal.emit(f"\nResults saved to {csv_file}")

        self.status_signal.emit("Analysis Complete")
        self.finished_signal.emit()

class PDFAnalyzerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF Page Analyzer")
        self.setGeometry(100, 100, 800, 600)

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.layout = QVBoxLayout(self.central_widget)
        self.layout.setSpacing(10)

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
        self.summary_table.setSortingEnabled(True)
        self.summary_table.horizontalHeader().setSectionsClickable(True)
        self.summary_table.setEditTriggers(QTableWidget.NoEditTriggers)
        summary_scroll = QScrollArea()
        summary_scroll.setWidget(self.summary_table)
        summary_scroll.setWidgetResizable(True)
        summary_layout.addWidget(summary_scroll)
        self.layout.addWidget(summary_widget)

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
        self.result_table.setSortingEnabled(True)
        self.result_table.horizontalHeader().setSectionsClickable(True)
        self.result_table.setEditTriggers(QTableWidget.NoEditTriggers)
        result_scroll = QScrollArea()
        result_scroll.setWidget(self.result_table)
        result_scroll.setWidgetResizable(True)
        result_layout.addWidget(result_scroll)
        self.layout.addWidget(result_widget)

        self.selected_files = []
        self.thread = None
        self.worker = None

    def browse_files(self):
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
        if not self.selected_files:
            QMessageBox.critical(self, "Error", "Please select at least one PDF file.")
            return

        self.log_text.clear()
        self.summary_table.setRowCount(0)
        self.result_table.setRowCount(0)
        self.status_label.setText("Processing...")
        self.analyze_button.setEnabled(False)

        self.thread = QThread()
        self.worker = AnalysisWorker(self.selected_files)
        self.worker.moveToThread(self.thread)

        self.worker.log_signal.connect(self.log_text.append)
        self.worker.status_signal.connect(self.status_label.setText)
        self.worker.result_signal.connect(self.update_tables)
        self.worker.finished_signal.connect(self.on_analysis_finished)
        self.thread.started.connect(self.worker.run)

        self.thread.start()

    def update_tables(self, result):
        # Populate Summary Table
        row = self.summary_table.rowCount()
        self.summary_table.insertRow(row)
        self.summary_table.setItem(row, 0, QTableWidgetItem(str(result['file'])))
        self.summary_table.setItem(row, 1, QTableWidgetItem(str(result['total_pages'])))
        self.summary_table.setItem(row, 2, QTableWidgetItem(str(result['blank_pages'])))
        self.summary_table.setItem(row, 3, QTableWidgetItem(str(result['gibberish_pages'])))
        self.summary_table.setItem(row, 4, QTableWidgetItem(str(result['billable_pages'])))

        # Populate Results Table
        for page_num, status, text_length in result['page_details']:
            row = self.result_table.rowCount()
            self.result_table.insertRow(row)

            # File
            file_item = QTableWidgetItem(str(result['file']))
            self.result_table.setItem(row, 0, file_item)

            # Page Number (set Qt.UserRole for numeric sorting)
            page_item = QTableWidgetItem(str(page_num))
            page_item.setData(Qt.UserRole, int(page_num))
            self.result_table.setItem(row, 1, page_item)

            # Status
            status_item = QTableWidgetItem(str(status))
            self.result_table.setItem(row, 2, status_item)

            # Text Length (set Qt.UserRole for numeric sorting)
            length_item = QTableWidgetItem(str(text_length))
            length_item.setData(Qt.UserRole, int(text_length))
            self.result_table.setItem(row, 3, length_item)

        # Force table repaint
        self.result_table.viewport().update()

    def on_analysis_finished(self):
        self.analyze_button.setEnabled(True)
        self.thread.quit()
        self.thread.wait()
        self.thread = None
        self.worker = None

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PDFAnalyzerApp()
    window.show()
    sys.exit(app.exec_())