import sys
import os
import fitz  # PyMuPDF
import enchant
import re
import logging
from PIL import Image
import pytesseract
import gc
import io
import csv
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                             QPushButton, QLabel, QFileDialog, QTableWidget, QTableWidgetItem,
                             QTextEdit, QLineEdit, QHeaderView, QMessageBox, QProgressBar)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QDoubleValidator, QFont

# Constants
BLANK_PAGE_CHAR_THRESHOLD = 100
VALID_WORD_THRESHOLD = 0.1
WHITENESS_THRESHOLD = 0.98
MIN_IMAGE_SIZE_BYTES = 1024
MAX_IMAGE_DIMENSION = 1000
PIXEL_SAMPLING_STRIDE = 1000
OCR_DPI = 100
TESSERACT_PSM_CONFIG = '--psm 6'

# Configure Tesseract path
if getattr(sys, 'frozen', False):
    bundle_dir = sys._MEIPASS
    tesseract_path = os.path.join(bundle_dir, 'tesseract', 'tesseract.exe' if sys.platform == 'win32' else 'tesseract')
    pytesseract.pytesseract.tesseract_cmd = tesseract_path
    os.environ['TESSDATA_PREFIX'] = os.path.join(bundle_dir, 'tessdata')
else:
    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(message)s')
logger = logging.getLogger(__name__)

# Initialize dictionary once for reuse
dictionary = enchant.Dict("en_US")


# Custom logging handler to emit logs to QTextEdit
class QTextEditLogger(logging.Handler):
    def __init__(self, text_edit):
        super().__init__()
        self.text_edit = text_edit

    def emit(self, record):
        msg = self.format(record)
        self.text_edit.append(msg)


# Thread for running PDF analysis
class AnalysisThread(QThread):
    update_progress = pyqtSignal(str, int, int)  # message, current, total
    log_message = pyqtSignal(str)
    analysis_complete = pyqtSignal(list)  # List of results for multiple PDFs
    analysis_failed = pyqtSignal(str)

    def __init__(self, pdf_paths, blank_threshold, valid_word_threshold):
        super().__init__()
        self.pdf_paths = pdf_paths  # List of PDF paths
        self.blank_threshold = blank_threshold
        self.valid_word_threshold = valid_word_threshold
        self.is_running = True

    def run(self):
        try:
            all_results = []
            total_files = len(self.pdf_paths)
            for file_idx, pdf_path in enumerate(self.pdf_paths):
                if not self.is_running:
                    break
                self.log_message.emit(f"\nProcessing file: {os.path.basename(pdf_path)}")
                result = self.analyze_pdf(pdf_path)
                if result:
                    # Add filename to page details for Results Table
                    filename = os.path.basename(pdf_path)
                    result['filename'] = filename
                    for detail in result['page_details']:
                        detail.insert(0, filename)  # Add filename to each page detail
                    all_results.append(result)
                # Update progress for file completion
                self.update_progress.emit(f"Processing {filename}", file_idx + 1, total_files)

            if self.is_running:
                if all_results:
                    self.analysis_complete.emit(all_results)
                else:
                    self.analysis_failed.emit("Analysis failed. Check the log for details.")
            else:
                self.analysis_failed.emit("Analysis cancelled.")
        except Exception as e:
            self.analysis_failed.emit(f"Error: {str(e)}")

    def stop(self):
        self.is_running = False

    def is_blank_page(self, text, char_threshold):
        return len(text.strip()) < char_threshold

    def is_gibberish_page(self, text, valid_word_threshold):
        if not text.strip():
            return False
        words = re.findall(r'\b\w+\b', text.lower())
        if not words:
            return True
        valid_words = sum(1 for word in words if dictionary.check(word) and len(word) > 1)
        valid_ratio = valid_words / len(words)
        return valid_ratio < valid_word_threshold

    def is_likely_blank_pixmap(self, pix, whiteness_threshold=WHITENESS_THRESHOLD):
        if pix.n != 3:
            pix = pix.convert_to_rgb()
        pixels = pix.samples
        total_pixels = pix.width * pix.height
        stride = max(1, total_pixels // PIXEL_SAMPLING_STRIDE)
        white_pixels = sum(
            1 for i in range(0, len(pixels), stride * 3)
            if pixels[i] > 240 and pixels[i + 1] > 240 and pixels[i + 2] > 240
        )
        return (white_pixels / (total_pixels / stride)) > whiteness_threshold

    def extract_text_with_pymupdf_image(self, page, page_num):
        try:
            pix = page.get_pixmap(dpi=OCR_DPI)
            if self.is_likely_blank_pixmap(pix):
                return ""
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.rgb)
            text = pytesseract.image_to_string(img, lang='eng', config=TESSERACT_PSM_CONFIG)
            pix = None
            img.close()
            del img
            gc.collect()
            return text
        except Exception as e:
            self.log_message.emit(f"PyMuPDF image OCR failed for page {page_num}: {e}")
            return ""

    def extract_images_from_page(self, page):
        image_list = []
        try:
            image_list.extend(page.get_images(full=True))
        except Exception as e:
            self.log_message.emit(f"Failed to extract images from page: {e}")
        return image_list

    def analyze_image(self, image_info, page_num, doc):
        try:
            xref = image_info[0]
            base_image = doc.extract_image(xref)
            image_bytes = base_image["image"]
            image_ext = base_image.get("ext", "").lower()
            if image_ext not in ["png", "jpeg", "jpg", "bmp", "tiff"]:
                self.log_message.emit(f"Unsupported image format '{image_ext}' on page {page_num}, skipping")
                return ""
            if len(image_bytes) < MIN_IMAGE_SIZE_BYTES:
                return ""
            img = Image.open(io.BytesIO(image_bytes))
            if img.width > MAX_IMAGE_DIMENSION or img.height > MAX_IMAGE_DIMENSION:
                img = img.resize((min(img.width, MAX_IMAGE_DIMENSION), min(img.height, MAX_IMAGE_DIMENSION)),
                                 Image.Resampling.LANCZOS)
            if img.mode != "RGB":
                img = img.convert("RGB")
            text = pytesseract.image_to_string(img, lang='eng', config=TESSERACT_PSM_CONFIG)
            img.close()
            del img, base_image
            gc.collect()
            return text
        except Exception as e:
            self.log_message.emit(f"Image analysis failed for page {page_num}: {e}")
            return ""

    def analyze_pdf(self, pdf_path):
        blank_pages = 0
        gibberish_pages = 0
        billable_pages = 0
        page_details = []

        try:
            doc = fitz.open(pdf_path)
            total_pages = len(doc)

            for i in range(total_pages):
                if not self.is_running:
                    doc.close()
                    return None

                page_num = i + 1
                self.update_progress.emit(f"Analyzing page {page_num}/{total_pages}", i + 1, total_pages)

                try:
                    page = doc[i]
                    text = page.get_text("text") or ""
                except Exception as e:
                    self.log_message.emit(f"Text extraction failed for page {page_num}: {e}")
                    text = ""

                if len(text.strip()) >= self.blank_threshold and not self.is_gibberish_page(text,
                                                                                            self.valid_word_threshold):
                    status = "Billable"
                    billable_pages += 1
                    self.log_message.emit(
                        f"Page {page_num}: {status}\n"
                        f"Text Length: {len(text)}\n"
                        f"Combined Text Length: {len(text)}"
                    )
                    page_details.append([page_num, status, len(text)])
                    continue

                if not text.strip():
                    text = self.extract_text_with_pymupdf_image(page, page_num)

                image_text = ""
                image_list = self.extract_images_from_page(page)
                if image_list:
                    for img_info in image_list:
                        if not self.is_running:
                            doc.close()
                            return None
                        image_text += self.analyze_image(img_info, page_num, doc) + " "

                combined_text = text + " " + image_text

                if self.is_blank_page(combined_text, self.blank_threshold):
                    status = "Blank"
                    blank_pages += 1
                elif self.is_gibberish_page(combined_text, self.valid_word_threshold):
                    status = "Gibberish"
                    gibberish_pages += 1
                else:
                    status = "Billable"
                    billable_pages += 1

                self.log_message.emit(
                    f"Page {page_num}: {status}\n"
                    f"Text Length: {len(text)}\n"
                    f"Combined Text Length: {len(combined_text)}"
                )
                page_details.append([page_num, status, len(text)])

            doc.close()
            del doc
            gc.collect()

        except Exception as e:
            self.log_message.emit(f"Error processing PDF: {e}")
            return None

        return {
            "total_pages": total_pages,
            "blank_pages": blank_pages,
            "gibberish_pages": gibberish_pages,
            "billable_pages": billable_pages,
            "page_details": page_details
        }


# Main GUI Window
class PDFAnalyzerGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF Analyzer")
        self.setGeometry(100, 100, 900, 700)

        self.pdf_paths = []  # List of selected PDF paths
        self.thread = None
        self.all_page_details = []  # Store all page details for export
        self.init_ui()

        # Apply modern stylesheet with new colors
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f5f7fa;
            }
            QPushButton {
                background-color: #ff7b1c;
                color: white;
                border-radius: 5px;
                padding: 8px;
                font-size: 14px;
                border: none;
            }
            QPushButton:hover {
                background-color: #d45b04;
            }
            QPushButton:disabled {
                background-color: #ffb07c;
            }
            QLineEdit {
                border: 1px solid #dcdcdc;
                border-radius: 5px;
                padding: 5px;
                font-size: 14px;
                background-color: white;
            }
            QTableWidget {
                border: 1px solid #dcdcdc;
                border-radius: 5px;
                background-color: white;
                font-size: 14px;
            }
            QTableWidget::item {
                padding: 5px;
            }
            QHeaderView::section {
                background-color: #ff7b1c;
                color: white;
                padding: 5px;
                border: none;
            }
            QHeaderView::section:hover {
                background-color: #d45b04;
            }
            QTextEdit {
                border: 1px solid #dcdcdc;
                border-radius: 5px;
                background-color: white;
                font-size: 14px;
            }
            QLabel {
                font-size: 14px;
                color: #333;
            }
            QProgressBar {
                border: 1px solid #dcdcdc;
                border-radius: 5px;
                text-align: center;
                font-size: 14px;
            }
            QProgressBar::chunk {
                background-color: #ff7b1c;
                border-radius: 3px;
            }
            QProgressBar::chunk:disabled {
                background-color: #ffb07c;
            }
        """)

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        layout.setSpacing(10)
        layout.setContentsMargins(15, 15, 15, 15)

        # File selection
        file_layout = QHBoxLayout()
        self.file_label = QLabel("No files selected")
        file_button = QPushButton("Select PDF Files")
        file_button.clicked.connect(self.select_files)
        file_layout.addWidget(self.file_label)
        file_layout.addWidget(file_button)
        layout.addLayout(file_layout)

        # Threshold inputs
        threshold_layout = QHBoxLayout()
        blank_label = QLabel("Blank Page Char Threshold:")
        self.blank_threshold_input = QLineEdit(str(BLANK_PAGE_CHAR_THRESHOLD))
        self.blank_threshold_input.setValidator(QDoubleValidator(0, 10000, 2))
        self.blank_threshold_input.setFixedWidth(100)
        valid_word_label = QLabel("Valid Word Threshold:")
        self.valid_word_threshold_input = QLineEdit(str(VALID_WORD_THRESHOLD))
        self.valid_word_threshold_input.setValidator(QDoubleValidator(0, 1, 2))
        self.valid_word_threshold_input.setFixedWidth(100)
        threshold_layout.addWidget(blank_label)
        threshold_layout.addWidget(self.blank_threshold_input)
        threshold_layout.addWidget(valid_word_label)
        threshold_layout.addWidget(self.valid_word_threshold_input)
        threshold_layout.addStretch()
        layout.addLayout(threshold_layout)

        # Start and Cancel buttons
        button_layout = QHBoxLayout()
        self.start_button = QPushButton("Start Analysis")
        self.start_button.clicked.connect(self.start_analysis)
        self.start_button.setEnabled(False)
        self.cancel_button = QPushButton("Cancel")
        self.cancel_button.clicked.connect(self.cancel_analysis)
        self.cancel_button.setEnabled(False)
        button_layout.addWidget(self.start_button)
        button_layout.addWidget(self.cancel_button)
        layout.addLayout(button_layout)

        # Status label and Progress bar
        self.status_label = QLabel("Status: Idle")
        layout.addWidget(self.status_label)

        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        layout.addWidget(self.progress_bar)

        # Summary Table
        self.summary_table = QTableWidget()
        self.summary_table.setRowCount(1)
        self.summary_table.setColumnCount(4)
        self.summary_table.setHorizontalHeaderLabels(
            ["Total Pages", "Blank Pages", "Gibberish Pages", "Billable Pages"])
        self.summary_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        layout.addWidget(self.summary_table)

        # Results Table
        self.results_table = QTableWidget()
        self.results_table.setColumnCount(4)  # Added File column
        self.results_table.setHorizontalHeaderLabels(["File", "Page Number", "Status", "Text Length"])
        self.results_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.results_table.setSortingEnabled(True)
        layout.addWidget(self.results_table)

        # Export to CSV button
        self.export_button = QPushButton("Export to CSV")
        self.export_button.clicked.connect(self.export_to_csv)
        self.export_button.setEnabled(False)
        layout.addWidget(self.export_button)

        # Log viewer
        self.log_viewer = QTextEdit()
        self.log_viewer.setReadOnly(True)
        self.log_viewer.setMinimumHeight(150)
        layout.addWidget(self.log_viewer)

        # Configure logging to display in log viewer
        handler = QTextEditLogger(self.log_viewer)
        handler.setFormatter(logging.Formatter('%(message)s'))
        global logger
        logger.handlers = []  # Clear existing handlers to avoid duplicates
        logger.addHandler(handler)

    def select_files(self):
        file_paths, _ = QFileDialog.getOpenFileNames(self, "Select PDF Files", "", "PDF Files (*.pdf)")
        if file_paths:
            self.pdf_paths = file_paths
            self.file_label.setText(f"{len(file_paths)} file(s) selected")
            self.start_button.setEnabled(True)
            self.status_label.setText("Status: Files selected, ready to analyze")

    def start_analysis(self):
        if not self.pdf_paths:
            QMessageBox.warning(self, "Error", "Please select at least one PDF file.")
            return

        try:
            blank_threshold = float(self.blank_threshold_input.text())
            valid_word_threshold = float(self.valid_word_threshold_input.text())
        except ValueError:
            QMessageBox.warning(self, "Error", "Invalid threshold values. Please enter numeric values.")
            return

        self.start_button.setEnabled(False)
        self.cancel_button.setEnabled(True)
        self.export_button.setEnabled(False)
        self.status_label.setText("Status: Analyzing...")
        self.summary_table.clearContents()
        self.results_table.setRowCount(0)
        self.log_viewer.clear()
        self.all_page_details = []
        self.progress_bar.setValue(0)

        self.thread = AnalysisThread(self.pdf_paths, blank_threshold, valid_word_threshold)
        self.thread.update_progress.connect(self.update_progress)
        self.thread.log_message.connect(self.log_viewer.append)
        self.thread.analysis_complete.connect(self.on_analysis_complete)
        self.thread.analysis_failed.connect(self.on_analysis_failed)
        self.thread.start()

    def update_progress(self, message, current, total):
        self.status_label.setText(f"Status: {message}")
        self.progress_bar.setMaximum(total)
        self.progress_bar.setValue(current)

    def cancel_analysis(self):
        if self.thread:
            self.thread.stop()
            self.status_label.setText("Status: Analysis cancelled")
            self.start_button.setEnabled(True)
            self.cancel_button.setEnabled(False)
            self.progress_bar.setValue(0)

    def on_analysis_complete(self, results):
        self.start_button.setEnabled(True)
        self.cancel_button.setEnabled(False)
        self.export_button.setEnabled(True)
        self.status_label.setText("Status: Analysis complete")
        self.progress_bar.setValue(self.progress_bar.maximum())

        # Aggregate summary
        total_pages = sum(result['total_pages'] for result in results)
        blank_pages = sum(result['blank_pages'] for result in results)
        gibberish_pages = sum(result['gibberish_pages'] for result in results)
        billable_pages = sum(result['billable_pages'] for result in results)

        # Update Summary Table
        self.summary_table.setItem(0, 0, QTableWidgetItem(str(total_pages)))
        self.summary_table.setItem(0, 1, QTableWidgetItem(str(blank_pages)))
        self.summary_table.setItem(0, 2, QTableWidgetItem(str(gibberish_pages)))
        self.summary_table.setItem(0, 3, QTableWidgetItem(str(billable_pages)))

        # Update Results Table
        row = 0
        for result in results:
            for page_detail in result['page_details']:
                self.results_table.setRowCount(row + 1)
                # page_detail = [filename, page_num, status, text_length]
                self.results_table.setItem(row, 0, QTableWidgetItem(page_detail[0]))  # File
                self.results_table.setItem(row, 1, QTableWidgetItem(str(page_detail[1])))  # Page Number
                self.results_table.setItem(row, 2, QTableWidgetItem(page_detail[2]))  # Status
                self.results_table.setItem(row, 3, QTableWidgetItem(str(page_detail[3])))  # Text Length
                row += 1
            self.all_page_details.extend(result['page_details'])

        # Log final summary
        self.log_viewer.append(f"\nFinal Summary:")
        self.log_viewer.append(f"Total Pages: {total_pages}")
        self.log_viewer.append(f"Blank Pages: {blank_pages}")
        self.log_viewer.append(f"Gibberish Pages: {gibberish_pages}")
        self.log_viewer.append(f"Billable Pages: {billable_pages}")

    def on_analysis_failed(self, message):
        self.start_button.setEnabled(True)
        self.cancel_button.setEnabled(False)
        self.status_label.setText("Status: Analysis failed")
        self.progress_bar.setValue(0)
        QMessageBox.critical(self, "Error", message)

    def export_to_csv(self):
        if not self.all_page_details:
            QMessageBox.warning(self, "Error", "No analysis data to export.")
            return

        file_path, _ = QFileDialog.getSaveFileName(self, "Save CSV File", "", "CSV Files (*.csv)")
        if file_path:
            try:
                with open(file_path, 'w', newline='') as csvfile:
                    writer = csv.writer(csvfile)
                    writer.writerow(["File", "Page Number", "Status", "Text Length"])
                    for detail in self.all_page_details:
                        writer.writerow(detail)
                QMessageBox.information(self, "Success", f"Results exported to {file_path}")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to export CSV: {str(e)}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PDFAnalyzerGUI()
    window.show()
    sys.exit(app.exec_())