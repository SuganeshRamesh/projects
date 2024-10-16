import sys
import fitz  # PyMuPDF
import numpy as np
from PIL import Image, ImageDraw, ImageFont
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                           QHBoxLayout, QLabel, QPushButton, QFileDialog, 
                           QScrollArea, QMessageBox, QSpinBox, QCheckBox,
                           QColorDialog, QComboBox, QGroupBox, QSlider)
from PyQt6.QtGui import QPixmap, QImage, QColor
from PyQt6.QtCore import Qt
import cv2
from difflib import SequenceMatcher

class PDFComparisonTool(QMainWindow):
    def __init__(self):
        super().__init__()
        self.zoom_level = 1.0
        self.highlight_differences = True
        self.difference_threshold = 30
        self.insertion_color = QColor(0, 255, 0, 64)  # Transparent green
        self.deletion_color = QColor(255, 0, 0, 64)  # Transparent red
        self.comparison_mode = 'rgb'  # Default comparison mode
        self.show_difference_map = False
        self.initUI()
        
    def initUI(self):
        self.setWindowTitle('PDF Comparison Tool')
        self.setGeometry(100, 100, 1400, 900)

        # Create main widget and layout
        self.main_widget = QWidget()
        self.setCentralWidget(self.main_widget)
        self.main_layout = QVBoxLayout(self.main_widget)

        # Create control panels
        self.create_file_control_panel()
        self.create_comparison_control_panel()
        
        # Create scroll area
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.main_layout.addWidget(self.scroll_area)

        # Create widget for comparison content
        self.comparison_widget = QWidget()
        self.comparison_layout = QVBoxLayout(self.comparison_widget)
        self.comparison_layout.setSpacing(20)
        self.scroll_area.setWidget(self.comparison_widget)

        # Initialize paths and status bar
        self.pdf1_path = ''
        self.pdf2_path = ''
        self.statusBar().showMessage('Ready')

    def create_file_control_panel(self):
        file_panel = QGroupBox("File Controls")
        file_layout = QHBoxLayout()
        
        # File selection buttons
        self.pdf1_button = QPushButton('Select PDF 1')
        self.pdf2_button = QPushButton('Select PDF 2')
        self.compare_button = QPushButton('Compare PDFs')
        
        # Zoom controls
        zoom_widget = QWidget()
        zoom_layout = QHBoxLayout(zoom_widget)
        zoom_layout.setContentsMargins(0, 0, 0, 0)
        
        zoom_out_btn = QPushButton('-')
        zoom_out_btn.setFixedWidth(30)
        self.zoom_spin = QSpinBox()
        self.zoom_spin.setRange(25, 400)
        self.zoom_spin.setValue(100)
        self.zoom_spin.setSuffix('%')
        zoom_in_btn = QPushButton('+')
        zoom_in_btn.setFixedWidth(30)
        
        zoom_layout.addWidget(zoom_out_btn)
        zoom_layout.addWidget(self.zoom_spin)
        zoom_layout.addWidget(zoom_in_btn)
        
        # Add controls to panel
        file_layout.addWidget(self.pdf1_button)
        file_layout.addWidget(self.pdf2_button)
        file_layout.addWidget(self.compare_button)
        file_layout.addStretch()
        file_layout.addWidget(QLabel('Zoom:'))
        file_layout.addWidget(zoom_widget)
        
        file_panel.setLayout(file_layout)
        self.main_layout.addWidget(file_panel)

        # Connect buttons
        self.pdf1_button.clicked.connect(lambda: self.select_pdf(1))
        self.pdf2_button.clicked.connect(lambda: self.select_pdf(2))
        self.compare_button.clicked.connect(self.compare_pdfs)
        zoom_in_btn.clicked.connect(self.zoom_in)
        zoom_out_btn.clicked.connect(self.zoom_out)
        self.zoom_spin.valueChanged.connect(self.zoom_value_changed)
    
    def create_comparison_control_panel(self):
        comparison_panel = QGroupBox("Comparison Controls")
        comparison_layout = QHBoxLayout()

        # Highlight toggle
        self.highlight_checkbox = QCheckBox('Highlight Differences')
        self.highlight_checkbox.setChecked(self.highlight_differences)
        
        # Difference map toggle
        self.diff_map_checkbox = QCheckBox('Show Difference Map')
        self.diff_map_checkbox.setChecked(self.show_difference_map)

        # Threshold control
        threshold_widget = QWidget()
        threshold_layout = QHBoxLayout(threshold_widget)
        threshold_layout.setContentsMargins(0, 0, 0, 0)
        
        threshold_label = QLabel('Threshold:')
        self.threshold_slider = QSlider(Qt.Orientation.Horizontal)
        self.threshold_slider.setRange(1, 100)
        self.threshold_slider.setValue(self.difference_threshold)
        self.threshold_value_label = QLabel(f'{self.difference_threshold}')
        
        threshold_layout.addWidget(threshold_label)
        threshold_layout.addWidget(self.threshold_slider)
        threshold_layout.addWidget(self.threshold_value_label)

        # Color selection
        self.insertion_color_button = QPushButton('Insertion Color')
        self.insertion_color_button.setStyleSheet(
            f'background-color: {self.insertion_color.name()}'
        )
        self.deletion_color_button = QPushButton('Deletion Color')
        self.deletion_color_button.setStyleSheet(
            f'background-color: {self.deletion_color.name()}'
        )

        # Comparison mode
        mode_widget = QWidget()
        mode_layout = QHBoxLayout(mode_widget)
        mode_layout.setContentsMargins(0, 0, 0, 0)
        
        mode_label = QLabel('Compare Mode:')
        self.mode_combo = QComboBox()
        self.mode_combo.addItems(['RGB', 'Grayscale', 'Text Only'])
        
        mode_layout.addWidget(mode_label)
        mode_layout.addWidget(self.mode_combo)

        # Add all controls to panel
        comparison_layout.addWidget(self.highlight_checkbox)
        comparison_layout.addWidget(self.diff_map_checkbox)
        comparison_layout.addWidget(threshold_widget)
        comparison_layout.addWidget(self.insertion_color_button)
        comparison_layout.addWidget(self.deletion_color_button)
        comparison_layout.addWidget(mode_widget)
        comparison_layout.addStretch()

        comparison_panel.setLayout(comparison_layout)
        self.main_layout.addWidget(comparison_panel)

        # Connect controls
        self.highlight_checkbox.stateChanged.connect(self.toggle_highlighting)
        self.diff_map_checkbox.stateChanged.connect(self.toggle_difference_map)
        self.threshold_slider.valueChanged.connect(self.threshold_changed)
        self.insertion_color_button.clicked.connect(self.select_insertion_color)
        self.deletion_color_button.clicked.connect(self.select_deletion_color)
        self.mode_combo.currentTextChanged.connect(self.mode_changed)

    def select_pdf(self, pdf_num):
        try:
            file_path, _ = QFileDialog.getOpenFileName(
                self, f'Select PDF {pdf_num}', '', 'PDF files (*.pdf)'
            )
            if file_path:
                if pdf_num == 1:
                    self.pdf1_path = file_path
                    self.pdf1_button.setText(f'PDF 1: {file_path.split("/")[-1]}')
                else:
                    self.pdf2_path = file_path
                    self.pdf2_button.setText(f'PDF 2: {file_path.split("/")[-1]}')
        except Exception as e:
            QMessageBox.critical(self, 'Error', f'Error selecting PDF: {str(e)}')

    def select_insertion_color(self):
        color = QColorDialog.getColor(self.insertion_color, self, "Select Insertion Highlight Color", 
                                    QColorDialog.ColorDialogOption.ShowAlphaChannel)
        if color.isValid():
            self.insertion_color = color
            self.insertion_color_button.setStyleSheet(f'background-color: {color.name()}')
            self.compare_pdfs()

    def select_deletion_color(self):
        color = QColorDialog.getColor(self.deletion_color, self, "Select Deletion Highlight Color", 
                                    QColorDialog.ColorDialogOption.ShowAlphaChannel)
        if color.isValid():
            self.deletion_color = color
            self.deletion_color_button.setStyleSheet(f'background-color: {color.name()}')
            self.compare_pdfs()

    def threshold_changed(self, value):
        self.difference_threshold = value
        self.threshold_value_label.setText(str(value))
        self.compare_pdfs()

    def mode_changed(self, mode):
        self.comparison_mode = mode.lower()
        self.compare_pdfs()

    def toggle_difference_map(self, state):
        self.show_difference_map = bool(state)
        self.compare_pdfs()

    def toggle_highlighting(self, state):
        self.highlight_differences = bool(state)
        if self.pdf1_path and self.pdf2_path:
            self.compare_pdfs()

    def zoom_in(self):
        self.zoom_spin.setValue(self.zoom_spin.value() + 10)

    def zoom_out(self):
        self.zoom_spin.setValue(self.zoom_spin.value() - 10)

    def zoom_value_changed(self, value):
        self.zoom_level = value / 100.0
        self.compare_pdfs()

    def get_text_and_positions(self, doc):
        """Extract text and their positions from all pages of a PDF document"""
        all_text = []
        for page_num in range(len(doc)):
            page = doc[page_num]
            text_instances = page.get_text("words")
            all_text.extend([(inst[4], inst[:4], page_num) for inst in text_instances])
        return all_text

    def align_content(self, text1, text2):
        """Align content from two PDFs"""
        s = SequenceMatcher(None, [t[0] for t in text1], [t[0] for t in text2])
        aligned1 = []
        aligned2 = []
        for tag, i1, i2, j1, j2 in s.get_opcodes():
            if tag == 'equal':
                aligned1.extend(text1[i1:i2])
                aligned2.extend(text2[j1:j2])
            elif tag == 'replace':
                max_len = max(i2 - i1, j2 - j1)
                aligned1.extend(text1[i1:i2] + [None] * (max_len - (i2 - i1)))
                aligned2.extend(text2[j1:j2] + [None] * (max_len - (j2 - j1)))
            elif tag == 'delete':
                aligned1.extend(text1[i1:i2])
                aligned2.extend([None] * (i2 - i1))
            elif tag == 'insert':
                aligned1.extend([None] * (j2 - j1))
                aligned2.extend(text2[j1:j2])
        return aligned1, aligned2

    def highlight_diff_areas(self, doc1, doc2):
        """Highlight differences between two PDF documents"""
        try:
            # Extract text and positions from both documents
            text1 = self.get_text_and_positions(doc1)
            text2 = self.get_text_and_positions(doc2)

            # Align content
            aligned1, aligned2 = self.align_content(text1, text2)

            # Create PIL Images for all pages
            images1 = [Image.frombytes("RGB", [page.get_pixmap(matrix=fitz.Matrix(self.zoom_level, self.zoom_level)).width, 
                                               page.get_pixmap(matrix=fitz.Matrix(self.zoom_level, self.zoom_level)).height], 
                                       page.get_pixmap(matrix=fitz.Matrix(self.zoom_level, self.zoom_level)).samples)
                       for page in doc1]
            images2 = [Image.frombytes("RGB", [page.get_pixmap(matrix=fitz.Matrix(self.zoom_level, self.zoom_level)).width, 
                                               page.get_pixmap(matrix=fitz.Matrix(self.zoom_level, self.zoom_level)).height], 
                                       page.get_pixmap(matrix=fitz.Matrix(self.zoom_level, self.zoom_level)).samples)
                       for page in doc2]

            # Create drawing objects for all pages
            draws1 = [ImageDraw.Draw(img, 'RGBA') for img in images1]
            draws2 = [ImageDraw.Draw(img, 'RGBA') for img in images2]

            # Highlight differences
            for word1, word2 in zip(aligned1, aligned2):
                if word1 is None:  # Insertion
                    word, pos, page_num = word2
                    draws2[page_num].rectangle(pos, fill=self.insertion_color.getRgb())
                elif word2 is None:  # Deletion
                    word, pos, page_num = word1
                    draws1[page_num].rectangle(pos, fill=self.deletion_color.getRgb())
                elif word1[0] != word2[0]:  # Replacement
                    word, pos, page_num = word1
                    draws1[page_num].rectangle(pos, fill=self.deletion_color.getRgb())
                    word, pos, page_num = word2
                    draws2[page_num].rectangle(pos, fill=self.insertion_color.getRgb())

            return [np.array(img) for img in images1], [np.array(img) for img in images2]
        
        except Exception as e:
            print(f"Error in highlight_diff_areas: {str(e)}")
            # Return original images if highlighting fails
            return [page.get_pixmap(matrix=fitz.Matrix(self.zoom_level, self.zoom_level)).tobytes("raw", "RGB") for page in doc1], \
                   [page.get_pixmap(matrix=fitz.Matrix(self.zoom_level, self.zoom_level)).tobytes("raw", "RGB") for page in doc2]

    def compare_pdfs(self):
        if not self.pdf1_path or not self.pdf2_path:
            QMessageBox.warning(self, 'Warning', 'Please select both PDF files first.')
            return

        try:
            # Clear previous comparison
            while self.comparison_layout.count():
                child = self.comparison_layout.takeAt(0)
                if child.widget():
                    child.widget().deleteLater()

            # Open PDF documents
            doc1 = fitz.open(self.pdf1_path)
            doc2 = fitz.open(self.pdf2_path)

            try:
                total_pages = max(len(doc1), len(doc2))
                self.statusBar().showMessage(f'Comparing {total_pages} pages...')
                
                if self.highlight_differences:
                    images1, images2 = self.highlight_diff_areas(doc1, doc2)
                else:
                    images1 = [page.get_pixmap(matrix=fitz.Matrix(self.zoom_level, self.zoom_level)).tobytes("raw", "RGB") for page in doc1]
                    images2 = [page.get_pixmap(matrix=fitz.Matrix(self.zoom_level, self.zoom_level)).tobytes("raw", "RGB") for page in doc2]

                for page_num in range(total_pages):
                    # Create container for this page pair
                    page_container = QWidget()
                    page_layout = QHBoxLayout(page_container)
                    page_layout.setSpacing(10)

                    # Add page number label
                    page_label = QLabel(f'Page {page_num + 1}')
                    page_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
                    self.comparison_layout.addWidget(page_label)

                    # Process and display pages
                    img1 = images1[page_num] if page_num < len(images1) else None
                    img2 = images2[page_num] if page_num < len(images2) else None

                    if img1 is not None:
                        img1 = self.process_image_for_comparison(img1)
                        page1_pixmap = QPixmap.fromImage(QImage(img1.data, img1.shape[1], img1.shape[0], QImage.Format.Format_RGB888))
                        label1 = QLabel()
                        label1.setPixmap(page1_pixmap)
                        label1.setAlignment(Qt.AlignmentFlag.AlignCenter)
                        page_layout.addWidget(label1)
                    else:
                        label1 = QLabel('No page in PDF 1')
                        label1.setAlignment(Qt.AlignmentFlag.AlignCenter)
                        page_layout.addWidget(label1)

                    if img2 is not None:
                        img2 = self.process_image_for_comparison(img2)
                        page2_pixmap = QPixmap.fromImage(QImage(img2.data, img2.shape[1], img2.shape[0], QImage.Format.Format_RGB888))
                        label2 = QLabel()
                        label2.setPixmap(page2_pixmap)
                        label2.setAlignment(Qt.AlignmentFlag.AlignCenter)
                        page_layout.addWidget(label2)
                    else:
                        label2 = QLabel('No page in PDF 2')
                        label2.setAlignment(Qt.AlignmentFlag.AlignCenter)
                        page_layout.addWidget(label2)

                    # Add difference map if enabled
                    if self.show_difference_map and img1 is not None and img2 is not None:
                        diff_map = self.get_difference_map(img1, img2)
                        diff_pixmap = QPixmap.fromImage(QImage(diff_map.data, diff_map.shape[1], diff_map.shape[0], QImage.Format.Format_RGB888))
                        diff_label = QLabel()
                        diff_label.setPixmap(diff_pixmap)
                        diff_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
                        page_layout.addWidget(diff_label)

                    self.comparison_layout.addWidget(page_container)
                    QApplication.processEvents()  # Keep UI responsive

                self.statusBar().showMessage('Comparison complete')
            
            finally:
                doc1.close()
                doc2.close()

        except Exception as e:
            QMessageBox.critical(self, 'Error', f'Error comparing PDFs: {str(e)}')
            self.statusBar().showMessage('Error during comparison')
            print(f"Error details: {str(e)}")  # For debugging

    def get_difference_map(self, img1_array, img2_array):
        """Generate a visual difference map between two images"""
        if img1_array.shape != img2_array.shape:
            # Resize to match
            height = max(img1_array.shape[0], img2_array.shape[0])
            width = max(img1_array.shape[1], img2_array.shape[1])
            
            if img1_array.shape[0] != height or img1_array.shape[1] != width:
                img1 = Image.fromarray(img1_array)
                img1_array = np.array(img1.resize((width, height)))
            
            if img2_array.shape[0] != height or img2_array.shape[1] != width:
                img2 = Image.fromarray(img2_array)
                img2_array = np.array(img2.resize((width, height)))

        # Calculate absolute difference
        diff = np.abs(img1_array.astype(np.float32) - img2_array.astype(np.float32))
        
        # Convert to grayscale for visualization
        diff_gray = np.max(diff, axis=2)
        
        # Normalize and enhance contrast
        diff_normalized = ((diff_gray > self.difference_threshold) * 255).astype(np.uint8)
        
        # Apply color map for better visualization
        diff_color = cv2.applyColorMap(diff_normalized, cv2.COLORMAP_JET)
        
        return diff_color

    def process_image_for_comparison(self, img_array):
        """Process image based on selected comparison mode"""
        if self.comparison_mode == 'grayscale':
            gray = cv2.cvtColor(img_array, cv2.COLOR_RGB2GRAY)
            return cv2.cvtColor(gray, cv2.COLOR_GRAY2RGB)
        elif self.comparison_mode == 'text only':
            # Apply threshold to focus on text
            gray = cv2.cvtColor(img_array, cv2.COLOR_RGB2GRAY)
            _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            return cv2.cvtColor(binary, cv2.COLOR_GRAY2RGB)
        return img_array

def main():
    app = QApplication(sys.argv)
    window = PDFComparisonTool()
    window.show()
    sys.exit(app.exec())

if __name__ == '__main__':
    main()