import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QLabel, QPushButton, QFileDialog, QTextEdit, QHBoxLayout, \
    QScrollArea, QGroupBox, QMessageBox
from PyQt5.QtGui import QImage, QPixmap, QPainter, QPen
from PyQt5.QtCore import Qt, QRect
import pytesseract
from pdfkit import from_string
from PyQt5.QtWidgets import QDialog
from PIL import Image

from qt_material import apply_stylesheet


class PDFConverterApp(QWidget):
    def __init__(self):
        super().__init__()

        # Initialize the UI
        self.setWindowTitle("Computer Vision Project")
        self.image_label = QLabel(self)
        self.load_button = QPushButton("Load Image(s)", self)
        self.convert_button = QPushButton("Convert to PDF", self)
        self.text_edit = QTextEdit(self)
        self.save_word_button = QPushButton("Save as Word", self)
        self.save_pdf_button = QPushButton("Save as PDF", self)
        self.history_button = QPushButton("View History", self)
        self.crop_button = QPushButton("Crop Image", self)
        self.footer_label = QLabel("Under the guidance of Prof. Sunil Survaia", self)

        # Set up the layout
        layout = QVBoxLayout()
        layout.addWidget(self.image_label)
        layout.addWidget(self.load_button)
        layout.addWidget(self.convert_button)
        layout.addWidget(self.text_edit)
        layout.addWidget(self.save_word_button)
        layout.addWidget(self.save_pdf_button)
        layout.addWidget(self.history_button)
        layout.addWidget(self.crop_button)

        dashboard_layout = QHBoxLayout()
        dashboard_layout.addLayout(layout)
        dashboard_layout.addWidget(self.create_history_group_box())
        self.setLayout(dashboard_layout)

        # Connect the buttons to their respective functions
        self.load_button.clicked.connect(self.load_images)
        self.convert_button.clicked.connect(self.convert_to_pdf)
        self.save_word_button.clicked.connect(self.save_as_word)
        self.save_pdf_button.clicked.connect(self.save_as_pdf)
        self.history_button.clicked.connect(self.view_history)
        self.crop_button.clicked.connect(self.crop_image)

        # Initialize image paths, extracted text, and history variables
        self.image_paths = []
        self.extracted_text = ""
        self.history = []

    def create_history_group_box(self):
        group_box = QGroupBox("History")
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area_content = QWidget()
        scroll_area_layout = QVBoxLayout(scroll_area_content)
        scroll_area_layout.setAlignment(Qt.AlignTop)
        scroll_area.setWidget(scroll_area_content)
        group_box_layout = QVBoxLayout()
        group_box_layout.addWidget(scroll_area)
        group_box.setLayout(group_box_layout)
        return group_box

    def add_to_history(self, image_path, extracted_text):
        self.history.append((image_path, extracted_text))

    def update_history_ui(self):
        scroll_area_content = self.history_button.parent().findChild(QWidget, "scrollAreaContent")
        scroll_area_layout = scroll_area_content.layout()
        if scroll_area_layout is not None:
            # Clear previous history entries
            while scroll_area_layout.count():
                item = scroll_area_layout.takeAt(0)
                widget = item.widget()
                if widget:
                    widget.deleteLater()

            # Add new history entries
            for image_path, extracted_text in self.history:
                history_entry = QLabel(f"<b>Image:</b> {image_path}<br><b>Extracted Text:</b> {extracted_text}")
                history_entry.setWordWrap(True)
                scroll_area_layout.addWidget(history_entry)

    def load_images(self):
        # Open a file dialog to select multiple image files
        file_dialog = QFileDialog()
        image_paths, _ = file_dialog.getOpenFileNames(
            self, "Select Image(s)", "", "Image Files (*.png *.jpg *.jpeg)"
        )

        if image_paths:
            # Clear previous data
            self.image_paths.clear()
            self.extracted_text = ""

            # Load the selected images
            for image_path in image_paths:
                self.image_paths.append(image_path)
                image = QImage(image_path)
                pixmap = QPixmap.fromImage(image)
                self.image_label.setPixmap(pixmap.scaled(640, 480, Qt.KeepAspectRatio))

    def convert_to_pdf(self):
        # Check if any images are loaded
        if self.image_paths:
            # Initialize the extracted text
            self.extracted_text = ""

            # Iterate through the loaded images and perform OCR
            for image_path in self.image_paths:
                pil_image = Image.open(image_path)
                extracted_text = pytesseract.image_to_string(pil_image)
                self.extracted_text += extracted_text + "\n"

            # Update the text edit widget
            self.text_edit.setPlainText(self.extracted_text)

            # Add to history
            self.add_to_history(image_path, self.extracted_text)

            # Update the history UI
            self.update_history_ui()

            print("OCR completed!")
        else:
            print("No images loaded!")

    def save_as_word(self):
        # Check if any text is extracted
        if self.extracted_text:
            # Open a file dialog to select the save location and filename for Word document
            file_dialog = QFileDialog()
            save_path, _ = file_dialog.getSaveFileName(
                self, "Save as Word", "", "Word Files (*.docx)"
            )

            if save_path:
                # Save the extracted text as a Word document
                with open(save_path, 'w') as file:
                    file.write(self.extracted_text)

                print("Text saved as Word document!")
        else:
            print("No extracted text to save!")

    def save_as_pdf(self):
        # Check if any text is extracted
        if self.extracted_text:
            # Open a file dialog to select the save location and filename for PDF document
            file_dialog = QFileDialog()
            save_path, _ = file_dialog.getSaveFileName(
                self, "Save as PDF", "", "PDF Files (*.pdf)"
            )

            if save_path:
                # Convert the extracted text to a PDF
                pdf_content = "<pre>" + self.extracted_text + "</pre>"
                from_string(pdf_content, save_path)

                print("Text saved as PDF document!")
        else:
            print("No extracted text to save!")

    def view_history(self):
        scroll_area_content = self.history_button.parent().findChild(QWidget, "scrollAreaContent")
        scroll_area_content.show()

    def crop_image(self):
        if self.image_paths:
            # Load the image
            image_path = self.image_paths[0]
            image = QImage(image_path)

            # Create a crop dialog
            crop_dialog = CropDialog(image)
            if crop_dialog.exec_() == CropDialog.Accepted:
                # Get the crop rectangle from the dialog
                crop_rect = crop_dialog.get_crop_rect()

                # Perform cropping
                cropped_image = image.copy(crop_rect)
                pixmap = QPixmap.fromImage(cropped_image)
                self.image_label.setPixmap(pixmap.scaled(640, 480, Qt.KeepAspectRatio))

                print("Image cropped!")
        else:
            print("No image loaded!")


class CropDialog(QDialog):
    def __init__(self, image):
        super().__init__()

        self.setWindowTitle("Crop Image")
        self.image_label = QLabel(self)
        self.crop_button = QPushButton("Crop", self)
        self.cancel_button = QPushButton("Cancel", self)

        layout = QVBoxLayout()
        layout.addWidget(self.image_label)
        layout.addWidget(self.crop_button)
        layout.addWidget(self.cancel_button)
        self.setLayout(layout)

        self.crop_button.clicked.connect(self.accept)
        self.cancel_button.clicked.connect(self.reject)

        self.image = image
        self.crop_rect = QRect()

    def showEvent(self, event):
        super().showEvent(event)

        pixmap = QPixmap.fromImage(self.image)
        self.image_label.setPixmap(pixmap.scaled(640, 480, Qt.KeepAspectRatio))

    def mousePressEvent(self, event):
        self.crop_rect.setTopLeft(event.pos())

    def mouseMoveEvent(self, event):
        self.crop_rect.setBottomRight(event.pos())
        self.update()

    def mouseReleaseEvent(self, event):
        self.crop_rect.setBottomRight(event.pos())
        self.accept()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setPen(QPen(Qt.red, 2, Qt.SolidLine))
        painter.drawRect(self.crop_rect)

    def get_crop_rect(self):
        return self.crop_rect.normalized()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    apply_stylesheet(app, theme='dark_teal.xml')  # Apply Material Design dark theme
    converter = PDFConverterApp()
    converter.show()
    sys.exit(app.exec_())
