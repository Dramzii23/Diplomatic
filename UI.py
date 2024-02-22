import fitz  
# PyMuPDF
from PyQt5.QtGui import QPixmap, QFont, QPainter, QPdfWriter, QColor, QPen, QImage
from PyQt5.QtWidgets import QApplication, QDialog, QSplashScreen, QFileDialog, QPushButton, QFontComboBox, QDial, QLabel
from PyQt5.uic import loadUi
from PyQt5.QtCore import Qt, QTimer, QUrl, QSizeF
from PyQt5.QtGui import QFontMetrics
import sys
import os

class MainWindow(QDialog):
    def __init__(self):
        # Call the constructor of the parent class (QMainWindow)
        super(MainWindow, self).__init__()

        # Load the user interface from the "mainwindow.ui" file
        loadUi("ui_files/mainwindow.ui", self)

        # Connect the currentFontChanged signal of the fontComboBox to the change_font method
        # This means that when the current font of the fontComboBox changes, the change_font method will be called
        self.fontComboBox.currentFontChanged.connect(self.change_font)

        # Connect the valueChanged signal of the dialFontSize to the change_font_size method
        # This means that when the value of the dialFontSize changes, the change_font_size method will be called
        self.dialFontSize.valueChanged.connect(self.change_font_size)

        # Connect the clicked signal of the uploadImage QPushButton to the upload_image method
        # This means that when the uploadImage button is clicked, the upload_image method will be called
        self.uploadImage.clicked.connect(self.upload_image)

        # Connect the clicked signal of the exportPDF QPushButton to the export_pdf method
        # This means that when the exportPDF button is clicked, the export_pdf method will be called
        self.exportPDF.clicked.connect(self.export_pdf)

        # Connect the clicked signal of the previewPDF QPushButton to the preview_pdf method
        # This means that when the previewPDF button is clicked, the preview_pdf method will be called
        # self.previewPDF.clicked.connect(self.preview_pdf)
        # self.preview_pdf.clicked.connect(self.preview_pdf)
        # self.previewButton.clicked.connect(self.preview_pdf)

        # Create a new QPushButton named 'Preview PDF'
        self.previewPDF = QPushButton('Preview PDF', self)
        # Connect the clicked signal of the new previewPDF QPushButton to the preview_pdf method
        # This means that when the new previewPDF button is clicked, the preview_pdf method will be called
        self.previewPDF.clicked.connect(self.preview_pdf)
    # END OF __init__ METHOD
        
    def upload_image(self):
        # Open a QFileDialog to select an image file
        # The getOpenFileName method returns a tuple where the first element is the selected file path
        # The second element is the selected filter, which we don't need, hence the underscore
        image_file, _ = QFileDialog.getOpenFileName(self, "Open Image", "", "Image Files (*.png *.jpg *.bmp)")

        # If a file was selected (i.e., if image_file is not an empty string)
        if image_file:
            # Create a QPixmap from the image file
            pixmap = QPixmap(image_file)
            # Scale the QPixmap to fit the size of the imageLabel, while keeping the aspect ratio
            pixmap = pixmap.scaled(self.imageLabel.size(), Qt.KeepAspectRatio)
            # Set the QPixmap as the pixmap of the imageLabel
            self.imageLabel.setPixmap(pixmap)
    # END OF upload_image METHOD
        
    def change_font(self, font):
        """
        Summary: Changes the font of the labelName QLabel to the selected font.
        
        Purpose: This function is used to change the font of the text in the labelName QLabel. 
        It can be used to customize the appearance of the text in the UI.
        
        Use: It is called with a QFont object as an argument, which represents the new font to be used for the labelName QLabel.
        """
        self.labelName.setFont(font)

    def change_font_size(self, value):
        # Get the current font of the labelName QLabel
        font = self.labelName.font()

        # Set the font size to the value of the dialFontSize
        font.setPointSize(value)

        # Set the font of the labelName QLabel
        self.labelName.setFont(font)

    def export_pdf(self):
        # Open a QFileDialog to select the PDF file
        pdf_file, _ = QFileDialog.getSaveFileName(self, "Export PDF", "", "PDF Files (*.pdf)")

        if pdf_file:
            # Create a QPdfWriter
            pdf_writer = QPdfWriter(pdf_file)

            # Set the page size to 11 x 8.5 inches (1 inch = 25.4 mm)
            pdf_writer.setPageSizeMM(QSizeF(11 * 25.4, 8.5 * 25.4))

            # Set the resolution to 300 DPI
            pdf_writer.setResolution(300)

            # Create a QPainter
            painter = QPainter(pdf_writer)

            # Draw the image
            painter.drawPixmap(0, 0, self.imageLabel.pixmap())

            # Set the color of the text to red
            painter.setPen(QColor(Qt.red))

            # Draw the text at a specific position
            painter.setFont(self.labelName.font())
            pixmap_width = self.imageLabel.pixmap().width()
            pixmap_height = self.imageLabel.pixmap().height()

            # Calculate the center of the page
            page_center_x = pdf_writer.width() // 2
            page_center_y = pdf_writer.height() // 2

            # Calculate the width and height of the text
            font_metrics = QFontMetrics(self.labelName.font())
            text_width = font_metrics.horizontalAdvance(self.labelName.text())
            text_height = font_metrics.height()

            # Calculate the position of the text so that its center aligns with the center of the page
            text_x = page_center_x - text_width // 2
            text_y = page_center_y - text_height // 2 
              
            painter.drawText(text_x, text_y, self.labelName.text())

            # End the QPainter
            painter.end()
    def preview_pdf(self):
        # Save the PDF to a temporary file
        temp_pdf_file = os.path.join(os.path.dirname(__file__), "temp.pdf")

        # Create a QPdfWriter
        pdf_writer = QPdfWriter(temp_pdf_file)

        # Set the page size to 11 x 8.5 inches (1 inch = 25.4 mm)
        pdf_writer.setPageSizeMM(QSizeF(11 * 25.4, 8.5 * 25.4))

        # Set the resolution to 300 DPI
        pdf_writer.setResolution(300)

        # Create a QPainter
        painter = QPainter(pdf_writer)

        # Draw the image
        painter.drawPixmap(0, 0, self.imageLabel.pixmap())

        # Set the color of the text to red
        painter.setPen(QColor(Qt.red))

        # Draw the text at a specific position
        painter.setFont(self.labelName.font())
        pixmap_width = self.imageLabel.pixmap().width()
        pixmap_height = self.imageLabel.pixmap().height()

        # Calculate the center of the page
        page_center_x = pdf_writer.width() // 2
        page_center_y = pdf_writer.height() // 2

        # Calculate the width and height of the text
        font_metrics = QFontMetrics(self.labelName.font())
        text_width = font_metrics.horizontalAdvance(self.labelName.text())
        text_height = font_metrics.height()

        # Calculate the position of the text so that its center aligns with the center of the page
        text_x = page_center_x - text_width // 2
        text_y = page_center_y - text_height // 2 

        painter.drawText(text_x, text_y, self.labelName.text())

        # End the QPainter
        painter.end()

        # Open the PDF with PyMuPDF
        doc = fitz.open(temp_pdf_file)

    # Render the first page to an image
        page = doc.load_page(0)  # 0 is the first page
        pix = page.get_pixmap()

        # Convert the PyMuPDF pixmap to a QImage
        img = QImage(pix.samples, pix.width, pix.height, QImage.Format_RGB888)

        self.pdfPreviewLabel = QLabel(self)
        # Display the QImage in a QLabel
        self.pdfPreviewLabel.setPixmap(QPixmap.fromImage(img))

if __name__ == "__main__":
    app = QApplication(sys.argv)

    # Load the splash image
    splash_image = QPixmap("img/logo dramzii - Text  only 2.png")

    # Scale the splash image
    splash_image = splash_image.scaledToWidth(400, Qt.SmoothTransformation)

    # Create a QSplashScreen with the splash image
    splash = QSplashScreen(splash_image)
    splash.show()

    # Create an instance of MainWindow
    window = MainWindow()

    # Use a QTimer to close the splash screen after a delay and start the main window
    QTimer.singleShot(3000, splash.close)
    QTimer.singleShot(3000, window.show)  # Pass the method itself, not the result of calling it

    sys.exit(app.exec_())