import fitz  
# PyMuPDF
from PyQt5.QtGui import QPixmap, QFont, QPainter, QPdfWriter, QColor, QPen, QImage
from PyQt5.QtWidgets import QApplication, QDialog, QSplashScreen, QFileDialog, QPushButton, QMessageBox, QFontComboBox, QDial, QLabel
from PyQt5.uic import loadUi
from PyQt5.QtCore import Qt, QTimer, QUrl, QSizeF, QMarginsF
from PyQt5.QtGui import QFontMetrics, QDesktopServices, QStandardItemModel, QStandardItem
from PyQt5 import QtWidgets, uic
import pandas as pd



# import pyglet
import sys
import os

# https://stackoverflow.com/questions/31836104/pyinstaller-and-onefile-how-to-include-an-image-in-the-exe-file
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS2
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

class MainWindow(QDialog):
    def __init__(self):
        # Call the constructor of the parent class (QMainWindow)
        super(MainWindow, self).__init__()

        # Load the user interface from the "mainwindow.ui" file
        loadUi(resource_path("ui_files\\mainwindow.ui"), self)
        self.image_file = resource_path('defaultImage.png')
        self.imageLabel.setPixmap(QPixmap(self.image_file))

        # Connect the currentFontChanged signal of the fontComboBox to the change_font method
        # This means that when the current font of the fontComboBox changes, the change_font method will be called
        # self.fontComboBox.currentFontChanged.connect(self.change_font)

        # Connect the valueChanged signal of the dialFontSize to the change_font_size method
        # This means that when the value of the dialFontSize changes, the change_font_size method will be called
       ######### # self.dialFontSize.valueChanged.connect(self.change_font_size)

        # Connect the clicked signal of the uploadImage QPushButton to the upload_image method
        # This means that when the uploadImage button is clicked, the upload_image method will be called
        self.uploadImage.clicked.connect(self.upload_image)

        # Connect the clicked signal of the exportPDF QPushButton to the export_pdf method
        # This means that when the exportPDF button is clicked, the export_pdf method will be called
        
        self.exportPDF.clicked.connect(self.export_pdf)
        
        
        # Assuming subirExcelButton and tableViewExcel are members of self
        self.subirExcelButton.clicked.connect(self.handle_subirExcelButton_clicked)
        
    
        # Connect the clicked signal of the previewPDF QPushButton to the preview_pdf method
        # This means that when the previewPDF button is clicked, the preview_pdf method will be called
        # self.previewPDF.clicked.connect(self.preview_pdf)
        # self.preview_pdf.clicked.connect(self.preview_pdf)
        # self.previewButton.clicked.connect(self.preview_pdf)

        # # Create a new QPushButton named 'Preview PDF'
        # self.previewPDF = QPushButton('Preview PDF ORC', self)
        # # Connect the clicked signal of the new previewPDF QPushButton to the preview_pdf method
        # # This means that when the new previewPDF button is clicked, the preview_pdf method will be called
        # self.previewPDF.clicked.connect(self.preview_pdf)
        
        
    # END OF __init__ METHOD
    def handle_subirExcelButton_clicked(self):
            # QMessageBox.information(self, "Title", "Hello 1") 
            file, _ = QFileDialog.getOpenFileName(self, 'Open Excel File', '', 'Excel Files (*.xlsx)')
            # QMessageBox.information(self, "Title", "Hello 2") 
            if file:
                df = pd.read_excel(file)
                # QMessageBox.information(self, "Title", "Hello 3") 
                model = QStandardItemModel(df.shape[0], df.shape[1])
                
                for row in df.iterrows():
                    for col, value in enumerate(row[1]):
                        
                        item = QStandardItem(str(value))
                        model.setItem(row[0], col, item)
                self.tableViewExcel.setModel(model)
                # Get the first column of the DataFrame as a list
                first_column = df.iloc[:, 0].tolist()

                # Join the list into a single string with newline characters between each item
                first_column_str = '\n'.join(str(item) for item in first_column)
                
                QMessageBox.information(self, "Title", first_column_str) 
                # Set the text of labelName to the string
                # self.labelName.setText(first_column_str)
                # Call the export_pdf function and set the text of labelName for each line in first_column_str
                # for i, line in enumerate(first_column_str.split('\n')):
                #     self.export_pdf(line, f"output{i}.pdf")
                #     self.labelName.setText(line)
                for i, line in enumerate(first_column_str.split('\n')):
                    
                    self.labelName.setText(line)
                    self.export_pdf(str(i) + " " + line)
                    
    def upload_image(self):
        # Open a QFileDialog to select an image file
        # The getOpenFileName method returns a tuple where the first element is the selected file path
        # The second element is the selected filter, which we don't need, hence the underscore
        
        
        
        # self.image_file, _ = QFileDialog.getOpenFileName(self, "Open Image", self.image_file, "Image Files (*.png *.jpg *.bmp)")
        bgImage, _ = QFileDialog.getOpenFileName(self, "Open Image", self.image_file, "Image Files (*.png *.jpg *.bmp)")
        if bgImage:
            self.image_file = bgImage


        # If a file was selected (i.e., if self.image_file is not an empty string)
        if self.image_file:
            # Create a QPixmap from the image file
            pixmap = QPixmap(self.image_file)
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

    def export_pdf(self, line_content):
        
        
        export_path = resource_path("PDF\\For print\\")
        
        # Check if the directory exists
        folder_path = os.path.dirname(export_path)
        
        os.makedirs(folder_path, exist_ok=True)
        if not os.path.exists(os.path.dirname(export_path)):
            # Create the directory
            os.makedirs(os.path.dirname(export_path))
        
        # Create the filename from the line_content
        filename = f"{line_content}.pdf"
        
        # Combine the default directory and the filename to create the full file path
        pdf_file = os.path.join(folder_path, filename)

        # Check if the file path is valid
        if not os.path.isdir(os.path.dirname(pdf_file)):
            print(f"Invalid file path: {pdf_file}")
            return
            
        # Open a QFileDialog to select the PDF file
        # pdf_file, _ = QFileDialog.getSaveFileName(self, "Export PDF",folder_path, "PDF Files (*.pdf)")
           
            
        if pdf_file:
            # Check if the file path is valid
            if not os.path.isdir(os.path.dirname(pdf_file)):
                print(f"Invalid file path: {pdf_file}")
                return

        # Create a QPdfWriter
        pdf_writer = QPdfWriter(pdf_file)
        # Set the page size to 11 x 8.5 inches (1 inch = 25.4 mm)
        pdf_writer.setPageSizeMM(QSizeF(11 * 25.4, 8.5 * 25.4))
        # Remove the margins
        pdf_writer.setPageMargins(QMarginsF(0, 0, 0, 0))
        # Set the resolution to 300 DPI
        pdf_writer.setResolution(300)

        # Create a QPainter object with the QPdfWriter object as the paint device
        painter = QPainter(pdf_writer)

        # Check if QPainter is active
        if not painter.isActive():
            
            # self.image_file = resource_path("defaultImage.png")
            self.image_file = bgImage
            
            pixmap = QPixmap(self.image_file)
            
            
            if pixmap:
                painter.drawPixmap(0, 0, pixmap)
            else:
                print("No pixmap in imageLabel.")
                return

            print("Failed to initialize QPainter.")
            QMessageBox.information(self, "Information", "No puedo sobreescribir un PDF abierto")
            return

        pixmap = QPixmap(self.image_file)
        if pixmap:
            painter.drawPixmap(0, 0, pixmap)
        else:
            print("No pixmap in imageLabel.")
            return

        pixmapWidth, pixmapHeight = pixmap.width(), pixmap.height()
        painter.setViewport(0, 0, pixmapWidth, pixmapHeight)
        painter.setWindow(0, 0, pixmapWidth, pixmapHeight)
        painter.drawPixmap(0, 0, pixmap)

        font = self.labelName.font()
        font.setPointSize(int(self.labelName.font().pointSize() * 1.5))
        self.labelName.setFont(font)
        font.setPointSize(int(self.labelName.font().pointSize()))
        painter.setFont(font)
        painter.setPen(QColor(Qt.white))

        font_metrics = QFontMetrics(self.labelName.font())
        text_width = int(font_metrics.horizontalAdvance(self.labelName.text()) * 1.55)
        text_height = int(font_metrics.height())

        page_center_x = pdf_writer.width() // 2
        page_center_y = int(pdf_writer.height() * .65)

        text_x = page_center_x - text_width // 2
        text_y = page_center_y - text_height // 2

        painter.drawText(page_center_x - text_width, page_center_y, self.labelName.text())

        painter.end()
        self.preview_pdf(pdf_file)

        font.setPointSize(int(self.labelName.font().pointSize() / 1.5))
        self.labelName.setFont(font)
        
    def preview_pdf(self, pdf_file):
        # Open the PDF file in the default PDF viewer
        QDesktopServices.openUrl(QUrl.fromLocalFile(pdf_file))
        # The openUrl method of the QDesktopServices class is used to open the PDF file in the default PDF viewer.
        # The QUrl.fromLocalFile method is used to create a QUrl object from the file path.
        # The openUrl method then opens the PDF file in the default PDF viewer.
        
        
        
if __name__ == "__main__":
    app = QApplication(sys.argv)

    # Load the splash image
    splash_image = QPixmap(resource_path("img\\logo dramzii - Text  only 2.png"))

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