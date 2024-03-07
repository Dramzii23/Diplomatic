from PyQt5.QtWidgets import QDialog
from PyQt5.uic import loadUi
from utils import upload_image, change_font, change_font_size, export_pdf, preview_pdf
from PyQt5.QtWidgets import QWidget

class MainWindow(QDialog):
    
    def __init__(self):
        super(MainWindow, self).__init__()
        loadUi("ui_files/mainwindow.ui", self)
        
        
        # self.fontComboBox.currentFontChanged.connect(lambda: change_font(self))
        self.dialFontSize.valueChanged.connect(lambda: change_font_size(self))
        self.uploadImage.clicked.connect(lambda: upload_image(self))
        self.exportPDF.clicked.connect(lambda: export_pdf(self))
        # self.previewPDF.clicked.connect(lambda: preview_pdf(self))
        self.print_widget_names()
        
    def print_widget_names(window):
        widgets = window.findChildren(QWidget)
        for widget in widgets:
            print(type(widget).__name__, widget.objectName())
            