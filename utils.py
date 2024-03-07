from PyQt5.QtWidgets import QFileDialog
from PyQt5.QtGui import QPixmap, QFont, QPainter, QPdfWriter, QColor, QFontMetrics, QImage
from PyQt5.QtCore import Qt, QSizeF
import fitz
import os

def upload_image(window):
    '''
    This function opens a file dialog for the user to select an image file.
    The selected image is then displayed in the imageLabel widget of the provided window.
    '''
    file_name, _ = QFileDialog.getOpenFileName(window, 'Open Image File', r"<Default dir>", "Image files (*.jpg *.gif *.png)")
    pixmap = QPixmap(file_name)
    window.imageLabel.setPixmap(pixmap)
    
def change_font(window, font):
    '''
    This function changes the font of the text in the textEdit widget of the
    provided window to the specified font.
    '''
    window.textEdit.setFont(QFont(font))

def change_font_size(window, value):
    '''
    This function changes the font size of the text in the textEdit widget of
    the provided window to the specified value
    '''
    font = window.textEdit.font()
    font.setPointSize(value)
    window.textEdit.setFont(font)

def export_pdf(window):
    '''
    This function exports the content of the imageLabel widget of the provided
    window to a PDF file. It creates a QPdfWriter object, sets the page size to
    the size of the imageLabel widget, and then draws the pixmap of the imageLabel
    widget onto the PDF.
    '''
    from PyQt5.QtGui import QPageSize, QPdfWriter
    from PyQt5.QtCore import QSizeF

    # Create a QPdfWriter object
    pdf_writer = QPdfWriter("output.pdf")

    # Get the size of the imageLabel widget
    size = window.imageLabel.size()

    # Create a QPageSize object with the size of the imageLabel widget
    page_size = QPageSize(QSizeF(size.width(), size.height()), QPageSize.Point)

    # Set the page size of the PDF writer
    pdf_writer.setPageSize(page_size)
    
    

    pdf_writer = QPdfWriter("certificate.pdf")
    pdf_writer.setPageSize(QSizeF(window.imageLabel.size()))
    painter = QPainter(pdf_writer)
    scale_x = pdf_writer.pageSize().width() / window.imageLabel.width()
    scale_y = pdf_writer.pageSize().height() / window.imageLabel.height()
    painter.scale(scale_x, scale_y)
    painter.drawPixmap(0, 0, QPixmap(window.imageLabel.pixmap()))
    painter.end()

def preview_pdf(window):
    '''
    This function opens a PDF file, extracts all images from it, and displays
    the first image in the labelImage widget of the provided window. It uses the
    fitz library to open the PDF and extract images.
    '''
    doc = fitz.open("certificate.pdf")
    for i in range(len(doc)):
        for img in doc.get_page_images(i):
            xref = img[0]
            base = os.path.splitext("certificate.pdf")[0]
            pix = fitz.Pixmap(doc, xref)
            if pix.n < 5:       # this is GRAY or RGB
                pix.writePNG("%s.png" % (base,))
            else:               # CMYK: convert to RGB first
                pix1 = fitz.Pixmap(fitz.csRGB, pix)
                pix1.writePNG("%s.png" % (base,))
                pix1 = None
            pix = None
    window.labelImage.setPixmap(QPixmap("%s.png" % (base,)))