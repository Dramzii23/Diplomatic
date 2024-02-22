# DIPLOMATIC

## List

# Modules and Imports

- **fitz**: PyMuPDF module used for PDF processing
- **PyQt5.QtGui**: Contains classes for various GUI operations
  - Imports:
    - `QPixmap`: Used for handling images
    - `QFont`: Used for specifying the font's properties
    - `QPainter`: Used for drawing 2D graphics
    - `QPdfWriter`: Used for creating PDF documents
    - `QColor`: Used for colors in terms of RGB, HSV, and CMYK values
    - `QPen`: Used for drawing lines, curves, and outlines of shapes
    - `QImage`: Used for hardware-independent image representation
    - `QFontMetrics`: Used for calculating font and text metrics
- **PyQt5.QtWidgets**: Contains classes for creating the application, dialog windows, splash screens, file dialogs, buttons, font combo boxes, dials, and labels
  - Imports:
    - `QApplication`: Manages the GUI application's control flow and main settings
    - `QDialog`: Used for creating dialog windows
    - `QSplashScreen`: Used for displaying a splash screen during application startup
    - `QFileDialog`: Used for displaying file dialogs
    - `QPushButton`: Used for creating push buttons
    - `QFontComboBox`: Used for displaying a combo box with font names
    - `QDial`: Used for creating a dial (like a knob)
    - `QLabel`: Used for creating a text or image display
- **PyQt5.uic**: Contains the function used to load a user interface from a .ui file
  - Imports:
    - `loadUi`: Used for loading a user interface from a .ui file
- **PyQt5.QtCore**: Contains classes and constants for various core operations such as timers, URLs, and size calculations
  - Imports:
    - `Qt`: Contains global constants and helper functions
    - `QTimer`: Used for creating and managing timers
    - `QUrl`: Used for parsing and constructing URLs
    - `QSizeF`: Used for specifying the size of graphical objects
- **sys**: Used for accessing command-line arguments and exiting the application
- **os**: Used for file path manipulations

# Classes and Methods

- **MainWindow.\_\_init\_\_**: This is a special method in Python classes known as a constructor. This method is called when an object is created from a class and it allows the class to initialize the attributes of the class.
- **MainWindow.upload_image**: This is a method of the `MainWindow` class. It's used to upload an image from a file.
- **MainWindow.change_font**: This is a method of the `MainWindow` class. It's used to change the font of a label.
- **MainWindow.change_font_size**: This is a method of the `MainWindow` class. It's used to change the font size of a label.
- **MainWindow.export_pdf**: This is a method of the `MainWindow` class. It's used to export the current image and text to a PDF file.
- **MainWindow.preview_pdf**: This is a method of the `MainWindow` class. It's used to preview the PDF by saving it to a temporary file and displaying the first page in a label.

# Sections

- **if \_\_name\_\_ == "\_\_main\_\_"**: This is a common Python idiom. In a Python file, `__name__` is a special variable that's the name of the module. If the module is being run directly (as opposed to being imported), `__name__` will be `"__main__"`. So, this condition is `True` if the file is being run directly. This is typically where the application is actually started, after all the classes and functions have been defined.
