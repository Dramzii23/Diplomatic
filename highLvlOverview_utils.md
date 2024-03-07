- **upload_image(window)**: This function opens a file dialog for the user to select an image file. The selected image is then displayed in the `imageLabel` widget of the provided window.

- **change_font(window, font)**: This function changes the font of the text in the `textEdit` widget of the provided window to the specified font.

- **change_font_size(window, value)**: This function changes the font size of the text in the `textEdit` widget of the provided window to the specified value.

- **export_pdf(window)**: This function exports the content of the `imageLabel` widget of the provided window to a PDF file. It creates a `QPdfWriter` object, sets the page size to the size of the `imageLabel` widget, and then draws the pixmap of the `imageLabel` widget onto the PDF.

- **preview_pdf(window)**: This function opens a PDF file, extracts all images from it, and displays the first image in the `labelImage` widget of the provided window. It uses the `fitz` library to open the PDF and extract images.
