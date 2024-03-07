import unittest
from PyQt5.QtWidgets import QApplication, QFileDialog
from UI import YourClassName

class TestUploadImage(unittest.TestCase):

    def setUp(self):
        self.app = QApplication([])
        self.window = YourClassName()  # Replace YourClassName with the actual class name of your UI

    def tearDown(self):
        self.window.close()

    def test_upload_image(self):
        # Simulate selecting an image file using QFileDialog
        file_path = "/path/to/image.png"
        QFileDialog.getOpenFileName = lambda *args: (file_path, "Image Files (*.png *.jpg *.bmp)")

        # Call the upload_image method
        self.window.upload_image()

        # Assert that the image file path is set correctly
        self.assertEqual(self.window.image_file, file_path)

        # Assert that the imageLabel pixmap is set correctly
        expected_pixmap = QPixmap(file_path).scaled(self.window.imageLabel.size(), Qt.KeepAspectRatio)
        self.assertEqual(self.window.imageLabel.pixmap(), expected_pixmap)

if __name__ == '__main__':
    unittest.main()