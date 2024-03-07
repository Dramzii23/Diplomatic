from PyQt5.QtWidgets import QApplication, QSplashScreen
from PyQt5.QtGui import QPixmap
from PyQt5.QtCore import QTimer, Qt
from main_window import MainWindow
import sys

if __name__ == "__main__":
    app = QApplication(sys.argv)
    splash_image = QPixmap("img/logo dramzii - Text  only 2.png")
    splash_image = splash_image.scaledToWidth(400, Qt.SmoothTransformation)
    splash = QSplashScreen(splash_image)
    splash.show()
    window = MainWindow()
    QTimer.singleShot(3000, splash.close)
    QTimer.singleShot(3000, window.show)
    print("All is good")
    sys.exit(app.exec_())