import sys
import os

from PyQt5.QtWidgets import QMainWindow, QPushButton, QApplication, QLineEdit, QLabel, QVBoxLayout, QScrollArea, QWidget, QMessageBox, QFileDialog
from PyQt5.QtGui import *
from PyQt5.QtCore import *

from time import gmtime, strftime
import magic


class PatientFinder(QMainWindow):

    def __init__(self):
        super().__init__()

        self.initUI()

    def initUI(self):

        btn1 = QPushButton("Find patients", self)
        btn1.setGeometry(10, 20, 200, 25)
        btn1.clicked.connect(self.buttonClicked)

        btn2 = QPushButton("Save to text file", self)
        btn2.setGeometry(200, 20, 200, 25)
        btn2.clicked.connect(self.saveToTxt)

        scroll = QScrollArea(self)
        scroll.move(15, 60)
        scroll.setFixedHeight(310)
        scroll.setFixedWidth(380)
        scroll.setStyleSheet('background-color: white')

        self.output = QLabel("Result will appear here", self)
        self.output.move(30, 90)
        self.output.setFixedSize(300, 300)
        self.output.setStyleSheet('background-color: white')
        self.output.setAlignment(Qt.AlignTop)

        scroll.setWidget(self.output)

        self.statusBar()

        self.setGeometry(300, 300, 410, 425)
        self.setFixedHeight(425)
        self.setFixedWidth(410)
        self.setWindowTitle('Patient Finder')
        self.show()

    def buttonClicked(self):
        self.outString = ""
        self.found = []

        directory = str(QFileDialog.getExistingDirectory(
            self, "Select directory to search from"))

        if directory:
            # Start from root dir
            for root, dirs, files in os.walk(directory):
                # Append each file name found to output string
                for file in files:
                    with open(os.path.join(root, file), 'rb') as input:
                        mime_type = magic.from_buffer(input.readline())
                        if "PDF" in mime_type:
                            print(file)
                            self.outString = self.outString + \
                                (str(file) + "\n")
                            self.found.append(file)

            # For debug purposes. Remove later
            ''' buttonReply = QMessageBox.question(
                self, 'Saved', self.outString, QMessageBox.Ok) '''

            # Display the output string on GUI
            self.output.setText(self.outString)
            self.output.setFixedHeight(20 * len(self.found))
            self.output.repaint()
            self.show()

    def saveToTxt(self):
        directory = str(QFileDialog.getExistingDirectory(
            self, "Select directory to save list in."))

        if directory:
            try:
                print(self.found)
                print("Saving to txt...")
                now = strftime("%d%m%d_%H%M", gmtime())
                fileName = directory + "/patients_list_" + now + ".txt"
                f = open(fileName, "w+")
                f.write(self.outString)
                f.close()
                QMessageBox.question(
                    self, 'Saved', "List saved to:\n\n" + fileName, QMessageBox.Ok)
                print("Done...")
            except Exception as e:
                QMessageBox.question(
                    self, 'Error', "The list is empty. Please press 'Find Patients' button before saving to text file", QMessageBox.Ok)
                print("List is empty. Abort command.")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = PatientFinder()
    sys.exit(app.exec_())
