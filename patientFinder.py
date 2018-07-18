# Written by Danish Bassi
#
# A trivial script which utilises the PyQt5 library to create a simple GUI
# for fetching data stored in PDF files stored on a drive

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

        # Create the layout
        btn1 = QPushButton("Find patients", self)
        btn1.setGeometry(10, 20, 200, 25)
        btn1.clicked.connect(self.findPatients)

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

    # Find patients button handler
    def findPatients(self):
        self.outString = ""
        self.found = []

        # Obtain the directory to begin searching from
        directory = str(QFileDialog.getExistingDirectory(
            self, "Select directory to search from"))

        # If a directory was selected, continue
        if directory:
            # Start from the directory and recursively fetch each file
            for root, dirs, files in os.walk(directory):
                # Append each file name found to output string
                for file in files:
                    # Check if file is of PDF type
                    with open(os.path.join(root, file), 'rb') as input:
                        mime_type = magic.from_buffer(input.readline())
                        # If it is a PDF file, append to string to display output
                        if "PDF" in mime_type:
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
        # Obtain the directory to save list in 
        directory = str(QFileDialog.getExistingDirectory(
            self, "Select directory to save list in."))

        # If a directory was selected, save list in a txt file in chosen directory with current timestamp
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

# Execute the program
if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = PatientFinder()
    sys.exit(app.exec_())
