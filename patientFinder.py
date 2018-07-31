# Written by Danish Bassi
#
# A trivial script which utilises the PyQt5 library to create a simple GUI
# for recursively fetching data stored in PDF files starting from a directory
# chosen by the user
#
# TO DO: Extract data from PDF files to analyze


import sys
import os
import re

from PyQt5.QtWidgets import QMainWindow, QPushButton, QApplication, QLineEdit, QLabel, QVBoxLayout, QScrollArea, QWidget, QMessageBox, QFileDialog
from PyQt5.QtGui import *
from PyQt5.QtCore import *

from datetime import datetime
# import magic
# from mimetypes import MimeTypes
import filetype
import winshell
import glob


class PatientFinder(QMainWindow):

    def __init__(self):
        super().__init__()

        self.initUI()

    def initUI(self):

        # Initialise global variables
        self.found = []

        # Create the layout
        self.textbox = QLineEdit(self)
        self.textbox.setGeometry(15, 10, 380, 25)
        self.textbox.setPlaceholderText("URN1 URN2 URN3 ...")

        btn1 = QPushButton("Find patients", self)
        btn1.setGeometry(10, 40, 200, 25)
        btn1.clicked.connect(self.findPatients)

        btn2 = QPushButton("Save to text file", self)
        btn2.setGeometry(200, 40, 200, 25)
        btn2.clicked.connect(self.saveToTxt)

        scroll = QScrollArea(self)
        scroll.move(15, 75)
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

        URNS = self.textbox.text()
        if re.match("^[0-9 ]+$", URNS) is not None:
            URNS = URNS.split()

            # Obtain the directory to begin searching from
            directory = str(QFileDialog.getExistingDirectory(
                self, "Select directory to search from"))

            # If a directory was selected, continue
            if directory:
                # Display message box to inform this process can take a while
                QMessageBox.question(
                    self, 'Information', "Patient searching will begin after you press OK.\n\nPlease be patient as this could possibly take a while", QMessageBox.Ok)
                # Start from the directory and recursively fetch each file
                for root, dirs, files in os.walk(directory):
                    # For each file that was found
                    for file in files:
                       
                        # Get the target path of the shortcut link
                        shortcut = winshell.shortcut(os.path.join(root, file))    
                      
                        # Check if file is of PDF type
                        kind = filetype.guess(shortcut.path)

                        # If it is not a PDF file
                        if kind is None:
                            # Continue to the next file
                            continue
                        # If it is a PDF
                        elif "pdf" in str(kind.mime):
                            
                            # Get the URN from the name and check if it is in the input list
                            fileURN = re.findall('\d+', file )                                

                            # Check if the URN is in the list of input URNs
                            for ur in fileURN:
                                if ur in URNS:
                                    # Added file name to the output string
                                    self.outString = self.outString + \
                                        (str(file[:-4]) + "\n")
                                    self.found.append(file[:-4])                        
                        
                # For debug purposes. Remove later
                ''' buttonReply = QMessageBox.question(
                    self, 'Saved', self.outString, QMessageBox.Ok) '''

                # Display the output string on GUI
                self.output.setText(self.outString)
                self.output.setFixedHeight(20 * len(self.found))
                self.output.repaint()
                self.show()
        elif URNS == "":
            QMessageBox.question(
                    self, 'Error', "Please input one or more URNs before proceeding.", QMessageBox.Ok)
        else:
            QMessageBox.question(
                    self, 'Error', "One or more URNs contain non-numeric numbers. Please fix before proceeding.", QMessageBox.Ok)

    # Save list button handler
    def saveToTxt(self):
        
        # Only allow to export as txt if items were found
        if len(self.found) != 0:
            # Obtain the directory to save list in
            directory = str(QFileDialog.getExistingDirectory(
                self, "Select directory to save list in."))
        else:
            QMessageBox.question(
                self, 'Error', "The list is empty. Please press 'Find Patients' button before saving to text file", QMessageBox.Ok)
            print("List is empty. Abort command.")
            return

        # If a directory was selected, save list in a txt file in chosen directory with current timestamp
        if directory:
            try:
                print(self.found)
                print("Saving to txt...")
                now = datetime.now().strftime("%d%m%y_%H%M")
                fileName = directory + "/patients_list_" + now + ".txt"
                f = open(fileName, "w+")
                f.write(self.outString)
                f.close()
                QMessageBox.question(
                    self, 'Saved', "List saved to:\n\n" + fileName, QMessageBox.Ok)
                print("Done...")
            except Exception as e:
                QMessageBox.question(
                    self, 'Error', e, QMessageBox.Ok)
                print("Error occured. Abort command.")


# Execute the program
if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = PatientFinder()
    sys.exit(app.exec_())
