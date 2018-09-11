# Written by Danish Bassi
#
# A trivial script which utilises the PyQt5 library to create a simple GUI
# for recursively fetching data stored in PDF files and Word documents starting
# from a directory chosen by the user

import sys
import os
import re
import subprocess
import uuid
from hashlib import blake2b

from PyQt5.QtWidgets import QMainWindow, QPushButton, QApplication, QLineEdit, QLabel, QVBoxLayout, QScrollArea, QWidget, QMessageBox, QFileDialog
from PyQt5.QtGui import *
from PyQt5.QtCore import *

from datetime import datetime
import filetype
import winshell
from win32com.client import Dispatch
import glob


class PatientFinder(QMainWindow):

    def __init__(self):
        super().__init__()
        
        # Insert MAC addresses of machines to run program on here
        licenses = {}
        
        # Encryption method removed for privacy reasons
        # (Insert encryption here)

        h = blake2b(digest_size=8)
        h.update(myMacAddress.encode())

        if h.hexdigest().upper() in licenses:   
            self.initUI()
        else:
            QMessageBox.question(
                    self, 'Not a valid device', "This machine is not licensed to use this software. Please contact Danish Bassi at bassidanish@gmail.com to receive a license. ", QMessageBox.Ok)
            sys.exit()

    def initUI(self):

        # Initialise global variables
        self.found = []

        # Create the layout
        self.textbox = QLineEdit(self)
        self.textbox.setGeometry(15, 10, 380, 25)
        self.textbox.setPlaceholderText("URN1 URN2 URN3 ...")
        self.textbox.setStyleSheet('border-radius: 5; padding-left: 5; background-color: white; border-style: solid; border-width: 1.2; border-color: rgb(7, 145, 134)')

        btn1 = QPushButton("Find with text input", self)
        btn1.setGeometry(15, 40, 185, 25)
        btn1.clicked.connect(self.useText)
        btn1.setStyleSheet('color: rgb(179, 43, 36); border-style: solid; background-color: white; border-radius: 5; border-width: 1.2; border-color: rgb(7, 145, 134)')

        btn2 = QPushButton("Save as shortcuts", self)
        btn2.setGeometry(110, 390, 185, 25)
        btn2.clicked.connect(self.saveToTxt)
        btn2.setStyleSheet('color: rgb(179, 43, 36); border-style: solid; background-color: white; border-radius: 5; border-width: 1.2; border-color: rgb(7, 145, 134)')

        scroll = QScrollArea(self)
        scroll.move(15, 75)
        scroll.setFixedHeight(310)
        scroll.setFixedWidth(380)
        scroll.setStyleSheet('background-color: white')
        scroll.setStyleSheet('background-color: white; padding: 5; border-style: solid; border-radius: 5;')

        self.output = QLabel("Result will appear here", self)
        self.output.setWordWrap(True)
        self.output.move(30, 90)
        self.output.setFixedSize(350, 300)
        self.output.setAlignment(Qt.AlignTop)

        scroll.setWidget(self.output)

        btn3 = QPushButton("Find with CSV", self)
        btn3.setGeometry(210, 40, 185, 25)
        btn3.clicked.connect(self.useCSV)
        btn3.setStyleSheet('color: rgb(179, 43, 36); border-style: solid; background-color: white; border-radius: 5; border-width: 1.2; border-color: rgb(7, 145, 134)')
        
        self.setGeometry(300, 200, 410, 425)
        self.setFixedHeight(425)
        self.setFixedWidth(410)
        self.setWindowTitle('Patient Finder - Royal Melbourne Hospital')
        self.setWindowIcon(QIcon('logo.jpg'))
        self.show()

    # Get URNs from a CSV file
    def useCSV(self):
        QMessageBox.question(
            self, 'Select CSV file', "Please select your CSV file in the following file dialog.", QMessageBox.Ok)
        
        csv = QFileDialog.getOpenFileName(self, "Select csv file")
        
        print(csv)
        if csv[0]:
            f = open(csv[0], 'r')
            try:
                with f :
                    content = f.readlines()
            except e as Exception:
                QMessageBox.question(
                    self, 'Error', "This is not a CSV file. Please try again.", QMessageBox.Ok)
            self.useCSV
            f.close()
            
            content = [x.strip(',.;\n ') for x in content]
            print(content)
            self.findPatients(content)

    # Use the text input to get URNs
    def useText(self):
        URNS = self.textbox.text()
        if re.match("^[0-9 ]+$", URNS) is not None:
            URNS = URNS.split()
            self.findPatients(URNS)
        elif URNS == "":
            QMessageBox.question(
                    self, 'Error', "Please input one or more URNs before proceeding.", QMessageBox.Ok)
        else:
            QMessageBox.question(
                    self, 'Error', "One or more URNs contain non-numeric numbers. Please fix before proceeding.", QMessageBox.Ok)

    # Find patients button handler
    def findPatients(self, URNS):
        self.outString = ""

        QMessageBox.question(
            self, 'Select directory', "Please select the directory to search from in the following file dialog.", QMessageBox.Ok)
      
        # Obtain the directory to begin searching from
        directory = str(QFileDialog.getExistingDirectory(
            self, "Select directory to search from", 'S:\\Neurology\\Neurology Files'))

        # If a directory was selected, continue
        if directory:
            self.found = []
            # Display message box to inform this process can take a while
            QMessageBox.question(
                self, 'Information', "Patient searching will begin after you press OK.\n\nPlease be patient as this could possibly take a while", QMessageBox.Ok)
                        
            # Start from the directory and recursively fetch each file
            for root, dirs, files in os.walk(directory):
                # For each file that was found
                for file in files:

                    # Update output text to show which directory is being searched
                    self.output.setFixedHeight(300)
                    self.output.setText("Now searching. Please do not close the program...\n\n" + "Searching in: \n" + str(os.path.join(root, file)))
                    self.output.repaint()
                    
                    # Get the target path of the shortcut link
                    shortcut = winshell.shortcut(os.path.join(root, file))
                   
                    # Get the mime type of file
                    kind = filetype.guess(shortcut.path)

                    # If it does not have a mime type
                    if kind is None:
                        # Continue to the next file
                        continue
                    # If it is a .pdf or .docx file 
                    elif ("pdf" in str(kind.mime)) or (("zip" in str(kind.mime)) and (".docx" in str(file))):
                        
                        # Get the URN from the name and check if it is in the input list
                        fileURN = re.findall('\d+', file )                                

                        # Check if the URN is in the list of input URNs
                        for ur in fileURN:
                            if ur in URNS:
                                # Added file name to the output string
                                self.outString = self.outString + \
                                    (str(file) + "\n")
                                self.found.append([shortcut.path, file])
                                
                                    
                    
            # For debug purposes. Remove later
            ''' buttonReply = QMessageBox.question(
                self, 'Saved', self.outString, QMessageBox.Ok) '''

            # Display the output string on GUI
            if self.outString != "":
                self.output.setText(self.outString)
                self.output.setFixedHeight(20 * len(self.found))
                self.output.repaint()
            else:
                self.output.setText("No results. Please ensure all URNs have been entered correctly.")
                self.output.setFixedHeight(20)
                self.output.repaint()
            self.show()
            print(self.found)
        

    # Save list button handler
    def saveToTxt(self):
        
        # Only allow to export as txt if items were found
        if len(self.found) != 0:
            # Obtain the directory to save list in
            directory = str(QFileDialog.getExistingDirectory(
                self, "Select directory to save list in."))
        else:
            QMessageBox.question(
                self, 'Empty List', "The list cannot be empty in order to saving to files as shortcuts.", QMessageBox.Ok)
            print("List is empty. Abort command.")
            return

        # If a directory was selected
        if directory:
            # Save list in a txt file in chosen directory with current timestamp 
            try:
                print(self.found)
                print("Saving to txt...")
                now = datetime.now().strftime("%d%m%y_%H%M")
                fileName = directory + "/patients_list_" + now + ".txt"
                f = open(fileName, "w+")
                f.write(self.outString)
                f.close()
                print("Done...")

                # Create shortcut links in the chosen directory for easy access to PDF files
                print("Creating shortcuts...")
                for f in self.found:

                    path = os.path.join(directory, (f[1]+'.lnk'))
                    target = f[0]

                    shell = Dispatch('WScript.Shell')

                    shortcut = shell.CreateShortCut(path)
                    shortcut.Targetpath = target
                    shortcut.save()
                print("Done...")

                print(os.path.normpath(directory))
                
                QMessageBox.question(
                    self, 'Saved', "Shortcuts created in:\n" + directory + "\n\n" + "List saved to:\n" + fileName, QMessageBox.Ok)

                subprocess.Popen('explorer "'+ os.path.normpath(directory) +'\\"')
            except Exception as e:
                QMessageBox.question(
                    self, 'Error', e, QMessageBox.Ok)
                print("Error occured. Abort command.")


# Execute the program
if __name__ == '__main__':
     
    app = QApplication(sys.argv)
    
    ex = PatientFinder()

    p = QPalette()
    p.setBrush(QPalette.Window, QBrush(QColor.fromRgb(175, 206, 255)))
    ex.setPalette(p)

    sys.exit(app.exec_()
