__author__ = 'Pascal COSTA'
__website__ = 'pascalcosta.fr'
__creationDate__ = '2024-02-02'
__license__ = 'free'

import sys
import os
import subprocess
import serial
import serial.tools.list_ports
from PyQt5.QtCore import *
from PyQt5.QtWidgets import *
from PyQt5.uic import loadUi
from PyQt5 import QtGui

directory = os.path.dirname(__file__)

try:
    from ctypes import windll
    myappid = 'AVR-FLASHER'
    windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
except ImportError:
    pass


class qt(QMainWindow):
    def __init__(self):

        QMainWindow.__init__(self)
        # Load the user interface
        loadUi('./flasherui.ui', self)

        # Create application path
        self.basedir = os.path.dirname(__file__)
        newBaseDir = ''
        self.spaceInPath = False
        for letter in self.basedir:
            if letter == ' ':
                self.spaceInPath = True
            if letter == '\\':
                letter = '/'
            newBaseDir += letter
        self.basedir = newBaseDir

        # Generate the Documents path
        self.username = os.environ.get('USERNAME')
        self.documents_path = 'C:/Users/' + self.username + '/Documents/'

        # Control the AVR_SAUVEGARDE folder is existing
        if os.path.exists(self.documents_path + 'AVR_SAUVEGARDE') == False:
            os.mkdir(os.path.join(self.documents_path, 'AVR_SAUVEGARDE'))
        # Created a file to receive the copy of EEPROM
        if os.path.exists(self.documents_path + 'AVR_SAUVEGARDE/eepromsaving.hex') == False:
            f = open(self.documents_path +
                     'AVR_SAUVEGARDE/eepromsaving.hex', 'w')
            f.write('File is initied')
            f.close()
        # Initiate some variables
        self.port = 'NONE'
        self.filename = self.eepromFilename = ''
        self.getHexFile = False
        self.getEepromHexFile = False
        self.eeprom_saving_ok = False
        self.prog_injection_ok = False
        self.eeprom_injection_ok = False
        self.bauds = "1000000"
        # Connect user interface buttons
        self.set_19200_bauds.clicked.connect(self.define_bauds_19200)
        self.set_115200_bauds.clicked.connect(self.define_bauds_115200)
        self.set_1000000_bauds.clicked.connect(self.define_bauds_1000000)
        self.button_open_file.clicked.connect(self.get_hex_file)
        self.button_open_eeprom.clicked.connect(self.get_eeprom_hex_file)
        self.research_port_button.clicked.connect(self.set_connection)
        self.save_eeprom.clicked.connect(self.save_eeprom_action)
        self.flash.clicked.connect(self.inject_prog)
        self.eeprom_injection.clicked.connect(self.inject_eeprom_action)
        self.save_eeprom.setEnabled(False)
        self.save_eeprom.setStyleSheet(
            'background-color: grey;'
            'color: black;'
        )
        self.flash.setEnabled(False)
        self.flash.setStyleSheet(
            'background-color: grey;'
            'color: black;'
        )
        self.eeprom_injection.setEnabled(False)
        self.eeprom_injection.setStyleSheet(
            'background-color: grey;'
            'color: black;'
        )
        self.init_langage()

    # init_langage() method is useful for developing future multi-language functionality
    def init_langage(self):
        self.title_h1.setText('AVR-FLASHER - Mettez à jour vos projets')
        self.loadfile_label.setText(
            'Cliquez ici pour charger votre programme.hex')
        self.button_open_file.setText('Chargez votre programme.hex')
        self.loading_filename.setText('Programme ???')
        if self.getHexFile and self.getHexFile[0][-4:] != '.hex':
            self.button_open_file.setText('Uniquement un fichier .hex')
            self.loading_filename.setText(
                'Erreur, uniquement un fichier .hex ...')
            self.loading_filename.setStyleSheet(
                'background-color: #C99A70;'
                'color: black;'
                'font-size: 16px;'
            )
        if self.getHexFile and self.getHexFile[0][-4:] == '.hex':
            self.button_open_file.setText('Chargez votre programme.hex')
            self.loading_filename.setText(self.filename)
        self.loadeeprom_label.setText(
            'Cliquez ici pour charger votre EEPROM.hex')
        self.button_open_eeprom.setText('Chargez votre EEPROM.hex')
        self.loading_eeprom.setText('EEPROM ???')
        if self.getEepromHexFile and self.getEepromHexFile[0][-4:] != '.hex':
            self.button_open_eeprom.setText('Uniquement un fichier .hex')
            self.loading_eeprom.setText(
                'Erreur, uniquement un fichier .hex ...')
            self.loading_eeprom.setStyleSheet(
                'background-color: #C99A70;'
                'color: black;'
                'font-size: 16px;'
            )
        if self.getEepromHexFile and self.getEepromHexFile[0][-4:] == '.hex':
            self.button_open_eeprom.setText('Chargez votre EEPROM.hex')
            self.loading_eeprom.setText(self.eepromFilename)
        if self.port == 'NONE':
            self.save_eeprom.setText('Attente connexion !')
            self.eeprom_injection.setText('Attente fichier .hex !')
        else:
            self.save_eeprom.setText('Lire et sauvegarder la configuration')
            self.eeprom_injection.setText('Injecter la configuration')
        self.save_eeprom_label.setText(
            '1 - Prélevez la configuration actuelle')
        self.flash_label.setText('2 - Mettez à jour votre projet')
        if self.getHexFile and self.getHexFile[0][-4:] == '.hex':
            self.flash.setText('Mettre à jour votre projet')
        else:
            self.flash.setText('Attente fichier .hex !')
        self.eeprom_injection_label.setText('3 - Injecté votre configuration')

    # Manage the baudrate selector

    def init_bauds_buttons(self):
        self.set_19200_bauds.setStyleSheet(
            'background-color: lightgrey;'
            'color: black;'
            'font-size: 12px;'
        )
        self.set_115200_bauds.setStyleSheet(
            'background-color: lightgrey;'
            'color: black;'
            'font-size: 12px;'
        )
        self.set_1000000_bauds.setStyleSheet(
            'background-color: lightgrey;'
            'color: black;'
            'font-size: 12px;'
        )

    # Define baudrate to 19200 bauds
    def define_bauds_19200(self):
        self.init_bauds_buttons()
        self.bauds = "19200"
        self.set_19200_bauds.setStyleSheet(
            'background-color: #70C976;'
            'color: black;'
            'font-size: 14px;'
        )

    # Define baudrate to 115200 bauds
    def define_bauds_115200(self):
        self.init_bauds_buttons()
        self.bauds = "115200"
        self.set_115200_bauds.setStyleSheet(
            'background-color: #70C976;'
            'color: black;'
            'font-size: 14px;'
        )

    # Define baudrate to 1000000 bauds
    def define_bauds_1000000(self):
        self.init_bauds_buttons()
        self.bauds = "1000000"
        self.set_1000000_bauds.setStyleSheet(
            'background-color: #70C976;'
            'color: black;'
            'font-size: 14px;'
        )

    # Load the binary intended for flash memory
    def get_hex_file(self):
        # Create the path to downloads
        hex_path = 'C:/Users/' + self.username + '/Downloads/'

        self.filename = ''
        # Select the binary file
        self.getHexFile = QFileDialog.getOpenFileName(
            self,
            "Open file",
            hex_path,
            ".hex (*)",
        )
        # Control the loading binary and update the user interface
        if self.getHexFile and self.getHexFile[0][-4:] == '.hex':
            self.filename = self.getHexFile[0].split('/')[-1]
            self.loading_filename.setText(self.filename)
            self.loading_filename.setStyleSheet(
                'background-color: #70C976;'
                'color: black;'
                'font-size: 16px;'
            )
        else:
            self.filename = ''
            self.button_open_file.setText('Uniquement un fichier .hex')
            self.loading_filename.setText(
                'Erreur, uniquement un fichier .hex ...')

        self.prog_injection_ok = False
        self.init_langage()
        self.set_port_com_connection()

    # Load the binary intended for EEPROM memory
    def get_eeprom_hex_file(self):
        # Create the path to downloads
        hex_path = 'C:/Users/' + self.username + '/Downloads/'

        self.eepromFilename = ''
        # Select the binary file
        self.getEepromHexFile = QFileDialog.getOpenFileName(
            self,
            "Open file",
            hex_path,
            ".hex (*)",
        )
        # Control the loading binary and update the user interface
        if self.getEepromHexFile and self.getEepromHexFile[0][-4:] == '.hex':
            self.eepromFilename = self.getEepromHexFile[0].split('/')[-1]
            self.loading_eeprom.setText(self.eepromFilename)
            self.loading_eeprom.setStyleSheet(
                'background-color: #70C976;'
                'color: black;'
            )
        else:
            self.eepromFilename = ''
            self.button_open_eeprom.setText('Uniquement un fichier .hex')
            self.loading_eeprom.setText(
                'Erreur, uniquement un fichier .hex ...')

        self.eeprom_injection_ok = False
        self.init_langage()
        self.set_port_com_connection()

    # Try to find the connection if connecteur is realy connected to computer
    def set_connection(self):
        self.port = 'NONE'
        # Read and analyse USB connections
        for p in serial.tools.list_ports.comports():
            # If USB is found, update the user interface
            if 'Arduino' or 'CP210' or 'CH340' in p.description:
                self.port = p.device
                self.research_port_button.setText(p.device)
                self.research_port_button.setStyleSheet(
                    'background-color: #70C976;'
                    'color: black;'
                    'font-size: 16px;'
                )
                self.set_19200_bauds.setEnabled(False)
                self.set_115200_bauds.setEnabled(False)
                self.set_1000000_bauds.setEnabled(False)
        # If connection is false, update the user interface
        if self.port == 'NONE':
            self.research_port_button.setText('Erreur')
            self.research_port_button.setStyleSheet(
                'background-color: #EC5A24;'
                'color: black;'
                'font-size: 16px;'
            )
            self.set_19200_bauds.setEnabled(True)
            self.set_115200_bauds.setEnabled(True)
            self.set_1000000_bauds.setEnabled(True)

        self.eeprom_saving_ok = False
        self.init_langage()
        self.set_port_com_connection()

    # If connection is update, update the user interface
    def set_port_com_connection(self):
        if self.port != 'NONE':
            if self.eeprom_saving_ok == False:
                self.save_eeprom.setEnabled(True)
                self.save_eeprom.setText(
                    'Lire et sauvegarder la configuration')
                self.save_eeprom.setStyleSheet(
                    'background-color: lightgrey;'
                    'color: black;'
                )
            if self.prog_injection_ok == False and self.filename != '':
                self.flash.setEnabled(True)
                self.flash.setText('Mettre à jour votre projet')
                self.flash.setStyleSheet(
                    'background-color: lightgrey;'
                    'color: black;'
                )
            if self.eeprom_injection_ok == False and self.eepromFilename != '':
                self.eeprom_injection.setEnabled(True)
                self.eeprom_injection.setText('Injecter la configuration')
                self.eeprom_injection.setStyleSheet(
                    'background-color: lightgrey;'
                    'color: black;'
                )
        else:
            if self.eeprom_saving_ok == False:
                self.save_eeprom.setEnabled(False)
                self.save_eeprom.setText('Attente connexion !')
                self.save_eeprom.setStyleSheet(
                    'background-color: grey;'
                    'color: black;'
                )
            if self.prog_injection_ok == False:
                self.flash.setEnabled(True)
                self.flash.setText('Attente fichier .hex !')
                self.flash.setStyleSheet(
                    'background-color: grey;'
                    'color: black;'
                )
            if self.eeprom_injection_ok == False:
                self.eeprom_injection.setEnabled(False)
                self.eeprom_injection.setText('Attente fichier .hex !')
                self.eeprom_injection.setStyleSheet(
                    'background-color: grey;'
                    'color: black;'
                )

    # Read the EEPROM content
    def save_eeprom_action(self):
        if self.port != 'NONE':
            # Choose the folder to save the content of EEPROM
            savingFolder = str(
                QFileDialog.getExistingDirectory(
                    self,
                    'Open File',
                    self.documents_path + 'AVR_SAUVEGARDE/',
                    QFileDialog.DontResolveSymlinks,
                )
            )
            # Create the AVRDUDE command depending on the presence of spaces in the paths
            if self.spaceInPath == True:
                command = '"' + self.basedir + '/avrdude.exe" "-C' + self.basedir + '/avrdude.conf" -c stk500v1 -P ' + \
                    self.port + ' -patmega2560 -b' + self.bauds + \
                    ' "-Ueeprom:r:' + savingFolder + '/eepromsaving.hex:i"'
            else:
                command = self.basedir + '/avrdude.exe -C' + self.basedir + '/avrdude.conf -c stk500v1 -P ' + self.port + \
                    ' -patmega2560 -b' + self.bauds + ' "-Ueeprom:r:' + \
                    savingFolder + '/eepromsaving.hex:i"'
            # Use the following line to display the command
            # print(command)
            subprocess.run(command)
            # After execution of the command, update the user interface
            self.save_eeprom.setText('L\'EEPROM est sauvegardée')
            self.save_eeprom.setStyleSheet(
                'background-color: lightgreen;'
                'color: black;'
            )

            self.eeprom_saving_ok = True

    def inject_prog(self):
        if self.getHexFile[0][-4:] == '.hex' and self.filename != '' and self.port != 'NONE':
            if self.spaceInPath == True:
                command = '"' + self.basedir + '/avrdude.exe" "-C' + self.basedir + '/avrdude.conf" -c stk500v1 -P ' + \
                    self.port + ' -patmega2560 -b' + self.bauds + \
                    ' "-Uflash:w:' + self.getHexFile[0] + ':i"'
            else:
                command = self.basedir + '/avrdude.exe -C' + self.basedir + '/avrdude.conf -c stk500v1 -P ' + self.port + \
                    ' -patmega2560 -b' + self.bauds + \
                    ' "-Uflash:w:' + self.getHexFile[0] + ':i"'
            # Use the following line to display the command
            # print(command)
            subprocess.run(command)
            # After execution of the command, update the user interface
            self.flash.setText('Attente fichier .hex !_3')
            self.flash.setStyleSheet(
                'background-color: lightgreen;'
                'color: black;'
            )
            self.prog_injection_ok = True

    def inject_eeprom_action(self):
        if self.getEepromHexFile[0][-4:] == '.hex' and self.eepromFilename != '' and self.port != 'NONE':
            if self.spaceInPath == True:
                command = '"' + self.basedir + '/avrdude.exe" "-C' + self.basedir + '/avrdude.conf" -c stk500v1 -P ' + \
                    self.port + ' -patmega2560 -b' + self.bauds + \
                    ' "-Ueeprom:w:' + self.getEepromHexFile[0] + ':i"'
            else:
                command = self.basedir + '/avrdude.exe -C' + self.basedir + '/avrdude.conf -c stk500v1 -P ' + self.port + \
                    ' -patmega2560 -b' + self.bauds + ' "-Ueeprom:w:' + \
                    self.getEepromHexFile[0] + ':i"'
            # Use the following line to display the command
            # print(command)
            subprocess.run(command)
            # After execution of the command, update the user interface
            self.eeprom_injection.setText('L\'EEPROM est injecté')
            self.eeprom_injection.setStyleSheet(
                'background-color: lightgreen;'
                'color: black;'
            )
            self.eeprom_injection_ok = True


def run():
    app = QApplication(sys.argv)
    widget = qt()
    widget.setWindowIcon(QtGui.QIcon(os.path.join(directory, 'favicon_3.ico')))
    widget.show()

    sys.exit(app.exec_())


if __name__ == "__main__":
    run()
