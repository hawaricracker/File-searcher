import os
import magic
from openpyxl import Workbook
from PyQt5 import QtWidgets
from ui4 import Ui_MainWindow

class MyWindow(QtWidgets.QMainWindow, Ui_MainWindow):
    def __init__(self):
        # Initialize the parent class
        QtWidgets.QMainWindow.__init__(self)
        self.setupUi(self)
        self.directories_textbox = self.findChild(QtWidgets.QLineEdit, "directories_textbox")
        self.directories_textbox.editingFinished.connect(lambda: self.directory_text_changed(self.directories_textbox.text()))
    
    def directory_text_changed(self,text):
        self.direct = text.split(",")

    def search_files(self):
        cnt = 0
        # get directories and extensions
        directories = self.direct
        extensions = ['exe', 'mp3', 'mp4', 'mkv', 'rar', 'jpeg', 'jpg', 'png', 'zip', '7z']
    
        # create a dictionary to store the files for each extension
        files_dict = {}
        for extension in extensions:
            files_dict[extension] = []

        files_dir = []

        # iterate over each directory and extension
        for directory in directories:
            for extension in extensions:
                # add the matching files to the dictionary
                for file in self.find_files(extension, directory):
                    cnt += 1
                    try:
                        tp = magic.from_file(file)
                        if 'executable (GUI)' in tp and extension == 'exe':
                            files_dict[extension].append([os.path.dirname(file), os.path.basename(file)])
                        elif extension != 'exe':
                            files_dict[extension].append([os.path.dirname(file), os.path.basename(file)])
                    except:
                        files_dict[extension].append([os.path.dirname(file), os.path.basename(file)])
                        continue
   
        #Add not accessible directory to array
        for directory in directories:
            for name in self.find_dir(directory):
                cnt += 1
                files_dir.append(name)

        # create a new Excel workbook
        workbook = Workbook()
        worksheet = workbook.active

        # write the headers to the worksheet
        worksheet.append(['Extension', 'Directory', 'File'])

        # iterate over the dictionary and write the data to the worksheet
        for extension, files in files_dict.items():
            for file in files:
                cnt += 1
                directory = file[0]
                filename = file[1]
                file_path = os.path.join(directory, filename)
                hyperlink = f'=HYPERLINK("{directory}","{directory}")'
                worksheet.append([extension.upper(), hyperlink, filename])

        # Add non accessible directory to excel
        for dirs in files_dir:
            cnt += 1
            directory = dirs
            hyperlink = f'=HYPERLINK("{directory}","{directory}")'
            worksheet.append([" ", hyperlink, "Error, Can't Access Folder!!!", " "])

        # save the workbook
        out = 'result.xlsx'
        workbook.save(out)
        print(cnt)
        
        # show message box
        QtWidgets.QMessageBox.information(self, "File Searcher", f"Search completed. Results saved to {out}.")

    # Function for search files
    def find_files(self, extension, directory='.'):
        for root, dirs, files in os.walk(directory):
            for file in files:
                if file.endswith(extension):
                    yield os.path.join(root, file)

    # Function for search non accessible directory 
    def find_dir(self, directory='.'):
        for root,dirs,files in os.walk(directory):
            for name in dirs:
            # Check if the directory is accessible
                try:
                    os.listdir(os.path.join(root, name))
                except :
                    yield os.path.join(root, name)

app = QtWidgets.QApplication([])
window = MyWindow()
window.show()
app.exec_()
