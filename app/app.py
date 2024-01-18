from PyQt6.QtCore import QSize, Qt
from PyQt6.QtWidgets import QApplication, QMainWindow, QComboBox, QPushButton, QVBoxLayout, QWidget, QLabel, QLineEdit, QFileDialog, QDialog
import functools
import sys
from util import create_template

class UploadButton(QPushButton):
    def get_filename(self):
        print(self.f_name)
        return self.f_name
    
    # def return_func_wrapper(self):
    #     return functools.partial(self.on_clicked)

    def upload(self, main, f_name):
        self.f_name = QFileDialog.getOpenFileName(self,
                                                  "Open File",
                                                  "${HOME}",
                                                  "All files(*)") # filter for excel worksheets later
        # this stores file
        # try to get root name

        f_name_split = self.f_name[0].split("/")
        f_name.setText("File name: " + f_name_split[-1])
        main.file_dict[self.id[7:len(self.id)]] = self.f_name
        
    def __init__(self, label):
        super().__init__(label)
        self.id = label
        self.f_name = ""

class MainWindow(QMainWindow):
    # store directories
    def store_si(self):
        # validate later
        self.si_no = self.si_input.text()
    
    def submit(self):
        # check first if input is complete

        # generate file here using util
        self.si_no = self.si_input.text()

        # then call util function here
        is_success = create_template(self.file_dict.get('Lazada SOA File')[0], self.file_dict.get('QNE Sales Order Report File')[0], self.file_dict.get('QNE Sales Invoice Report File')[0], self.file_dict.get('QNE Sales Order Register File')[0])
        
        dlg = QDialog(self)
        dlg.setWindowTitle("Success")
        
        
        if is_success:
            message = QLabel("Convert Successful")
        else:
            message = QLabel("Convert Unsuccessful")

        dlg_layout = QVBoxLayout()
        dlg_layout.addWidget(message)
        dlg.setLayout(dlg_layout)

        dlg.exec()
        # then prompt output file later.
        
    def __init__(self):
        super().__init__()

        #dialog_wrap = functools.partial(self.on_clicked, self)
        self.f_name = ""
        self.file_dict = dict()

        # set window defaults
        self.setWindowTitle("Sales Invoice Uploader")
        self.setFixedSize(QSize(450, 450))
        
        self.si_label = QLabel("SI Number Start")
        self.si_input = QLineEdit()
        self.si_input.textChanged.connect(self.store_si)

        # add elements

        self.month_label = QLabel("Month")
        self.month_selection = QComboBox()
        self.month_selection.addItems(['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'])

        #self.soa_label = QLabel("Lazada SOA File")
        #self.soa_upload = QFileDialog()

        self.sorep_label = QLabel("QNE Sales Order Report File")
        self.sorep_upload = UploadButton("Upload QNE Sales Order Report File")
        self.sorep_fname = QLabel("")
        self.sorep_upload.clicked.connect(functools.partial(self.sorep_upload.upload, self, self.sorep_fname))
        # receive file

        self.sirep_label = QLabel("QNE Sales Invoice Report File")
        self.sirep_upload = UploadButton("Upload QNE Sales Invoice Report File")
        self.sirep_fname = QLabel("")
        self.sirep_upload.clicked.connect(functools.partial(self.sirep_upload.upload, self, self.sirep_fname))

        self.soreg_label = QLabel("QNE Sales Order Register File")
        self.soreg_upload = UploadButton("Upload QNE Sales Order Register File")
        self.soreg_fname = QLabel("")
        self.soreg_upload.clicked.connect(functools.partial(self.soreg_upload.upload, self, self.soreg_fname))

        self.soa_label = QLabel("Lazada SOA File")
        self.soa_upload = UploadButton("Upload Lazada SOA File")
        self.soa_upload_fname = QLabel("")
        self.soa_upload.clicked.connect(functools.partial(self.soa_upload.upload, self, self.soa_upload_fname))

        submit_button = QPushButton("Submit")
        submit_button.clicked.connect(self.submit)

        layout = QVBoxLayout()
        layout.addWidget(self.si_label)
        layout.addWidget(self.si_input)

        layout.addWidget(self.month_label)
        layout.addWidget(self.month_selection)

        layout.addWidget(self.soa_label)
        layout.addWidget(self.soa_upload)
        layout.addWidget(self.soa_upload_fname)

        layout.addWidget(self.sorep_label)
        layout.addWidget(self.sorep_upload)
        layout.addWidget(self.sorep_fname)

        layout.addWidget(self.sirep_label)
        layout.addWidget(self.sirep_upload)
        layout.addWidget(self.sirep_fname)

        layout.addWidget(self.soreg_label)
        layout.addWidget(self.soreg_upload)
        layout.addWidget(self.soreg_fname)


        layout.addWidget(submit_button)
        
        container = QWidget()
        container.setLayout(layout)
        self.setMenuWidget(container)

# Create instance of Qt application
app = QApplication(sys.argv)

# Create window instance
window = MainWindow()
window.show() # show window

# start application loop
app.exec()