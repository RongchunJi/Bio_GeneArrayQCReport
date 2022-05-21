from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QMainWindow
from MainWindow_ui import Ui_MainWindow
from PyQt5.QtCore import QTimer, QDateTime
import os
from generate_report import ReportGenerater


class MainWindow(QMainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()

        # Set up the user interface from Designer.
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)

        self.log_file = open(os.getcwd()+'/log.txt', 'a')

        self.sampleQC_file_name = ''
        self.sampleQC_file_type = None
        self.midQC_file_name = ''
        self.midQC_file_type = None
        self.summary_file_name = ''
        self.plates_file_name = ''
        self.pics_dir = ''
        self.plate_name = ''

        self.ui.pb_sampleQC_open.clicked.connect(self.on_open_sample_qc_file)
        self.ui.pb_midQC_open.clicked.connect(self.on_open_mid_qc_file)
        self.ui.pb_pics_dir_open.clicked.connect(self.on_open_pics_dir)
        self.ui.pb_generation.clicked.connect(self.on_generate_report)

        self.generator = ReportGenerater()

        self.generator.log_signal.connect(self.write_log)
        self.generator.state_signal.connect(self.on_state_signal)

    def __del__(self):
        self.log_file.close()

    def write_log(self, log_str):
        # pass
        time_str = QDateTime.currentDateTime().toString("yy-MM-dd hh:mm:ss  ")
        log_str = time_str + log_str
        self.ui.te_log.append(log_str)
        self.ui.te_log.moveCursor(self.ui.te_log.textCursor().End)

        self.log_file.write(log_str + '\n')
        self.log_file.flush()

    def on_open_sample_qc_file(self):
        self.sampleQC_file_name,  self.sampleQC_file_type = \
            QtWidgets.QFileDialog.getOpenFileName(self, "选取文件", os.getcwd()+'/../', "All Files(*);;Text Files(*.txt)")

        self.ui.le_sampleQC.setText(self.sampleQC_file_name)

    def on_open_mid_qc_file(self):
        self.midQC_file_name,  self.midQC_file_type = \
            QtWidgets.QFileDialog.getOpenFileName(self, "选取文件", os.getcwd()+'/../', "All Files(*);;Text Files(*.txt)")

        self.ui.le_midQC.setText(self.midQC_file_name)

    def on_open_pics_dir(self):
        self.pics_dir = \
            QtWidgets.QFileDialog.getExistingDirectory(self, "请选择文件夹路径", os.getcwd()+'/../')

        self.ui.le_pics_dir.setText(self.pics_dir)

        self.plate_name = self.pics_dir.split('/').pop()
        self.plates_file_name = self.pics_dir+'/plates_name.txt'
        self.summary_file_name = self.pics_dir+'/'+self.plate_name[2:]+'.txt'

        self.write_log('Get plate name: ' + self.plate_name)
        self.write_log('Get summary file name: ' + self.summary_file_name)

    def on_generate_report(self):
        try:
            if '.xls' not in self.sampleQC_file_name:
                self.write_log('Error: Sample QC file error.')
                return
            if '.xls' not in self.midQC_file_name:
                self.write_log('Error: Middle QC file error.')
                return
            if 'TZ' not in self.pics_dir:
                self.write_log('Error: Resource folder error.')
                return

            self.ui.pb_generation.setEnabled(False)
            self.ui.pb_midQC_open.setEnabled(False)
            self.ui.pb_pics_dir_open.setEnabled(False)
            self.ui.pb_sampleQC_open.setEnabled(False)

            self.generator.update_file(self.plate_name, self.sampleQC_file_name, self.midQC_file_name,
                                   self.summary_file_name, self.plates_file_name, self.pics_dir)

            self.generator.start()
        except Exception as e:
            self.write_log(str(e))

    def on_state_signal(self, state_str):
        self.write_log(state_str)
        self.ui.pb_generation.setEnabled(True)
        self.ui.pb_midQC_open.setEnabled(True)
        self.ui.pb_pics_dir_open.setEnabled(True)
        self.ui.pb_sampleQC_open.setEnabled(True)
