'''
    22/01/11
    22/01/12
'''
import re, os, sys, datetime
from PyQt4 import QtCore, QtGui
import readbibcls as rbc
import readbib as rb

#import icon_word
import icon_docx
import icon_excel
import icon_setting
import icon_folder_open
#import icon_color_picker
#import icon_font_picker

import icon_font_picker01
import icon_font_picker02
import icon_color_picker01
import icon_color_picker02

scheduler_keys = ('WORD', 'EXCEL')

def get_scheduler_key(key):
    return scheduler_keys[key]
    
def get_word_scheduler_key(): return get_scheduler_key(0)
def get_excel_scheduler_key(): return get_scheduler_key(1)
    
nyear_range = 10

month_list = [
    '1' , '2' , '3', '4', '5', 
    '6' , '7' , '8', '9', '10', 
    '11', '12'
]

_find_rgb = re.compile("(\d{1,3}),\s*(\d{1,3}),\s*(\d{1,3})")

def get_rgb(c):
    c1 = _find_rgb.search(c)
    r = int(c1.group(1))
    g = int(c1.group(2))
    b = int(c1.group(3))
    assert 0 <= r <=255    
    assert 0 <= g <=255
    assert 0 <= b <=255
    return rbc.color(r, g, b)

class QKeyButton(QtGui.QPushButton):
    def __init__(self, key):
        super(QKeyButton, self).__init__()
        self.key = key
        
class QReadBible(QtGui.QWidget):
    def __init__(self):
        super(QReadBible, self).__init__()
        self.initUI()

    def common_var(self):
        self.excel_info = rbc.excel_info()
        self.word_info = rbc.word_info()
        
    def initUI(self):
        self.common_var()
        layout = QtGui.QFormLayout()

        file_group = QtGui.QGroupBox('Output')
        file_layout = QtGui.QGridLayout()
        file_layout.addWidget(QtGui.QLabel('File'), 1, 0) 
        self.file_name  = QtGui.QLineEdit("rb300")
        file_layout.addWidget(self.file_name, 1, 1)
        
        file_layout.addWidget(QtGui.QLabel('Dest'), 2, 0)
        self.save_directory_path  = QtGui.QLineEdit(os.getcwd())
        self.save_directory_path_btn = QtGui.QPushButton('', self)
        self.save_directory_path_btn.clicked.connect(self.change_save_folder)
        self.save_directory_path_btn.setIcon(QtGui.QIcon(QtGui.QPixmap(icon_folder_open.table)))
        self.save_directory_path_btn.setIconSize(QtCore.QSize(16,16))
        self.save_directory_path_btn.setToolTip('save folder')

        file_layout.addWidget(self.save_directory_path, 2, 1)
        file_layout.addWidget(self.save_directory_path_btn, 2, 2)
        file_layout.setContentsMargins(3,3,3,3)
        file_layout.setSpacing(3) 
        file_group.setLayout(file_layout)
        
        date_group = QtGui.QGroupBox('Date')
        current_year = datetime.date.today().year
        years = [*range(current_year, current_year+nyear_range, 1)]
        year_range = [str(x) for x in years]
        
        date_grid = QtGui.QGridLayout()
        date_grid.addWidget(QtGui.QLabel("Year" ), 0, 0)
        date_grid.addWidget(QtGui.QLabel("Start"), 0, 1)
        date_grid.addWidget(QtGui.QLabel("End"  ), 0, 2)
        self.date_year   = QtGui.QComboBox()
        self.date_month1 = QtGui.QComboBox()
        self.date_month2 = QtGui.QComboBox()
        
        self.date_year  .addItems(year_range)
        self.date_month1.addItems(month_list)
        self.date_month2.addItems(month_list)

        date_grid.addWidget(self.date_year, 1, 0)
        date_grid.addWidget(self.date_month1, 1,1)
        date_grid.addWidget(self.date_month2, 1,2)  
        date_grid.setContentsMargins(3,3,3,3)
        date_grid.setSpacing(3) 
        date_group.setLayout(date_grid)
        
        excel_group = QtGui.QGroupBox('Excel')
        excel_grid  = QtGui.QGridLayout()
        
        excel_grid.addWidget(QtGui.QLabel("Columns"), 1, 0)
        self.excel_columns = QtGui.QLineEdit("5")
        self.excel_columns.setFixedWidth(30)
        excel_grid.addWidget(self.excel_columns, 1,1)
        
        excel_grid.addWidget(QtGui.QLabel("Auto Fit"), 2, 0)
        self.excel_autofit = QtGui.QCheckBox()
        excel_grid.addWidget(self.excel_autofit, 2, 1)
        
        excel_grid.addWidget(QtGui.QLabel("Delay(sec)"), 2, 2)
        self.excel_delay = QtGui.QLineEdit("1")
        self.excel_delay.setFixedWidth(40)
        excel_grid.addWidget(self.excel_delay, 2,3)
        
        excel_grid.addWidget(QtGui.QLabel("Name"), 3, 0)
        self.excel_font_name  = QtGui.QLineEdit(self.excel_info.font.font_name)
        excel_grid.addWidget(self.excel_font_name, 3, 1, 1, 2)
        self.excel_font_btn = QtGui.QPushButton('', self)
        self.excel_font_btn.setIcon(QtGui.QIcon(QtGui.QPixmap(icon_font_picker01.table)))
        self.excel_font_btn.setIconSize(QtCore.QSize(16,16))
        self.excel_font_btn.clicked.connect(self.choose_excel_font)
        excel_grid.addWidget(self.excel_font_btn, 3, 3)
        
        excel_grid.addWidget(QtGui.QLabel("Color"), 4, 0)
        self.excel_font_color = QtGui.QLineEdit(str(self.excel_info.font.font_color))
        excel_grid.addWidget(self.excel_font_color, 4, 1, 1, 2)
        self.excel_font_color_btn = QtGui.QPushButton('', self)
        self.excel_font_color_btn.setIcon(QtGui.QIcon(QtGui.QPixmap(icon_color_picker02.table)))
        self.excel_font_color_btn.setIconSize(QtCore.QSize(16,16))
        self.excel_font_color_btn.clicked.connect(self.choose_excel_font_color)
        excel_grid.addWidget(self.excel_font_color_btn, 4, 3)
        
        excel_grid.addWidget(QtGui.QLabel("Sunday"), 5, 0)
        self.excel_sunday = QtGui.QCheckBox()
        self.excel_sunday.setChecked(self.excel_info.sunday)
        excel_grid.addWidget(self.excel_sunday, 5, 1)
        
        excel_grid.setContentsMargins(2,2,2,2)
        excel_grid.setSpacing(2)
        
        excel_group.setLayout(excel_grid)
        
        # WORD group
        word_group = QtGui.QGroupBox('Word')
        word_grid  = QtGui.QGridLayout()

        word_grid.addWidget(QtGui.QLabel("Rows"), 1, 0)
        self.word_rows = QtGui.QLineEdit("25")
        self.word_rows.setFixedWidth(40)
        word_grid.addWidget(self.word_rows, 1,1)

        word_grid.addWidget(QtGui.QLabel(" Columns"), 1, 2)
        self.word_columns = QtGui.QLineEdit("2")
        self.word_columns.setFixedWidth(40)
        word_grid.addWidget(self.word_columns, 1,3)

        word_grid.addWidget(QtGui.QLabel("Name"), 2, 0)
        self.word_font_name  = QtGui.QLineEdit(self.word_info.font.font_name)
        word_grid.addWidget(self.word_font_name, 2, 1, 1, 2)
        self.word_font_btn = QtGui.QPushButton('', self)
        self.word_font_btn.setIcon(QtGui.QIcon(QtGui.QPixmap(icon_font_picker01.table)))
        self.word_font_btn.setIconSize(QtCore.QSize(16,16))
        self.word_font_btn.clicked.connect(self.choose_word_font)
        word_grid.addWidget(self.word_font_btn, 2, 3)
        
        word_grid.addWidget(QtGui.QLabel("Color"), 3, 0)
        self.word_font_color = QtGui.QLineEdit(str(self.word_info.font.font_color))
        word_grid.addWidget(self.word_font_color, 3, 1, 1, 2)
        self.word_font_color_btn = QtGui.QPushButton('', self)
        self.word_font_color_btn.setIcon(QtGui.QIcon(QtGui.QPixmap(icon_color_picker02.table)))
        self.word_font_color_btn.setIconSize(QtCore.QSize(16,16))
        self.word_font_color_btn.clicked.connect(self.choose_word_font_color)
        word_grid.addWidget(self.word_font_color_btn, 3, 3)
        
        word_grid.addWidget(QtGui.QLabel("Sunday"), 4, 0)
        self.word_sunday = QtGui.QCheckBox()
        self.word_sunday.setChecked(self.word_info.sunday)
        word_grid.addWidget(self.word_sunday, 4, 1)
        
        word_grid.setContentsMargins(2,2,2,2)
        word_grid.setSpacing(2)
        
        word_group.setLayout(word_grid)

        isz = 32
        run_layout = QtGui.QHBoxLayout()
        self.run_word_btn = QKeyButton(get_word_scheduler_key())
        self.run_word_btn.setIcon(QtGui.QIcon(QtGui.QPixmap(icon_docx.table)))
        self.run_word_btn.setIconSize(QtCore.QSize(isz,isz))
        self.run_word_btn.clicked.connect(self.create_bible_reading_schedule)
        self.run_word_btn.setToolTip('Word Schedule')
    
        self.run_excel_btn = QKeyButton(get_excel_scheduler_key())
        self.run_excel_btn.setIcon(QtGui.QIcon(QtGui.QPixmap(icon_excel.table)))
        self.run_excel_btn.setIconSize(QtCore.QSize(isz,isz))
        self.run_excel_btn.clicked.connect(self.create_bible_reading_schedule)
        self.run_excel_btn.setToolTip('Excel Schedule')
    
        self.run_setting_btn = QtGui.QPushButton('', self)
        self.run_setting_btn.setIcon(QtGui.QIcon(QtGui.QPixmap(icon_setting.table)))
        self.run_setting_btn.setIconSize(QtCore.QSize(isz,isz))
        #self.run_setting_btn.clicked.connect(self.create_shadow_text)
        self.run_setting_btn.setToolTip('Settings')
        
        run_layout.addWidget(self.run_word_btn)
        run_layout.addWidget(self.run_excel_btn)
        run_layout.addWidget(self.run_setting_btn)
        
        only_int = QtGui.QIntValidator()
        self.excel_columns.setValidator(only_int)
        self.excel_delay.setValidator(only_int)
        self.word_rows.setValidator(only_int)
        self.word_columns.setValidator(only_int)
        
        layout.addRow(file_group)
        layout.addRow(date_group)
        layout.addRow(word_group)
        layout.addRow(excel_group)
        layout.addRow(run_layout)
        self.setLayout(layout)
        self.show()
        
    def choose_color(self, col):
        new_col = QtGui.QColorDialog.getColor(QtGui.QColor(col.r, col.g, col.b))
        if new_col.isValid():
            r,g,b,a = new_col.getRgb()
            col.r, col.g, col.b = r, g, b
            return True
        else: return False
        
    def choose_font(self, font):
        new_font, valid = QtGui.QFontDialog.getFont()
        
        if valid:
            font.font_name = new_font.family()
            return True
        else: return False
            
    def choose_excel_font(self):
        if self.choose_font(self.excel_info.font):
            self.excel_font_name.setText(self.excel_info.font.font_name)
        
    def choose_excel_font_color(self):
        if self.choose_color(self.excel_info.font.font_color):
            self.excel_font_color.setText(str(self.excel_info.font.font_color))

    def choose_word_font(self):
        if self.choose_font(self.word_info.font):
            self.word_font_name.setText(self.word_info.font.font_name)
        
    def choose_word_font_color(self):
        if self.choose_color(self.word_info.font.font_color):
            self.word_font_color.setText(str(self.word_info.font.font_color))
             
    def create_bible_reading_schedule(self):
        btn = self.sender()
        
        if btn.key == get_word_scheduler_key():
            self.word_info.fname = "%s.docx"%self.file_name.text()
            self.word_info.fpath = self.save_directory_path.text()
            self.word_info.year  = int(self.date_year.currentText())
            self.word_info.month1= int(self.date_month1.currentText())
            self.word_info.month2= int(self.date_month2.currentText())
            self.word_info.nrow  = int(self.word_rows.text())
            self.word_info.ncol  = int(self.word_columns.text())
            self.word_info.sunday= self.word_sunday.isChecked()
            
            rb.create_bible_reading_schedule_word(
                os.path.join(self.word_info.fpath, self.word_info.fname),
                self.word_info.year  ,
                self.word_info.month1,
                self.word_info.month2,
                self.word_info.nrow,
                self.word_info.ncol  ,
                self.word_info.sunday)
        
        elif btn.key == get_excel_scheduler_key():
            self.excel_info.fname = "%s.xlsx"%self.file_name.text()
            self.excel_info.fpath = self.save_directory_path.text()
            self.excel_info.year  = int(self.date_year.currentText())
            self.excel_info.month1= int(self.date_month1.currentText())
            self.excel_info.month2= int(self.date_month2.currentText())
            self.excel_info.ncol  = int(self.excel_columns.text())
            self.excel_info.autofit   = self.excel_autofit.isChecked()
            self.excel_info.delay_time= int(self.excel_delay.text())
            self.excel_info.sunday= self.excel_sunday.isChecked()
        
            rb.create_bible_reading_schedule_excel(
                os.path.join(self.excel_info.fpath, self.excel_info.fname),
                self.excel_info.year  ,
                self.excel_info.month1,
                self.excel_info.month2,
                self.excel_info.ncol  ,
                self.excel_info.autofit,
                self.excel_info.delay_time,
                self.excel_info.sunday)
            
    def change_save_folder(self):
        startingDir = os.getcwd() 
        path = QtGui.QFileDialog.getExistingDirectory(None, 'Save folder', startingDir, 
        QtGui.QFileDialog.ShowDirsOnly)
        if not path: return
        self.save_directory_path.setText(path)
    
def main(): 
    app = QtGui.QApplication(sys.argv)
    QtGui.QApplication.setStyle(QtGui.QStyleFactory.create(u'Plastique'))
    run = QReadBible()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
        