'''



'''
import datetime

class color:
    def __init__(self, r=0, g=0, b=0):
        self.r = r
        self.g = g
        self.b = b
    def __str__(self):
        return "%3d,%3d,%3d"%(self.r, self.g, self.b)
        
class font_info:
    def __init__(self):
        self.font_name  = "Calibri"
        self.font_size  = 11
        self.font_bold  = False
        self.font_color = color()

class schedule_info:
    def __init__(self):
        self.fname  = ""
        self.fpath  = ""
        self.year   = datetime.date.today().year
        self.month1 = 1    # start month
        self.month2 = 12   # end month
        self.sunday = False

class word_info(schedule_info):
    def __init__(self):
        super(word_info, self).__init__()
        self.nrow      = 30 
        self.ncol      = 2
        self.font      = font_info()
        
class excel_info(schedule_info):
    def __init__(self):
        super(excel_info, self).__init__()
        self.ncol      = 7
        self.font      = font_info()        
        self.auto_fit  = True  # run Excel to fit the columns
        self.delay_time= 1     # sec
        