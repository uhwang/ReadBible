'''



'''
from datetime import date

class schedule_info:
    def __init__(self):
        today = date.today()
        self.year = today.year
        self.include_sunday = False

class color:
    def __init__(self, r=0, g=0, b=0):
        self.r = r
        self.g = g
        self.b = b
    def __str__(self):
        return "%3d,%3d,%3d"%(self.r, self.g, self.b)
        
class font_info:
    def __init__(self)"
        self.font_name  = "Calibri"
        self.font_size  = 11
        self.font_bold  = False
        self.font_color = color()
   
class word_info(schedule_info):
    def __init__(self):
        super.__init__()
        self.start_mon = 1
        self.end_mon   = 12 
        self.nrow      = 30 
        self.ncol      = 2
        
class excel_info(schedule_info):
    def __init__(self):
        super.__init__()
        self.start_mon = 1
        self.end_mon   = 12 
        self.ncol      = 7
        
        self.auto_fit  = True  # run Excel to fit the columns
        self.delay_time= 1     # sec
        