import pandas as pd
import xlwings as xw
from xlwings import Range
import partial_ownership

ui = xw.Book('UI_Mockup.xlsm')  #connect to xlsm ui within the active wd
file = pd.Series(partial_ownership.total_MW) #read in series with utilities and mw values

def index():
    ui = xw.Book.caller()
    ui.sheets['Sheet2'].Range['A1'].value = file.keys()

def capacity():
    ui = xw.Book.caller() #connects to xlsm macro
    utility = Range('C2').value #reads in selected utility
    mw = file.get(key = utility) #matches with series from partial_ownership
    Range('C3').value = mw #returns value in xlsm ui