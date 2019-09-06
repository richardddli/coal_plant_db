import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import xlwings as xw
from xlwings import Range
import prototype as pt

def create_csv():
    sht = xw.Book.caller().sheets['inputs'] #for calling parameters
    pst = xw.Book.caller().sheets['csv'] #for pasting results

    #data subsetting selections
    State = Range('F11').value  # reads in selected utility
    #Market = sht.Range('F13').value  # reads in selected utility
    Utility = Range('F20').value  # reads in selected utility
    #Utility_type = sht.Range('F22').value  # reads in selected utility
    Sierra = Range('F29').value  # reads in selected utility

    #data output selections
    Level = Range('N11').value  # reads in selected utility
    Sort = Range('N18').value  # reads in selected utility
    #file_path = sht.range('N25').value  # reads in selected utility

    #define df depending on analysis type
    #df = pt.plants
    ### DIRECTLY READS IN CURRENT VERSION OF DATABASE ###
    cwd = os.path.dirname(os.path.realpath(__file__))
    df = pd.read_csv(os.path.join(cwd, 'current database', 'plants_df.csv'))

    #analysis functions
    if State:
        newdf = pt.select_by_attribute(df, 'State', State)
    elif Utility:
        newdf = pt.select_by_attribute(df, 'Utility Name', Utility)
    else:
        newdf = df

    if Sierra:
        newdf = pt.select_by_attribute(newdf, 'Current Designation', Sierra)
    else:
        newdf = newdf

    if Level:
        newdf = pt.aggregate_by_level(newdf, Level)
    else:
        newdf = newdf

    if Sort:
        newdf = pt.sort_by_attribute(newdf, 'Nameplate Capacity (MW)', ascending=False)
    else:
        newdf = newdf

    #use file_path to create new csv file with data
    xw.view(newdf, sheet=pst)

if __name__ == '__main__':
    # Expects the Excel file next to this source file, adjust accordingly.
    xw.Book('UI_Mockup.xlsm').set_mock_caller()
    CreateCsv()

def create_chart():
    # read in prototype df and python visualizations
    #df = pt.plants
    ### DIRECTLY READS IN CURRENT VERSION OF DATABASE ###
    cwd = os.path.dirname(os.path.realpath(__file__))
    df = pd.read_csv(os.path.join(cwd, 'current database', 'plants_df.csv'))
    
    pb = pt.pb
    cf = pt.cf
    pr = pt.pr

    # Create a reference to the calling Excel Workbook
    sht = xw.Book.caller().sheets['inputs'] #for calling parameters
    pst = xw.Book.caller().sheets['vis'] #for pasting results

    chart = Range('V11').value

    if Range('V21').value == 'Python':
        # create visualization in python
        if chart == 'Plant Balance':
            fig = pb
        elif chart == 'CF Offset':
            fig = cf
        elif chart == 'Profits':
            fig = pr
        pst.pictures.add(fig, name='MyPlot', update=True, left=sht.range('E2').left, top=sht.range('E2').top)
    else:
        df.sort_values(by=[chart])
        largest = df.nlargest(10, chart)
        largest = largest[['Utility Name', 'Plant Name', chart]]
        pst.range('A1').options(index=False).value = largest

if __name__ == '__main__':
    # Expects the Excel file next to this source file, adjust accordingly.
    xw.Book('UI_Mockup.xlsm').set_mock_caller()
    CreateChart()