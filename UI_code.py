# -*- coding: utf-8 -*-

import os
import pandas as pd
import xlwings as xw
from xlwings import Range
import prototype as pt

"""
Pastes formatted "plants" dataframe from prototype.py into the csv tab on the UI.
Takes inputs from the inputs tab on the UI.
RUNTIME: ~1-5 sec
"""

def create_csv():
    sht = xw.Book.caller().sheets['inputs'] #for calling parameters
    pst = xw.Book.caller().sheets['csv'] #for pasting results

    #data subsetting selections
    State = Range('F10').value  # reads in selected utility
    #Market = sht.Range('F12').value  # reads in selected utility
    Utility = Range('F14').value  # reads in selected utility
    #Utility_type = sht.Range('F16').value  # reads in selected utility
    Sierra = Range('F21').value  # reads in selected utility

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

    if Level:
        newdf = pt.aggregate_by_level(newdf, Level)
    else:
        newdf = newdf

    if Sort:
        newdf = pt.sort_by_attribute(newdf, 'Nameplate Capacity (MW)', ascending=False)
    else:
        newdf = newdf

    if Sierra == 'No Retirement':
        newdf = pt.select_remaining_plants(newdf)
    elif Sierra:
        newdf = pt.select_by_attribute(newdf, 'Current Designation', Sierra)
    else:
        newdf = newdf

    #use file_path to create new csv file with data
    xw.view(newdf, sheet=pst)

if __name__ == '__main__':
    # Expects the Excel file next to this source file, adjust accordingly.
    xw.Book('UI_Mockup.xlsm').set_mock_caller()
    CreateCsv()

"""
Creates visualization on the vis tab on the UI.
Takes inputs from the inputs tab on the UI.
User has option to paste visualization created in python or pre-formatted table in xlsm.
"""

def create_chart():
    # read in prototype df and python visualizations
    #df = pt.plants
    ### DIRECTLY READS IN CURRENT VERSION OF DATABASE ###
    cwd = os.path.dirname(os.path.realpath(__file__))
    df = pd.read_csv(os.path.join(cwd, 'current database', 'plants_df.csv'))

    # Create a reference to the calling Excel Workbook
    sht = xw.Book.caller().sheets['inputs'] #for calling parameters
    pst = xw.Book.caller().sheets['vis'] #for pasting results

    #data subsetting selections
    State = Range('F10').value  # reads in selected utility
    #Market = sht.Range('F12').value  # reads in selected utility
    Utility = Range('F14').value  # reads in selected utility
    #Utility_type = sht.Range('F16').value  # reads in selected utility
    Sierra = Range('F21').value  # reads in selected utility

    #data output selections
    label = Range('V10').value #reads in Top 10 Chart label value
    #xaxis = Range('V15').value #reads in Plant Balance Chart xaxis value

    #Top 10 Chart
    if label:
        if Range('V21').value == 'Python':
            fig = pt.graph_top(df, label)
            pst.pictures.add(fig, name='MyPlot', update=True, left=sht.range('E2').left, top=sht.range('E2').top)
        elif Range('V21').value == 'Excel':
            ascending = False if label == 'Plant Balance' else True
            df = df.sort_values(label, ascending=ascending)
            df = df.iloc[:10]
            pst.range('A1').options(index=False).value = df
        else:
            return
    else:
        return

    """
    #Plant Balance Chart
    if xaxis:
        if Range('V21').value == 'Python':
        elif Range('V21').value == 'Excel':
        else:
            return
    else:
        return
    """

if __name__ == '__main__':
    # Expects the Excel file next to this source file, adjust accordingly.
    xw.Book('UI_Mockup.xlsm').set_mock_caller()
    CreateChart()