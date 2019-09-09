# -*- coding: utf-8 -*-

import os
import pandas as pd
import xlwings as xw
from xlwings import Range
import prototype as pt
import matplotlib.pyplot as plt
import numpy as np

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
    Level = Range('N10').value  # reads in selected utility
    Sort = Range('N15').value  # reads in selected utility
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

    if Sierra == 'No Retirement':
        newdf = pt.select_remaining_plants(newdf)
    elif Sierra:
        newdf = pt.select_by_attribute(newdf, 'Current Designation', Sierra)
    else:
        newdf = newdf

    if Level:
        newdf = pt.aggregate_by_level(newdf, Level)
    else:
        newdf = newdf

    if Sort:
        newdf = pt.sort_by_attribute(newdf, Sort, ascending=False)
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
    # Create a reference to the calling Excel Workbook
    sht = xw.Book.caller().sheets['inputs'] #for calling parameters
    pst = xw.Book.caller().sheets['vis'] #for pasting results

    #data subsetting selections
    state = Range('F10').value  # reads in selected utility
    #Market = sht.Range('F12').value  # reads in selected utility
    utility = Range('F14').value  # reads in selected utility
    #Utility_type = sht.Range('F16').value  # reads in selected utility
    sierra = Range('F21').value  # reads in selected utility

    #data output selections
    label = Range('V10').value #reads in Top 10 Chart label value
    xaxis = Range('V15').value #reads in Plant Balance Chart xaxis value
    output = Range('V20').value #Excel or Python

    if label:
        # Top 10 Chart

        # read in prototype df and python visualizations
        #df = pt.plants
        ### DIRECTLY READS IN CURRENT VERSION OF DATABASE ###
        cwd = os.path.dirname(os.path.realpath(__file__))
        df = pd.read_csv(os.path.join(cwd, 'current database', 'plants_df.csv'))

        # analysis functions
        if state:
            newdf = pt.select_by_attribute(df, 'State', state)
        elif utility:
            newdf = pt.select_by_attribute(df, 'Utility Name', utility)
        else:
            newdf = df

        if sierra == 'No Retirement':
            newdf = pt.select_remaining_plants(newdf)
        elif sierra:
            newdf = pt.select_by_attribute(newdf, 'Current Designation', sierra)
        else:
            newdf = newdf

        if output == 'Python':
            #fig = pt.graph_top(newdf, label, 10, label, label) #note - not functions uses df and not newdf for some reason
            if label not in newdf.columns:
                print('%s not in columns, try again' % (label))
            ascending = False if label == 'Plant Balance' else True
            newdf = newdf.sort_values(label, ascending=ascending)
            newdf = newdf.iloc[:10]
            fig, ax = plt.subplots()
            y_ticks = np.arange(len(newdf))
            newdf.loc[newdf['Generator ID'] == 'ALL', 'Generator ID'] = " "
            y_tick_labels = newdf['Plant Name'] + ' ' + newdf['Generator ID']

            ax.barh(y_ticks, width=newdf[label])
            ax.set_yticks(y_ticks)
            ax.set_yticklabels(y_tick_labels, rotation='horizontal')
            ax.set_xlabel(label)
            ax.set_title(label)
            #return fig
            pst.pictures.add(fig, name='MyPlot', update=True, left=sht.range('E2').left, top=sht.range('E2').top)
        elif output == 'Excel':
            ascending = False if label == 'Plant Balance' else True
            newdf = newdf.sort_values(label, ascending=ascending)
            newdf = newdf.iloc[:10]
            pst.range('A1').options(index=False).value = newdf
        else:
            print('enter Python or Excel')

    elif xaxis:
        # read in prototype df and python visualizations
        #df = pt.plants
        ### DIRECTLY READS IN CURRENT VERSION OF DATABASE ###
        cwd = os.path.dirname(os.path.realpath(__file__))
        df = pd.read_csv(os.path.join(cwd, 'current database', 'plants_df.csv'))

        if output == 'Python':
            fig = pt.plot_plant_balance(df, xaxis, designations=None, labels=True, title=None)
            pst.pictures.add(fig, name='MyPlot', update=True, left=sht.range('E2').left, top=sht.range('E2').top)
        #elif output == 'Excel':
        else:
            print('enter Python or Excel')
    else:
        print('select chart type')

if __name__ == '__main__':
    # Expects the Excel file next to this source file, adjust accordingly.
    xw.Book('UI_Mockup.xlsm').set_mock_caller()
    CreateChart()