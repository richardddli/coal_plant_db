import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import xlwings as xw
from xlwings import Range
import partial_ownership
import prototype

#ui = xw.Book('UI_Mockup.xlsm')  #connect to xlsm ui within the active wd

#read in partial ownership series, turn series into a dataframe
file = pd.Series(partial_ownership.total_MW) #read in series with utilities and mw values
file = file.to_frame()
file.rename(columns={0: 'Capacity (MW)'}, inplace=True)

#read in prototype df
main = prototype.plants2
pb = prototype.pb

def capacity():
    ui = xw.Book.caller() #connects to xlsm macro
    utility = Range('C2').value #reads in selected utility
    Range('C3').value = file.get_value(utility, 'Capacity (MW)') #returns value in xlsm ui

if __name__ == '__main__':
    # Expects the Excel file next to this source file, adjust accordingly.
    xw.Book('UI_Mockup.xlsm').set_mock_caller()
    HelloWorld()

def get_figure():
    main.set_index('Plant Name', drop=True, inplace=True)
    largest = main.nlargest(10, 'Plant Balance')
    ax = largest['Plant Balance'].plot(kind='bar', figsize=(3,3), title='Top 10 Plant Balances')
    fig = ax.get_figure()
    return fig

def plotcap():
    # Create a reference to the calling Excel Workbook
    sht = xw.Book.caller().sheets['Sheet1'] #connects to xlsm macro

    if Range('B6').value == 'Python':
        #reate visualization in python
        #fig = get_figure()
        fig = pb
        sht.pictures.add(fig, name='MyPlot', update=True, left=sht.range('E2').left, top=sht.range('E2').top)
    else:
        # Get dataframe selection and show it in Excel
        main.sort_values(by=['Plant Balance'])
        largest = main.nlargest(10, 'Plant Balance')
        largest = largest[['Utility Name', 'Plant Name', 'Plant Balance']]
        sht.range('A12').value = largest
