# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""

import os
import pandas as pd
from fuzzywuzzy import fuzz, process
import pdb
import matplotlib.pyplot as plt
import numpy as np

## This script generates a prototype database with the following attributes
## for each coal plant:
## (1) Ownership (how many owners, what types of owners)
## (2) Remaining plant balance (taken directly from Depreciation model)
## (3) Self committing vs Market committing (taken from JD Sierra Club paper)

# import EIA 860
cwd = os.path.dirname(os.path.realpath(__file__))
plant_cols = ['Plant Name', 'Plant Code', 'Utility Name', 
              'Generator ID', 'Technology', 'Ownership']
#plant_cols = ['Plant Name', 'Plant Code', 'Utility Name', 'Utility ID', 'State',
#              'Generator ID', 'Technology', 'Prime Mover', 'Unit Code',
#              'Ownership', 'Nameplate Capacity (MW)', 'Nameplate Power Factor',
#              'Operating Month', 'Operating Year', 'Planned Retirement Month', 
#              'Planned Retirement Year', 'Sector Name']
plants = pd.read_csv(os.path.join(cwd, 'eia860', '3_1_Generator_Y2017.csv'), skiprows=1, usecols=plant_cols)
owners = pd.read_csv(os.path.join(cwd, 'eia860', '4___Owner_Y2017.csv'), skiprows=1)

# selecting only coal plants for now
#plants2 = plants[plants['Technology'] == 'Conventional Steam Coal']
plants2=plants[plants['Plant Name'].notnull()]

# import depreciation model outputs
balance_or = pd.read_csv(os.path.join(cwd, 'plant balance.csv'))
balance = balance_or[balance_or['Current Net Plant Balance Incl. Removal Net of Salvage ($)'].notnull()
 & balance_or['Current Net Plant Balance Incl. Removal Net of Salvage ($)'] != 0]

unit_list = pd.read_csv(os.path.join(cwd, 'master unit list.csv'), skiprows=2)
unit_list = unit_list[unit_list['Plant Code'].notnull()]
unit_list['Plant Code'] = unit_list['Plant Code'].astype(int).astype(str)

# matching plant & unit between depreciation model and EIA 860
plant_names = sorted(list(set(plants2['Plant Name'])))
for i, row in balance.iterrows():
    if '_' in row['Plant ID']:
        match = unit_list[unit_list['Plant Generator ID'] == row['Plant ID']].iloc[0]
        eia_name = process.extractOne(match['Plant Name'], plant_names)[0]
        plants2.loc[(plants2['Plant Name'] == eia_name) & 
                    (plants2['Generator ID'] == match['Generator ID']), 
                    'Plant Balance'] = row['Current Net Plant Balance Incl. Removal Net of Salvage ($)']
    else:
        match = unit_list[unit_list['Plant Code'] == row['Plant ID']].iloc[0]
        eia_name = process.extractOne(match['Plant Name'], plant_names)[0]
        new_row = plants2[plants2['Plant Name'] == eia_name].iloc[0]
        new_row['Generator ID'] = 'ALL'
        new_row['Plant Balance'] = row['Current Net Plant Balance Incl. Removal Net of Salvage ($)']

        plants2 = plants2.append(new_row)
#        plants2.loc[plants2['Plant Name'] == eia_name, 
 #                   'Plant Balance'] = row['Current Net Plant Balance Incl. Removal Net of Salvage ($)']

# select only coal plants or plants with plant balances
plants2 = plants2[(plants2['Technology'] == 'Conventional Steam Coal') |
                  (plants2['Technology'] == 'Coal Integrated Gasification Combined Cycle') |
                  (plants2['Plant Balance'].notnull())]


# adding multiple ownership information
for i, row in plants2[plants2['Ownership'] != 'S'].iterrows():
    own = owners[(owners['Plant Code'] == row['Plant Code']) & 
                 (owners['Generator ID'] == row['Generator ID'])]
    own.reset_index(inplace=True)
    count = 1
    for j, r in own.iterrows():
        if r['Percent Owned'] > .1:
            plants2.loc[i, 'Owner #' + str(count)] = r['Owner Name']
            plants2.loc[i, '% Owned #' + str(count)] = r['Percent Owned']
            count += 1


## ADD SELF COMMITTING DATA
sc = pd.read_csv(os.path.join(cwd, 'self_committing.csv'))

# fuzzy matching for plant names between self-committing dataset and EIA.
plant_names = sorted(list(set(plants2['Plant Name'])))
name_mapping = {}
for i, row in sc.iterrows():
    name_mapping[row['Plant Name']] = process.extractOne(row['Plant Name'], plant_names, scorer = fuzz.partial_ratio)[0]

# the top match from above was chosen for each except for the following exceptions,
# which were manually matched:
name_mapping["Scottsbluff ST Plant"] = "Western Sugar Coop - Scottsbluff"
name_mapping["Grand River Energy Center (GRDA)"] = "GREC"

# add self-committing data to main dataset
for i, row in sc.iterrows():
    plants2.loc[(plants2['Plant Name'] == name_mapping[row['Plant Name']]) & 
                (plants2['Generator ID'] == row['Unit No']), "Profits"] = \
    row['2-year Cumulative Profits (Losses) ($ Millions)']
    plants2.loc[(plants2['Plant Name'] == name_mapping[row['Plant Name']]) & 
                (plants2['Generator ID'] == row['Unit No']), "CF Offset"] = \
    row['Expected vs Actual']
              
    
    
    
## CREATING VISUALIZATIONS
# graph the top 10 plants by any label
def graph_top(label, number, xlabel, title):
    if label not in plants2.columns:
        print('%s not in columns, try again'%(label))
    ascending = False if label=='Plant Balance' else True
    df = plants2.sort_values(label, ascending=ascending)
    df = df.iloc[:10]
    fig, ax = plt.subplots()
    y_ticks = np.arange(len(df))
    df.loc[df['Generator ID'] == 'ALL', 'Generator ID'] = " "
    y_tick_labels = df['Plant Name'] + ' ' + df['Generator ID']
    
    ax.barh(y_ticks, width=df[label])
    ax.set_yticks(y_ticks)
    ax.set_yticklabels(y_tick_labels, rotation='horizontal')
    ax.set_xlabel(xlabel)
    ax.set_title(title)
    return fig

# graphing the top 10 plant balances; top 10 self-committing plants
pb = graph_top('Plant Balance', 10, 'Undepreciated Plant Balance ($)', 'Candidates for Securitization (WI, KY)')
cf = graph_top('CF Offset', 10, 'Difference of Expected vs. Actual Capacity Factor (%)', 'Self-Committing Plants (SPP)')
pr = graph_top('Profits', 10, 'Profits/Losses in 2015-17 ($ Millions)', 'Least Profitable Plants (SPP)')
    

plan_tool = pd.read_csv('/Users/richardli/Documents/Coal Database/SC planning tool.csv', skiprows=2)
name_mapping = pd.read_csv('/Users/richardli/Documents/Coal Database/sc-eia plant mapping.csv', skiprows=2)
coal_plants = plants[(plants['Technology'] == 'Conventional Steam Coal') |
                     (plants['Technology'] == 'Coal Integrated Gasification Combined Cycle')]
plant_names = list(set(coal_plants['Plant Name']))

remaining = plan_tool.loc[plan_tool['Current Designation'] == 'Other Remaining']

for i, row in remaining.iterrows():
    if name_mapping['Sierra Club'].str.contains(row['Plant Name']).any():
        match = name_mapping.loc[name_mapping['Sierra Club'] == row['Plant Name'], 'EIA 860'].iloc[0]
    else:
        match = process.extractOne(row['Plant Name'], plant_names)[0]
    units = plants[plants['Plant Name'] == match]
    if units['Generator ID'].str.contains(row['GEN ID']).any():
        plants.loc[(plants['Plant Name'] == match) & (plants['Generator ID'] == row['GEN ID']), 'Current Designation'] = row['Current Designation']
    
