# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""

import os
import pandas as pd
from fuzzywuzzy import fuzz, process
import matplotlib.pyplot as plt
import numpy as np

# define working directory where all files live
cwd = os.path.dirname(os.path.realpath(__file__))


###############################################################################
########################### BUILDING THE DATABASE #############################
###############################################################################

"""
Imports EIA 860 to construct initial database, using tables 3 (plants) and 4 (owners).
INPUTS
    cols:   list of columns to import from Table 3. If None, uses default columns.
OUTPUTS
    plants: database of generators
    owners: database of multiple ownership information for generators
RUNTIME: ~1 sec
"""
def import_eia860(cols=None):
    if cols is None:
        cols = ['Plant Name', 'Plant Code', 'Utility Name', 'Utility ID',
                'Generator ID', 'State', 'Technology', 'Ownership', 'Nameplate Capacity (MW)']
        # potential columns of interest:
        # plant_cols = ['Plant Name', 'Plant Code', 'Utility Name', 'Utility ID', 'State',
        #              'Generator ID', 'Technology', 'Prime Mover', 'Unit Code',
        #              'Ownership', 'Nameplate Capacity (MW)', 'Nameplate Power Factor',
        #              'Operating Month', 'Operating Year', 'Planned Retirement Month', 
        #              'Planned Retirement Year', 'Sector Name']
    plants = pd.read_csv(os.path.join(cwd, 'eia860', '3_1_Generator_Y2017.csv'), skiprows=1, usecols=cols)
    plants['Nameplate Capacity (MW)'] = plants['Nameplate Capacity (MW)'].astype(float)
    plants[plants['Plant Name'].notnull()]
    owners = pd.read_csv(os.path.join(cwd, 'eia860', '4___Owner_Y2017.csv'), skiprows=1)
    return plants, owners


"""
Merges RMI depreciation model to existing EIA database. Currently can be used
on both the 'plant balance clean.csv' and 'cost of electricity clean.csv'
INPUTS
    plants: current database
    cols:   list of columns to import from depreciation database. If None,
            only imports net plant balance including removal net of salvage.
OUTPUTS
    plants: updated database with plant balances
RUNTIME: ~15 sec
"""
def add_depreciation_model(plants, filename=None, cols=None):
    if cols is None:
        cols = ['Current Net Plant Balance Incl. Removal Net of Salvage ($)',
                'Retirement Year']
        new_cols = ['Plant Balance', 'Retirement Year']
    else:
        new_cols = cols
    if filename is None:
        filename = 'plant balance clean.csv'
    depr_model = pd.read_csv(os.path.join(cwd, filename))
    # do not include 'other production' rows
    depr_model = depr_model[~(depr_model['Plant ID'].str.contains('Other') | 
                        depr_model['Plant ID'].str.contains('Mining'))]
    for col in cols:
        # remove zero and nan values
        depr_model = depr_model[depr_model[col].notnull() & (depr_model[col] != 0)]
        
    for i, row in depr_model.iterrows():
        if '_' in row['Plant ID']:
            plant_id = row['Plant ID'].split('_')[0]
            unit_id = row['Plant ID'].split('_')[1]
            match = plants[(plants['Plant Code'] == float(plant_id)) & 
                           (plants['Generator ID'] == unit_id)]
            if not (match.empty):
                for col in zip(cols, new_cols):
                    plants.loc[(plants['Plant Code'] == float(plant_id)) & 
                               (plants['Generator ID'] == unit_id), col[1]] = \
                                   row[col[0]]
        else:
            new_row = plants[plants['Plant Code'] == float(row['Plant ID'])].iloc[0]
            new_row['Generator ID'] = 'ALL'
            for col in zip(cols, new_cols):
                new_row[col[1]] = row[col[0]]
            plants = plants.append(new_row)
    return plants


"""
Merges Sierra Club planning tool to database. Currently only merges "Other
Remaining" plants – the rest will require manual name matching for ~300 plants.
INPUTS
    plants:     current database (only includes coal plants and plants with balances)
    cols:       list of columns to import from planning tool. If None, will
                import Current Designation and Predicted Retirement Year.
OUTPUTS
    plants:     current database
RUNTIME: ~20 sec
"""
def add_sc_planning_tool(plants, cols=None):
    # Note for updating SC planning tool: manually updated plant names for 
    # Marshall and Columbia due to redundancy issues.
    if cols is None:
        cols = ['Current Designation', 'Predicted Retirement Year']
    plan_tool = pd.read_csv(os.path.join(cwd, 'SC planning tool.csv'), skiprows=2)
    name_mapping = pd.read_csv(os.path.join(cwd, 'sc-eia plant mapping.csv'), skiprows=2)
    coal_plants = plants[(plants['Technology'] == 'Conventional Steam Coal') |
                         (plants['Technology'] == 'Coal Integrated Gasification Combined Cycle')]
    plant_names = list(set(coal_plants['Plant Name']))
    
    remaining_groups = ['2020 Vulnerable', '2025 Vulnerable', 'Announced', 'Other Remaining']
    remaining = plan_tool[plan_tool['Current Designation'].isin(remaining_groups)]
    for i, row in remaining.iterrows():
        if name_mapping['Sierra Club'].str.contains(row['Plant Name']).any():
            match = name_mapping.loc[name_mapping['Sierra Club'] == row['Plant Name'], 'EIA 860'].iloc[0]
        else:
            match = process.extractOne(row['Plant Name'], plant_names)[0]
        units = plants[plants['Plant Name'] == match]
        if units['Generator ID'].str.contains(row['GEN ID']).any():
            for col in cols:
                plants.loc[(plants['Plant Name'] == match) & 
                           (plants['Generator ID'] == row['GEN ID']), col] = row[col]
    return plants


"""
Adds self-committing data from Joe Daniel's Sierra Club paper.
INPUTS
    plants: current database
OUTPUTS
    plants: updated database with self-committing data
RUNTIME: ~5 sec
"""
def add_self_committing(plants):
    sc = pd.read_csv(os.path.join(cwd, 'self_committing.csv'))
    
    # fuzzy matching for plant names between self-committing dataset and EIA.
    plant_names = sorted(list(set(plants['Plant Name'])))
    name_mapping = {}
    for i, row in sc.iterrows():
        name_mapping[row['Plant Name']] = process.extractOne(row['Plant Name'], plant_names, scorer = fuzz.partial_ratio)[0]
    
    # the top match from above was chosen for each except for the following exceptions,
    # which were manually matched:
    name_mapping["Scottsbluff ST Plant"] = "Western Sugar Coop - Scottsbluff"
    name_mapping["Grand River Energy Center (GRDA)"] = "GREC"
    
    # add self-committing data to main dataset
    for i, row in sc.iterrows():
        plants.loc[(plants['Plant Name'] == name_mapping[row['Plant Name']]) & 
                   (plants['Generator ID'] == row['Unit No']), "Profits"] = \
                        row['2-year Cumulative Profits (Losses) ($ Millions)']
        plants.loc[(plants['Plant Name'] == name_mapping[row['Plant Name']]) & 
                   (plants['Generator ID'] == row['Unit No']), "CF Offset"] = \
                        row['Expected vs Actual']
    return plants


""" 
Adds multiple ownership information from EIA 860 table 4.
INPUTS
    plants: current database
    owners: database of multiple ownership information
OUTPUTS
    plants: updated database with multiple ownership information
RUNTIME: ~30 sec
"""
def add_multiple_ownership(plants, owners):
    for i, row in plants[plants['Ownership'] != 'S'].iterrows():
        matches = owners[(owners['Plant Code'] == row['Plant Code']) & 
                         (owners['Generator ID'] == row['Generator ID'])]
        matches.reset_index(inplace=True)
        count = 1
        for j, row2 in matches.iterrows():
            if row2['Percent Owned'] > .1:
                plants.loc[i, 'Owner #' + str(count)] = row2['Owner Name']
                plants.loc[i, '% Owned #' + str(count)] = row2['Percent Owned']
                count += 1
    return plants


"""
Wrapper function that builds a database using all available datasets (EIA 860,
RMI depreciation model, Sierra Club planning tool, Joe Daniel's self-committing 
data, multiple ownership information)

INPUTS: none
OUTPUTS
    plants:     database of mostly coal generators with all available data
    all_plants: database of all generators with limited data
RUNETIME: ~70 sec
"""
def build_database():
    all_plants, owners = import_eia860()
    all_plants = add_depreciation_model(all_plants, filename='plant balance clean.csv')
    all_plants = add_depreciation_model(all_plants, filename='cost of electricity clean.csv', cols=['Total Cost of Electricity Excluding ADIT ($/MWh)'])
    all_plants = add_sc_planning_tool(all_plants)
    # select only coal plants or plants with plant balances (~1000 out of 20000)
    plants = all_plants[(all_plants['Technology'] == 'Conventional Steam Coal') |
                        (all_plants['Technology'] == 'Coal Integrated Gasification Combined Cycle') |
                        (all_plants['Plant Balance'].notnull()) |
                        (all_plants['Current Designation'].notnull())]
    
    plants = add_self_committing(plants)
    plants = add_multiple_ownership(plants, owners)
    return plants, all_plants



#plants, all_plants = build_database()




###############################################################################
########################## CREATING VISUALIZATIONS ############################
###############################################################################

""" Graphs the top 10 plants by any label.
TODO: finish documentation
"""
def graph_top(plants, label, number, xlabel, title):
    if label not in plants.columns:
        print('%s not in columns, try again'%(label))
    ascending = False if label=='Plant Balance' else True
    df = plants.sort_values(label, ascending=ascending)
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

# graphs the top 10 plant balances; top 10 self-committing plants
#pb = graph_top(plants, 'Plant Balance', 10, 'Undepreciated Plant Balance ($)', 
#               'Candidates for Securitization (WI, KY)')
#cf = graph_top(plants, 'CF Offset', 10, 'Difference of Expected vs. Actual Capacity Factor (%)', 
#               'Self-Committing Plants (SPP)')
#pr = graph_top(plants, 'Profits', 10, 'Profits/Losses in 2015-17 ($ Millions)', 
#               'Least Profitable Plants (SPP)')
    
"""
Graphs plant balance as a function of another variable, financial or operational.
Valid x_axis arguments currently: 'Retirement Year', 'Nameplate Capacity (MW)',
'Total Cost of Electricity Excluding ADIT ($/MWh)'.

THIS IS NOT YET FUNCTIONAL / CLEANED UP
"""
def plot_plant_balance(plants, x_axis, designations=None, labels=True, title=None):    
    if designations is None:
        designations = ['2020 Vulnerable', '2025 Vulnerable', 'Announced', 'Other Remaining']
        colors = ['indianred','darkorchid','cyan','green']
    if x_axis not in plants.columns[plants.dtypes == np.float64]:
        print("Invalid x axis. Select one of the following:")
        print(*plants.columns[plants.dtypes == np.float64], sep=', ')
        return
    plants = select_plants_with_balances(select_remaining_plants(plants))
    plants = plants[plants[x_axis].notnull()]
    fig, ax = plt.subplots()

    for designation, color in zip(designations, colors):
        subset = plants[plants['Current Designation'] == designation]
        ax.scatter(subset[x_axis], subset['Plant Balance'], c=color, label=designation)
        # this labels the points
        if labels:
            for i, txt in enumerate(subset['Plant Name'] + ' ' + subset['Generator ID']):
                ax.annotate(txt, (subset[x_axis].iloc[i]+.7, subset['Plant Balance'].iloc[i]+.5))
    ax.legend()
    if title is not None:
        ax.set_title(title)
    ax.set_xlabel(x_axis)
    ax.set_ylabel('Plant Balance ($ billions)')
    #ax.set_ylim([-.15*np.power(10,9), 1.85*np.power(10,9)])
    return fig


###############################################################################
############################### ANALYSIS TOOLS ################################
###############################################################################
"""
Selects a subset of the database by attribute (column) value(s).
NOT CURRENTLY IMPLEMENTED: selecting by market or utility type (coop, muni, etc)

INPUTS
    df:        database
    attribute: the attribute (or column) 
    values:    the value(s) to select
OUTPUTS
    df:        subset of database with selected values
RUNTIME: ~1-10 sec
EXAMPLE USAGE (can be called in series):
    subset = select_by_attribute(plants, 'State', 'AL)
    subset = select_by_attribute(plants, 'Utility Name', 'Brandon Shores LLC')
    subset = select_by_attribute(plants, 'Current Designation', 'Other Remaining')
"""
def select_by_attribute(df, attribute, values):
    if attribute not in df.columns:
        print('%s is not a valid attribute. Please select one of the following:\n' %attribute)
        print(*df.columns, sep=', ')
        return
    if not isinstance(values, list):
        values = [values]
    return df[df[attribute].isin(values)]


"""
Selects all plants without a retirement date, per SC planning tool.
"""
def select_remaining_plants(df):
    remaining_groups = ['2020 Vulnerable', '2025 Vulnerable', 'Announced', 'Other Remaining']
    return df[df['Current Designation'].isin(remaining_groups)]

"""
Selects all plants with plant balances available.
"""
def select_plants_with_balances(df):
    return df[df['Plant Balance'].notnull()]


"""
Aggregates data by level (plant, utility, geography).
TO IMPLEMENT: more interesting statstics, i.e. % of coal for all generation in a state

INPUTS
    df:    database
    level: level by which to aggregate, i.e. Unit, Plant, State
OUTPUTS: 
    df:    database grouped by level
RUNTIME: ~1-10 sec
EXAMPLE USAGE:
    agg = aggregate_by_level(plants, 'Plant')
"""
def aggregate_by_level(df, level):
    col_to_keep = ['Plant Balance', 'Nameplate Capacity (MW)']
    if level == 'Unit/Subunit':
        return df
    if level == 'Plant':
        df = df.groupby(['Plant Code', 'Plant Name', 'Utility Name', 'Technology']).sum()
        return df[col_to_keep].reset_index().set_index('Plant Code')
    if level == 'Technology':
        df = df.groupby('Technology').sum()
        return df[col_to_keep].reset_index().set_index('Technology')
    if level == 'Utility':
        df = df.groupby('Utility Name').sum()
        return df[col_to_keep].reset_index().set_index('Utility Name')
    if level == 'State':
        df = df.groupby('State').sum()
        return df[col_to_keep].reset_index().set_index('State')
        
"""
Sorts data by an attribute (column).
"""
def sort_by_attribute(df, attribute, ascending=True):
    try:
        sorted = df.sort_values(attribute, ascending=ascending)
    except(KeyError):
        print('%s is not a valid attribute. Please select one of the following:\n' %attribute)
        print(*df.columns, sep=', ')
        return
    return sorted

        

        