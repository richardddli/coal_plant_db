# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""

import os
import pandas as pd

## This script generates a prototype database with the following attributes
## for each coal plant:
## (1) Ownership (how many owners, what types of owners [TODO])
## (2) Remaining plant balance (taken directly from Depreciation model)
## (3) Self committing vs Market committing (taken from JD Sierra Club paper TODO)

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

# import depreciation model outputs (NEEDS TO BE UPDATED)
balance_or = pd.read_csv(os.path.join(cwd, 'plant balance [wrong].csv'), skiprows=1)
balance = balance_or[balance_or['EIA Plant ID Best Match'].notnull()]
balance = balance[balance['Account Title'].notnull()]

# matching plant and unit between depreciation model and EIA 860
plants['Generator ID'] = plants['Generator ID'].astype(str)
for i, row in balance.iterrows():
    if row['EIA Plant ID Best Match'] in set(plants['Plant Code']):
        matches = plants[plants['Plant Code'] == row['EIA Plant ID Best Match']]  
        if row['Account Title'][-1].isdigit():
            matches = matches[matches['Generator ID'] == row['Account Title'][-1]]
            if(not matches.empty):
                plants.loc[matches.index[0], 'Plant Balance'] = row['ANNUAL AMOUNT']
        else:
            plants.loc[matches.index[0], 'Plant Balance'] = row['ANNUAL AMOUNT']
            
# selecting only coal plants with a plant balance for now
plants2 = plants[plants['Plant Balance'].notnull()]
plants2 = plants2[plants2['Technology'] == 'Conventional Steam Coal']

# adding multiple ownership information
for i, row in plants2[plants2['Ownership'] != 'S'].iterrows():
    own = owners[(owners['Plant Code'] == row['Plant Code']) & 
                 (owners['Generator ID'] == row['Generator ID'])]
    own.reset_index(inplace=True)
    for j, r in own.iterrows():
        plants2.loc[i, 'Owner #' + str(j)] = r['Owner Name']
        plants2.loc[i, '% Owned #' + str(j)] = r['Percent Owned']

#plants2.to_csv('/Users/richardli/Documents/Coal Database/prototype output.csv')




