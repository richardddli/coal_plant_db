# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""

import os
import pandas as pd

## This script aggregates the total MW of coal generation for each utility,
## broken down by partial ownership given the tables in EIA 860


# import plant & ownership data
cwd = os.path.dirname(os.path.realpath(__file__))
plant_cols = ['Plant Name', 'Plant Code', 'Utility Name', 'Utility ID', 'State',
              'Generator ID', 'Technology', 'Prime Mover', 'Unit Code',
              'Ownership', 'Nameplate Capacity (MW)', 'Nameplate Power Factor',
              'Operating Month', 'Operating Year', 'Planned Retirement Month', 
              'Planned Retirement Year', 'Sector Name']
plants = pd.read_csv(os.path.join(cwd, 'eia860', '3_1_Generator_Y2018_Early_Release.csv'), skiprows=2, usecols=plant_cols)
owners = pd.read_csv(os.path.join(cwd, 'eia860', '4___Owner_Y2018_Early_Release.csv'), skiprows=2)

# select coal plants in electric sector
plants = plants[(plants['Technology'] == 'Conventional Steam Coal') &
                (plants['Sector Name'] == 'Electric Utility')]
# note for sam: another way to select data in pandas:
# plants.query('Technology == "Conventional Steam Coal" | Sector Name == "Electric Utility"')
# but query does not allow for column names with spaces so I'm not a fan...

# aggregate MW for singly owned plants
plants_S = plants[plants['Ownership'] == 'S']
total_MW = plants_S.groupby('Utility Name').sum()['Nameplate Capacity (MW)']

# add in jointly owned plants
total = {}
# cast Percent Owned column as floats
owners.loc[owners['Percent Owned'] == ' ', 'Percent Owned'] = '0'
owners['Percent Owned'] = owners['Percent Owned'].astype('float')

plants_J = plants[plants['Ownership'] != 'S']
for i, plant in plants_J.iterrows():
    owner_list = owners[(owners['Plant Code'] == plant['Plant Code']) & 
                        (owners['Generator ID'] == plant['Generator ID'])]
    for j, owner in owner_list.iterrows():
        if owner['Owner Name'] in total:
            total[owner['Owner Name']] += owner['Percent Owned'] * plant['Nameplate Capacity (MW)']
        else:
            total[owner['Owner Name']] = owner['Percent Owned'] * plant['Nameplate Capacity (MW)']

total_MW_joint = pd.Series(total)
total_MW_joint.sort_index(inplace=True)

#now concatenate two tables
total_MW = total_MW.append(total_MW_joint)
total_MW.groupby(total_MW.index).sum()

#print table
total_MW.to_csv(os.path.join(cwd, 'eia860', 'Total Coal By Owner.csv'))