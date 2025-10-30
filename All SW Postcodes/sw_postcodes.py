import pandas as pd
import numpy as np
import os

#https://geoportal.statistics.gov.uk/datasets/ons::output-area-2021-to-parncp-to-lad-to-rgn-to-ctry-december-2024-best-fit-lookup-in-ew/about
districts = (pd.read_csv('Output_Area_(2021)_to_PARNCP_to_LAD_to_RGN_to_CTRY_(December_2024)_Best_Fit_Lookup_in_EW.csv')
             [['LAD24CD', 'LAD24NM', 'CTYUA24CD', 'CTYUA24NM']].drop_duplicates())
districts.columns = ['LAD24CD', 'Local Authority', 'cty', 'County']
keep_ctys = ['Cornwall', 'Devon', 'Plymouth', 'Torbay']
keep_dist = districts.loc[districts['County'].isin(keep_ctys)]


####OS Dataset
cols = ['Postcode','Positional_quality_indicator','PO_Box_indicator',
        'Total_number_of_delivery_points',
        'Delivery_points_used_to_create_the_CPLC','Domestic_delivery_points',
        'Non_domestic_delivery_points','PO_Box_delivery_points',
        'Matched_address_premises','Unmatched_delivery_points','Eastings',
        'Northings','Country_code',
        'NHS_regional_HA_code','NHS_HA_code','Admin_county_code',
        'Admin_district_code','Admin_ward_code','Postcode_type']

sw_pcds = []
for file in os.listdir('Data/CSV'):
    temp = pd.read_csv(f'Data/CSV/{file}', header=None)
    temp.columns=cols
    sw_pcds += temp.loc[temp['NHS_HA_code'] == 'E18000010'].values.tolist()

OS = pd.DataFrame(sw_pcds, columns=cols).merge(districts, left_on='Admin_district_code', right_on='LAD24CD', how='left')
OS = OS[['Postcode', 'Local Authority', 'County']].copy()
OS = OS.loc[OS['County'].isin(keep_ctys)]

#####ONS data is in download folder to validate against.
#https://www.data.gov.uk/dataset/0ed89962-3d60-4ac7-b659-46bd765f6de4/national-statistics-postcode-lookup-may-2025-for-the-uk
#https://open-geography-portalx-ons.hub.arcgis.com/datasets/ons::national-statistics-postcode-lookup-may-2025-for-the-uk/about
ONS = pd.read_csv('C:/Users/obriene/Projects/Mapping/All SW Postcodes/NSPL_MAY_2025/Data/NSPL_MAY_2025_UK.csv').rename(columns={'pcds':'Postcode'})
ONS = ONS.merge(districts, left_on='laua', right_on='LAD24CD', how='left')
ONS = ONS[['Postcode', 'Local Authority', 'County']]
ONS = ONS.loc[ONS['County'].isin(keep_ctys)]




#####All postcodes in keep_ctys from both data sets
full_pcd_list = OS.merge(ONS, on='Postcode', how='outer', suffixes=[' OS Data', ' ONS Data'])
full_pcd_list['OS and ONS match'] = True
full_pcd_list.loc[~((full_pcd_list['Local Authority OS Data'] == full_pcd_list['Local Authority ONS Data'])
                    & (full_pcd_list['County OS Data'] == full_pcd_list['County ONS Data'])),
                    'OS and ONS match'] = False
full_pcd_list.to_excel('All Plymouth, Devon, Cornwall and Torbay postcodes.xlsx', index=False)


#####Check postcodes
dfs = []
for sheet in ['Plymouth & South Hams', 'Cornwall Postcodes', 'Devon Postcodes']:
    temp = pd.read_excel('PHR Postcodes - 250424 (copy).xlsx', sheet_name=sheet).dropna(subset='Postcode')
    temp['sheet'] = sheet
    dfs.append(temp)
check_pcds = pd.concat(dfs)[['Postcode', 'Region', 'Region 2', 'sheet']]
check_pcds['Postcode join'] = check_pcds['Postcode'].str.replace('  ', ' ')


#Compare postcode data with the 2 datasets
compare = check_pcds.merge(full_pcd_list, left_on='Postcode join', right_on='Postcode', how='left', suffixes=[None, '_x'])

keep_cols = ['Postcode', 'Region', 'Region 2', 'Local Authority OS Data', 'County OS Data', 'Local Authority ONS Data', 'County ONS Data', 'OS and ONS match']
with pd.ExcelWriter('PHR Postcodes validation.xlsx') as writer:  
    compare.loc[compare['sheet'] == 'Plymouth & South Hams', keep_cols].to_excel(writer, sheet_name='Plymouth & South Hams', index=False)
    compare.loc[compare['sheet'] == 'Cornwall Postcodes', keep_cols].to_excel(writer, sheet_name='Cornwall Postcodes', index=False)
    compare.loc[compare['sheet'] == 'Devon Postcodes', keep_cols].to_excel(writer, sheet_name='Devon Postcodes', index=False)
