import pandas as pd
import pyodbc
import folium
import textwrap as tw
import os
os.chdir('C:/Users/obriene/Projects/Mapping')

#Read in patient postcode data, remove double spaces from postcode.
df = pd.read_excel('Patients for CKD heat map.xlsx', 
                   usecols=['Hospital Number', 'Forename', 'Surname',
                            'Patient Postcode', 'Status', 'CKD Stage', 'eGFR'],
                   skiprows=2)
df['Patient Postcode'] = df['Patient Postcode'].replace(r'\s+', ' ', regex=True)
df['Name'] = df['Forename'] + ' ' + df['Surname'] + ', '
#list of postcodes and their lat long
pcode_LL = pd.read_csv('ukpostcodes.csv').rename(columns={
                                                 'postcode':'FullPostCode'})

#travel times
travel_times = pd.read_csv(r"G:/PerfInfo/Performance Management/PIT Adhocs/"\
						   "2021-2022/Hannah/PeripheralClinic/"\
							   "travel_times_Full_Clinics.csv",\
 						   usecols = ['pcode', 'Tavistock Hospital']).rename(
                                           columns={'pcode':'Patient Postcode'})
#tavistock latlong
tavistock = pcode_LL.loc[pcode_LL['FullPostCode'] == 'PL19 8LD']

#Merge postcodes onto latlong lookup
df = df.merge(pcode_LL, left_on='Patient Postcode', right_on='FullPostCode',
              how='left')

#get count by postcode
pcode_counts = (df.groupby(['Patient Postcode', 'latitude', 'longitude'],
                          as_index=False).agg({'Hospital Number':'count',
                                               'Name':'sum'})
                  .rename(columns={'Hospital Number':'Number of Patients'}))
pcode_counts = pcode_counts.merge(travel_times, on='Patient Postcode',
                                  how='left')

#heatmap
from folium.plugins import HeatMap
m = folium.Map(location = [50.4163,-4.11849],
               zoom_start=10,
               tiles="cartodbpositron")
#Repeat by number of patients in each pcode
heat_df = pcode_counts.loc[pcode_counts.index.repeat(
                                            pcode_counts['Number of Patients']),
                            ['latitude','longitude']].dropna()
#Make a list of values
heat_data = [[row['latitude'], row['longitude']] for index,
             row in heat_df.iterrows()]
#create heatmap
HeatMap(heat_data).add_to(m)
#Add opaque circle markers to add tooltips of distance to tavistock
for i, row in pcode_counts.iterrows():
    folium.CircleMarker([row['latitude'], row['longitude']], radius=20,
                        fill_opacity=0, fill=True, color=None,
                        tooltip=f'Time to Tavistock Hospital: {row['Tavistock Hospital']} mins, {row['Number of Patients']} Patients: {row['Name'][:-2]}',
                        icon=None).add_to(m)
#Add tavistock hospital as a marker
folium.Marker([tavistock['latitude'].iloc[0], tavistock['longitude'].iloc[0]],
			  tooltip='Tavistock Hospital',
			  icon=folium.Icon(color='darkpurple', icon="h-square",
					  prefix = 'fa')).add_to(m)
title = "The information in this report contains Personal Confidential Data (PCD) and must not be sent to another organisation or non-NHS.NET email address without IG consent"
title_html = f'<h1 style="position:absolute;z-index:100000;left:40vw;color:red;font-size:160%;" >{title}</h1>'
m.get_root().html.add_child(folium.Element(title_html))
#save
m.save(r"Outputs/nephrology_peripheral_clinic.html")
