import pandas as pd
import pyodbc
import folium
import textwrap as tw
import os
os.chdir('C:/Users/obriene/Projects/test/Mapping')

#Read in patient postcode data
df = pd.read_excel('Copy of 26205 GR Obs Scan Postcodes.xlsx', sheet_name='DATA',
                   usecols=['pasid', 'Scan Appt', 'clinic_code', 'FullPostCode', 'PostCode_Area'],
                   skiprows=8)
df.loc[df['FullPostCode'] == 'TQ9  7RU']
df['FullPostCode'] = df['FullPostCode'].replace(r'\s+', ' ', regex=True) # double spaces in some postcodes need to be delt with

#Postcode lookup table for lat long
#2021
#pcode_LL = pd.read_csv("G:/PerfInfo/Performance Management/PIT Adhocs/2021-2022/Hannah/Maps/pcode_LSOA_latlong.csv",
 #                           usecols = ['pcds', 'lat', 'long']).rename(columns={'pcds':'FullPostCode'})
#2023 taken from https://www.freemaptools.com/download-uk-postcode-lat-lng.htm
pcode_LL = pd.read_csv('ukpostcodes.csv').rename(columns={'postcode':'FullPostCode'})
#Merge postcodes onto latlong lookup
df = df.merge(pcode_LL, on='FullPostCode', how='left')


#Grouped by postcode
df_group = df.groupby(['FullPostCode', 'latitude', 'longitude'], as_index=False)[['pasid', 'clinic_code']].agg({'pasid':[lambda x: list(x), 'nunique'], 'clinic_code': lambda y: list(y)})
df_group.columns = ['FullPostCode', 'lat', 'long', 'pasid', 'no_patients', 'clinic_code']
#produce the map
m = folium.Map(location = [50.4163,-4.11849],
               zoom_start=10,
               tiles="cartodbpositron")

for row in df_group.values.tolist():
   p_code, lat, long, pasid, no_pat, clinic_code = row
   folium.Circle(
      location=[lat, long],
      popup=p_code,
      tooltip = 'Number of Patients: ' + str(no_pat)+\
						 ', Patient ids: ' + str(pasid).replace('[','').replace(']','').replace("'", '')+\
							 ', clinic code: '+ str(set(clinic_code)).replace('{','').replace('}','').replace("'", ''),
      radius=float(no_pat)*40,
      color='lightseagreen',
      fill=True,
      fill_color='lightseagreen'
   ).add_to(m)

m.save(r"C:/Users/obriene/Projects/test/Mapping/GR Obs Scan Map.html")


#Grouped by postcode area
df_area = df.groupby('PostCode_Area', as_index=False).agg({'latitude':['mean'], 'longitude':['mean'], 'pasid':['nunique']})
n = folium.Map(location = [50.4163,-4.11849],
               zoom_start=10,
               tiles="cartodbpositron")
#produce the map
for row in df_area.values.tolist():
   p_code, lat, long, no_pat = row
   folium.Circle(
      location=[lat, long],
      popup=p_code,
      tooltip = 'Number of Patients: ' + str(no_pat),
      radius=float(no_pat),
      color='lightseagreen',
      fill=True,
      fill_color='lightseagreen'
   ).add_to(n)

n.save(r"C:/Users/obriene/Projects/test/Mapping/GR Obs Scan Area Map.html")