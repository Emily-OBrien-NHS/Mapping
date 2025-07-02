import pandas as pd
import pyodbc
import folium
from folium.plugins import HeatMap
import textwrap as tw
import os
import datetime
os.chdir('C:/Users/obriene/Projects/Mapping/Nephrology Mapping')
today = datetime.date.today().strftime('%d-%m-%Y')

################################################################################
################################Read in Data####################################
################################################################################
#########list of postcodes and their lat long
pcode_LL = pd.read_csv('C:/Users/obriene/Projects/Mapping/ukpostcodes.csv'
                       ).rename(columns={'postcode':'FullPostCode'})

#########All Patients
all_pats = pd.read_excel('Vital Data Patient Listing (P2489).xlsx',
                         usecols=['Hospital Number', 'Forename', 'Surname',
                                  'Patient Postcode', 'Status', 'CKD Stage',
                                  'eGFR'], skiprows=1)
#Replace double spaces in postcode
all_pats['Patient Postcode'] = all_pats['Patient Postcode'].replace(r'\s+', ' ', regex=True)
all_pats['Name'] = all_pats['Forename'] + ' ' + all_pats['Surname'] + ', '
#Merge postcodes onto latlong lookup
all_pats = all_pats.merge(pcode_LL, left_on='Patient Postcode',
                          right_on='FullPostCode', how='left')

##########CKD Patients with GFR<15
CKD_lt_15 = pd.read_excel('VitalData CKD Stage 4 and 5 Active Patients (P2407).xlsx',
                          usecols=['Hospital Number', 'Forename', 'Surname',
                                   'Status', 'CKD Stage', 'eGFR Result',
                                   'Choice RRT'], skiprows=1)
#create name column
CKD_lt_15['Name'] = CKD_lt_15['Forename'] + ' ' + CKD_lt_15['Surname'] + ', '
#Filter to current patients with GFR < 15
CKD_lt_15 = CKD_lt_15.loc[(CKD_lt_15['Choice RRT'] != 'Discharged')
                          & (CKD_lt_15['eGFR Result'] < 15)].copy()
#Get postcodes
CKD_lt_15 = CKD_lt_15.merge(
                      all_pats[['Hospital Number', 'Patient Postcode', 'latitude', 'longitude']].copy(),
                      on='Hospital Number', how='left')

#########HD (hosps) patients
HD_hosp = all_pats.loc[all_pats['Status'] == 'HD (Hosp)'].copy()

#########HD (home) patients
HD_home = all_pats.loc[all_pats['Status'] == 'HD (Home)'].copy()

#########travel times
travel_times = pd.read_csv(r"G:/PerfInfo/Performance Management/PIT Adhocs/"\
						   "2021-2022/Hannah/PeripheralClinic/"\
							   "travel_times_Full_Clinics.csv",\
 						   usecols = ['pcode', 'Tavistock Hospital']).rename(
                                           columns={'pcode':'Patient Postcode'})
#tavistock latlong
tavistock = pcode_LL.loc[pcode_LL['FullPostCode'] == 'PL19 8LD']

#Get patient counts per postcode
def pcode_counts(df):
    #get count by postcode
    pcode_counts = (df.groupby(['Patient Postcode', 'latitude', 'longitude'],
                            as_index=False).agg({'Hospital Number':'count',
                                                'Name':'sum'})
                    .rename(columns={'Hospital Number':'Number of Patients'}))
    pcode_counts = pcode_counts.merge(travel_times, on='Patient Postcode',
                                    how='left')
    
    return pcode_counts

all_pats = pcode_counts(all_pats)
CKD_lt_15 = pcode_counts(CKD_lt_15)
HD_hosp = pcode_counts(HD_hosp)
HD_home = pcode_counts(HD_home)

maps_dict = {'All Patients':all_pats,
             'CKD Patients with GFR under 15':CKD_lt_15,
             'HD (hosp) Patients':HD_hosp,
             'HD (home) Patients':HD_home}
################################################################################
##################################Heatmaps######################################
################################################################################
for name, df in maps_dict.items():
    #heatmap
    m = folium.Map(location = [50.4163,-4.11849],
                zoom_start=10,
                tiles="cartodbpositron")
    #Repeat by number of patients in each pcode
    heat_df = df.loc[df.index.repeat(df['Number of Patients']),
                     ['latitude','longitude']].dropna()
    #Make a list of values
    heat_data = [[row['latitude'], row['longitude']] for index,
                row in heat_df.iterrows()]
    #create heatmap
    HeatMap(heat_data).add_to(m)
    #Add opaque circle markers to add tooltips of distance to tavistock
#    for i, row in df.iterrows():
#        folium.CircleMarker([row['latitude'], row['longitude']], radius=20,
#                            fill_opacity=0, fill=True, color=None,
#                            tooltip=f'Time to Tavistock Hospital: {row['Tavistock Hospital']} mins, {row['Number of Patients']} Patients: {row['Name'][:-2]}',
#                            icon=None).add_to(m)
    #Add tavistock hospital as a marker
#    folium.Marker([tavistock['latitude'].iloc[0], tavistock['longitude'].iloc[0]],
#                tooltip='Tavistock Hospital',
#                icon=folium.Icon(color='darkpurple', icon="h-square",
#                        prefix = 'fa')).add_to(m)
    title = "The information in this map contains Personal Confidential Data (PCD) and must not be sent to another organisation or non-NHS.NET email address without IG consent"
    title_html = f'<h1 style="position:absolute;z-index:100000;left:40vw;color:red;font-size:160%;" >{title}</h1>'
    m.get_root().html.add_child(folium.Element(title_html))
    #save
    m.save(fr"Outputs/nephrology-{name} {today}.html")
