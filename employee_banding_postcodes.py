import os
import folium
import xlsxwriter
import numpy as np
import pandas as pd
import geopandas as gpd
from pathlib import Path
from sqlalchemy import create_engine
from folium.plugins import HeatMap
from datetime import datetime
os.chdir('C:/Users/obriene/Projects/Mapping/Outputs')
run_date = datetime.today().strftime('%Y-%m-%d')

#readin postcode to latlong data
pcode_LL = pd.read_csv("G:/PerfInfo/Performance Management/PIT Adhocs/2021-2022/Hannah/Maps/pcode_LSOA_latlong.csv",
                            usecols = ['pcds', 'lat', 'long'])

#Read in employee data
cl3_engine = create_engine('mssql+pyodbc://@cl3-data/DataWarehouse?'\
                           'trusted_connection=yes&driver=ODBC+Driver+17'\
                               '+for+SQL+Server')
Band_pcds_query = """SELECT *
FROM [HumanResources].[vw_CurrentStaffPostcodes]
WHERE Banding LIKE '%Band%'
AND PostCode IS NOT NULL
"""
Band_pcds = pd.read_sql(Band_pcds_query, cl3_engine).rename(columns={'PostCode':'pcds'})
cl3_engine.dispose()

#Exttract band numbers as group up
Band_pcds['Band No'] = Band_pcds['Banding'].str.extract(r'(\d+)').astype(int)
Band_pcds['Band Groups'] = np.select([((Band_pcds['Band No'] >= 1) & (Band_pcds['Band No'] <= 3)),
                                      ((Band_pcds['Band No'] >= 4) & (Band_pcds['Band No'] <= 5)),
                                      ((Band_pcds['Band No'] >= 6) & (Band_pcds['Band No'] <= 7)),
                                      (Band_pcds['Band No'] >= 8)],
                                      ['Bands 1-3', 'Bands 4-5', 'Bands 6-7', 'Bands 8+'])

#Add in space and capitalise postcodes
Band_pcds['pcds'] = [pcd.upper()  if ' ' in pcd else (pcd[:-3]+' '+pcd[-3:]).upper() for pcd in Band_pcds['pcds'].tolist()]

#Merge data together, group up data to postcode up to last 2 charecters to keep annonymity
LL_df = pcode_LL.merge(Band_pcds, on='pcds', how='inner')
LL_df['Area'] = LL_df['pcds'].str[:-2]
LL_df = LL_df.groupby(['Band Groups', 'Area'], as_index=False).agg({'lat':'mean', 'long':'mean', 'Band No':'count'})


# =============================================================================
# #HEATMAP
# =============================================================================
for group in LL_df['Band Groups'].drop_duplicates().tolist():
     m = folium.Map(location = [50.4163,-4.11849],
                    zoom_start=10,
                    min_zoom = 7,
                    tiles='cartodbpositron')
     heat_df = LL_df.loc[LL_df['Band Groups'] == group].copy()
     heat_df = heat_df.loc[heat_df.index.repeat(heat_df['Band No']), ['lat','long']].dropna()#Repeat by number of patients in each pcode
     #Make a list of values
     heat_data = [[row['lat'],row['long']] for index, row in heat_df.iterrows()]
     HeatMap(heat_data).add_to(m)
     m.save(run_date + ' ' + group + ' heatmap.html')

# =============================================================================
# #Table Data
# =============================================================================
#read in pension data
pension = pd.read_excel('C:/Users/obriene/Projects/Mapping/Pension Opt Out Month 2 2024.xlsx')
pension = pension.rename(columns={'Band ':'Banding', 'Age Band':'AgeBand'})

#Functions
def parse_date(td):
    #Conerts difference between two dates into string of xY xm
    resYear = ((td.dt.days)/364.0).astype(float)                 # get the number of years including the the numbers after the dot
    resMonth = ((resYear - resYear.astype(int))*364/30).astype(int).astype(str)  # get the number of months, by multiply the number after the dot by 364 and divide by 30.
    resYear = (resYear).astype(int).astype(str)
    return resYear + "Y " + resMonth + "m"

def group_by_band(df, pen, cols):
     #Function to group up data by different values and produce the results.
     df = df.groupby(cols).agg({'Band No':'count', 'FTE':'sum', 'Days in Position':['mean', 'max']}).reset_index()
     df.columns = cols + ['Headcount', 'FTE', 'Average Time in Position', 'Max Time in Position']
     df['Average Time in Position'] = parse_date(df['Average Time in Position'])
     df['Max Time in Position'] = parse_date(df['Max Time in Position'])
     #Add pension data
     df = df.merge(pen.groupby(cols, as_index=False)['Pension Opt Out'].count(), on=cols, how='left')
     return df

#Select/add required columns
df = Band_pcds[['Banding', 'Band No', 'Band Groups', 'FTE', 'AgeBand', 'StartDateInPosition']].copy()
df['Days in Position'] = (pd.Timestamp.now() - pd.to_datetime(df['StartDateInPosition']))
df['FTE'] = df['FTE'].astype(float)

#Get the grouped tables required
band_name = group_by_band(df, pension, ['Banding'])
age_band = group_by_band(df, pension, ['Banding', 'AgeBand'])
lookup_col = age_band[['Banding', 'AgeBand']].astype(str).agg(' '.join, axis=1)
age_band.insert(loc=0, column='lookup', value=lookup_col)


##############To excel ##################
#Lists for formatting
ages = ['<=20 Years', '21-25', '26-30', '31-35', '36-40', '41-45', '46-50', '51-55', '56-60', '61-65', '66-70', '>=71 Years']
cols = ['B', 'C', 'D', 'E', 'F']
col_names = band_name.columns[1:].tolist()

#Excel Writer
writer = pd.ExcelWriter(f"C:/Users/obriene/Projects/Mapping/Outputs/{run_date}_aggregate_employee_band_data.xlsx", engine='xlsxwriter')
workbook = writer.book

####COVER SHEET####
#Add formats
white = workbook.add_format({'bg_color':'white'})
yellow = workbook.add_format({'align':'center', 'border':True, 'bg_color':'yellow'})
center = workbook.add_format({'align':'center'})
center_border = workbook.add_format({'align':'center', 'border':True})
bold_right = workbook.add_format({'bold':True, 'align':'right', 'border':2})
bold_wrap = workbook.add_format({'bold': True, 'align':'center', 'valign':'center', 'text_wrap':True, 'border':2})

#Add filter worksheet
worksheet = workbook.add_worksheet('Filter')

#Set general column formats
worksheet.set_column(0, 27, 8, white)
worksheet.set_column(0, 0, 21, white)
worksheet.set_column(1, 4, 10, white)

#Add band drop down cell
worksheet.write('A2', 'Band:', bold_right)
worksheet.write('B2', '', yellow)
worksheet.data_validation('B2', {'validate':'list',
                                 'source':band_name['Banding'].drop_duplicates().sort_values().tolist()})

#Add in text for band level lookups
for i, age in enumerate(ages):
     worksheet.write(f'A{i+11}', age, bold_right)
#Add band level vlookups
worksheet.write_formula('B4', "=IFERROR(VLOOKUP(B2,'Band Data'!A:F,2,0), 0)", center_border)
worksheet.write_formula('B5', "=IFERROR(VLOOKUP(B2,'Band Data'!A:F,3,0), 0)", center_border)
worksheet.write_formula('B6', '''=IFERROR(VLOOKUP(B2,'Band Data'!A:F,4,0), "-")''', center_border)
worksheet.write_formula('B7', '''=IFERROR(VLOOKUP(B2,'Band Data'!A:F,5,0), "-")''', center_border)
worksheet.write_formula('B8', '''=IFERROR(VLOOKUP(B2,'Band Data'!A:F,6,0), "-")''', center_border)

#Add text for age band level lookups
for i, col in enumerate(zip(cols, col_names)):
     worksheet.write(f'A{i+4}', col[1], bold_right)
     worksheet.write(f'{col[0]}10', col[1], bold_wrap)
#Add age band lookups
for i in range(11,23):
     #Headcount
     worksheet.write_formula(f'B{i}',f'''=IFERROR(VLOOKUP((B2&" "&A{i}),'Age Band Data'!A:H,4,0), 0)''', center_border)
     #FTE
     worksheet.write_formula(f'C{i}', f'''=IFERROR(VLOOKUP((B2&" "&A{i}),'Age Band Data'!A:H,5,0), 0)''', center_border)
     #Average Time in Position
     worksheet.write_formula(f'D{i}', f'''=IFERROR(VLOOKUP((B2&" "&A{i}),'Age Band Data'!A:H,6,0), "-")''', center_border)
     #Longest Time in Position
     worksheet.write_formula(f'E{i}', f'''=IFERROR(VLOOKUP((B2&" "&A{i}),'Age Band Data'!A:H,7,0), "-")''', center_border)
     #Pension Opt Out
     worksheet.write_formula(f'F{i}', f'''=IFERROR(VLOOKUP((B2&" "&A{i}),'Age Band Data'!A:H,8,0), "-")''', center_border)

#Add in bar charts
headcount_chart = workbook.add_chart({'type':'column'})
headcount_chart.add_series({'name':'Filter!$B$10',
                            'categories': 'Filter!$A$11:$A$22',
                            'values': 'Filter!$B$11:$B$22'})
headcount_chart.add_series({'name':'Filter!$C$10',
                            'categories': 'Filter!$A$11:$A$22',
                            'values': 'Filter!$C$11:$C$22'})
headcount_chart.set_title({'name':'Headcount and FTE'})
headcount_chart.set_x_axis({'name':'Age Bands', 'num_font':{'rotation':45}})
headcount_chart.set_style(2)
worksheet.insert_chart('H2', headcount_chart, {'x_scale':1.04, 'y_scale':1.07})

pension_chart = workbook.add_chart({'type':'column'})
pension_chart.add_series({'name':'Filter!$F$10',
                            'categories': 'Filter!$A$11:$A$22',
                            'values': 'Filter!$F$11:$F$22'})
pension_chart.set_title({'name':'Pension Opt Out'})
pension_chart.set_x_axis({'name':'Age Bands', 'num_font':{'rotation':45}})
pension_chart.set_style(2)
pension_chart.set_legend({'none':True})
worksheet.insert_chart('H17', pension_chart, {'x_scale':1.04, 'y_scale':1.07})

####Add lookup sheets####
band_name.to_excel(writer, sheet_name='Band Data', index=False, engine='xlsxwriter')
band_worksheet = writer.sheets['Band Data']
band_worksheet.set_column(0, 0, 16)
band_worksheet.set_column(1, 4, 11, center)

age_band.to_excel(writer, sheet_name='Age Band Data', index=False, engine='xlsxwriter')
age_worksheet = writer.sheets['Age Band Data']
age_worksheet.set_column(0, 1, 26)
age_worksheet.set_column(2, 6, 11, center)

####save and close the workbook####
writer.close()
