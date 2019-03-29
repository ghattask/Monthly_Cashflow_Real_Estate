import numpy as np
import pandas as pd
import os
import glob
import calendar
import datetime
import re
#from tkinter import filedialog
#from tkinter import Tk




"""root = Tk()
root.withdraw()
cwd = os.getcwd()
tempdir = filedialog.askdirectory(parent= root, initialdir=cwd, title = 'Please Select a Directory')
if len(tempdir) > 0:
   print ('You Chose: {}'.format(cwd))
"""
cwd = os.getcwd()
print('The Current Path to the files is: {}'.format(cwd))
os.chdir('/Users/karimghattas/desktop/La_Jeunesse')
print(cwd)




# Rename all the file in the directory with the following rules:
   # 1. Make the file names lower case
   # 2. Rename files from camelcase to underscore notation
   # 3. Please note that the order is important so that different rename functions do not affect each other.
i = 0 
for f in os.listdir(cwd):
    [os.rename(os.path.join(cwd, f), os.path.join(cwd, f).replace(' - ', '_').lower())]
    [os.rename(os.path.join(cwd, f), os.path.join(cwd, f).replace(' ', '_').lower())]
    [os.rename(os.path.join(cwd, f), os.path.join(cwd, f).replace('-', '_').lower())]
    i += 1 
print('Files Successfully Renamed')

list_of_rentroll_files = glob.glob('[0-9][0-9]*rentroll*[0-9][0-9][0-9][0-9].xlsx') #Declare variable with filenames that include a pattern in the selected Directory.
latest_rentroll_file = sorted(list_of_rentroll_files,key=os.path.getctime)[-1] #Determine the latest file in the directory using the ctime function.
list_of_banque_files = glob.glob('[0-9][0-9]*banque*[0-9][0-9][0-9][0-9].xlsx') #Declare variable with filenames that include a pattern in the selected Directory.
latest_banque_file = sorted(list_of_banque_files,key=os.path.getctime)[-1] #Determine the latest file in the directory using the ctime function.

# Check that the month and year for both files are the same.
if latest_rentroll_file[:2] == latest_banque_file[:2] and latest_rentroll_file[-9:-5] == latest_banque_file[-9:-5]: #Condition to make sure the year of the Rentroll Excel file equals the year of the Banque Excvel file.
   month_and_year_date = str("".join(re.findall(r'\d+', str(latest_rentroll_file)))) #Extract the month and year of the latest file in the Directory from a list into concatnated variable.
   start_date = datetime.datetime.strptime(month_and_year_date, '%m%Y').date() #Join the month and year into a date variable from the list defined above.
   end_date = datetime.date(start_date.year, start_date.month, calendar.monthrange(start_date.year, start_date.month)[-1]) #Calculates the last day of the month for any given month. This provides the date range of the analysis.
   print('Both excel files found. The files are: {} and {} for dates between {} and {}.'.format(latest_rentroll_file,latest_banque_file,start_date, end_date))
elif int(latest_rentroll_file[:2]) < int(latest_banque_file[:2]) and int(latest_rentroll_file[-9:-5]) == int(latest_banque_file[-9:-5]):
   print('Rentroll file not found; Please check the current path and download the most recent rentroll file.')
elif int(latest_rentroll_file[:2]) > int(latest_banque_file[:2]) and int(latest_rentroll_file[-9:-5]) == int(latest_banque_file[-9:-5]):
   print('Banque file not found; Please check the current path and download the most recent banque file.')
else: 
   print('Files not found; Please check the current path and that the files were downloaded.')
file_list = (latest_rentroll_file, latest_banque_file) #Define a tuple with both filenames




if 'df' in locals(): # delete variable if it exists
    del df
df_this_month = pd.read_excel(latest_rentroll_file) #upload file into dataframe

#Cleaning Dataframe Columns/Rows
df_this_month = df_this_month.drop(df_this_month.index[66], axis = 0) # remove specific row where there is an ellipsis
df_this_month = df_this_month.dropna(how='all') #drop all rows that are empty.
new_header = df_this_month.iloc[0] #define the headers as the first row
df_this_month = df_this_month[1:] #Remove the original header row
df_this_month.columns = new_header #define the new header row
df_this_month.columns.values[1] = 'Bed_Bath/Sq_Footage' #Rename 2nd row to 'Bed_bath/Sq_Footage'
df_this_month.columns.values[3] = 'Lease_Start_Date' #rename 4th row to 'Lease_Start_Date'
df_this_month.rename(columns={'Lease Expiry\nDate':'Lease_Expiry_Date'}, inplace=True) #rename 'Lease Expiry\nDate' column. Different approach.
df_this_month.columns = df_this_month.columns.str.replace(' ','_') #replace whitespace with underscores for column names.
rows_to_be_removed = ('COMMERCE','Residential Sub-total','Sous-Total Autres','Sous-Total Commercial','TOTAL') #List defining values in TENANT column.
df_this_month = df_this_month[~df_this_month['TENANT'].isin(rows_to_be_removed)] #remove rows that are included in the rows_to_be_removed list
df_this_month = df_this_month.reset_index(drop=True) #rest index
#Populate 'Lease_Start_Date' if 'Lease_Expiry_Date' is not NaN. Assume that lease was signed 1 year before its expiry date.
df_this_month['Lease_Start_Date'].replace(np.nan, df_this_month['Lease_Expiry_Date'] - datetime.timedelta(days=364), inplace = True)
df_this_month['Lease_Expiry_Date'] = df_this_month['Lease_Expiry_Date'].dt.date #convert from Datetime to date
df_this_month['Lease_Start_Date'] = df_this_month['Lease_Start_Date'].dt.date #convert from Datetime to date
#clean columns to remove NaN and replace with 0
columns_to_be_cleaned = (df_this_month['Collected_By_Checks'], df_this_month['Collected_By_Cash'], df_this_month['Vacant'], df_this_month['Autre_et_Bad_Debts'], df_this_month['Solde_ouverture'], df_this_month['RENT']) #define a list of columns that will be cleaned
i = 0
for columns in columns_to_be_cleaned:
    columns.replace(np.nan,0,inplace=True)
    i+= 1
#clean Solde Ouverture column where numbers are less than one cent.
df_this_month['Solde_ouverture']= df_this_month['Solde_ouverture'].mask(np.logical_and(df_this_month['Solde_ouverture'] >=-0.01, df_this_month['Solde_ouverture'] <=0.01),0)
df_this_month = df_this_month.round(2) #Round all columns in dataframe to two decimal places
df_this_month.insert(df_this_month.columns.get_loc('RENT') + 1,'Paid', df_this_month['Collected_By_Checks'] + df_this_month['Collected_By_Cash']) #insert Paid column after rent column.
df_this_month = df_this_month.drop(['Collected_By_Checks','Collected_By_Cash','COMMENTS'], axis=1) #drop collected_by_checks and collected_by_cash columns
df_this_month.insert(0,'Month_Ending', end_date ) #adding data flag for timeseries. end_date defined earlier
df_this_month.insert(1,'Unit_Type', np.where((df_this_month['#_APT.']<=17) | (df_this_month['TENANT']=='LAVEUSE/SÉCHEUSE'), 'Residential', 'Commercial'))
with pd.option_context('display.max_rows', None, 'display.max_columns', None):
    display(df_this_month)




#Data Manipulation 1
filename = 'lajeunesse_'+str(end_date.month)+str(end_date.year)+'TD'+'.xlsx'
last_month_filename = 'lajeunesse_'+str((end_date.month)-1)+str(end_date.year)+'TD'+'.xlsx'
columns = pd.date_range(end_date, periods=1, freq='m') #declare columns to be used for base table or current month table
index = ['Total_paid','Commercial_paid','Residential__paid','Total_percent_paid','Commercial_percent_paid','Residential_percent_paid','Total_vacant','Commercial_vacant','Residential_vacant','Total_percent_vacant','Commercial_percent_vacant','Residential_percent_vacant']
df_tm_this_month = pd.DataFrame(index=index,columns=columns) # create a dataframe for this month's data
if end_date.month == 1:
    df_this_month['Solde_Fin_de_mois'] = df_this_month['RENT'] - df_this_month['Paid'] - df_this_month['Vacant'] #calculate'solde fin de mois' column for January
    df_this_month['Solde_ouverture'] = 0 #set 'solde_ouverture' column to zero for January
elif end_date.month != 1 and os.path.isfile(last_month_filename): #import last month's cumulative data
    df_base = pd.read_excel(last_month_filename,sheetname='base') #import base table with historical raw data
    df_tm_base = pd.read_excel(last_month_filename,sheetname='tm_base') #import base table with historical raw data
    if end_date.month == 12:
        df_this_month['Autre_et_Bad_Debts'] = df_this_month['Solde_Fin_de_mois'] - df_this_month['Paid'] + df_this_month['RENT'] + df_this_month['Solde_ouverture']#For the end of year (Year defined as 1st of January to 31st of December), any money owned is considered bad debits for tax reasons.
        df_this_month['Solde_Fin_de_mois'] = 0
    else:
        df_this_month['Solde_Fin_de_mois'] = df_this_month['Solde_ouverture'] + df_this_month['RENT'] - df_this_month['Paid'] - df_this_month['Vacant']
    df = pd.concat([df_this_month,df_base], ignore_index=True) #dataframe for all the raw data
else:
    print ('''Try Again - Cannot find last month's file''' )





#Data Manipulation 2
#calculations for this month's dataframe
total_rent = df_base['RENT'].sum()
paid = df_base.groupby('Unit_Type')['Paid'].sum()
percent_paid = round(paid/Total_rent,2)*100
vacant= df_base.groupby('Unit_Type')['Vacant'].sum()
percent_vacant = round(vacant/Total_rent,2)*100

#populating this month's dataframe
df_tm_this_month.iloc[0,0] = 'CA$ '+str(total_rent)
df_tm_this_month.iloc[1,0] = 'CA$ '+str(df_base['Paid'].sum())
df_tm_this_month.iloc[2,0] = 'CA$ '+str(paid[0])
df_tm_this_month.iloc[3,0] = 'CA$ '+str(paid[1])
df_tm_this_month.iloc[4,0] = str(round(df_base['Paid'].sum()/total_rent*100,2))+'%'
df_tm_this_month.iloc[5,0] = str(round(percent_paid[0],2))+'%'
df_tm_this_month.iloc[6,0] = str(round(percent_paid[1],2))+'%'
df_tm_this_month.iloc[7,0] = 'CA$ '+str(df_base['Vacant'].sum())
df_tm_this_month.iloc[8,0] = 'CA$ '+str(vacant[0])
df_tm_this_month.iloc[9,0] = 'CA$ '+str(vacant[1])
df_tm_this_month.iloc[10,0] = str(round(df_base'Vacant'].sum()/total_rent*100,2))+'%'
df_tm_this_month.iloc[11,0] = str(round(percent_vacant[0],2))+'%'
df_tm_this_month.iloc[12,0] = str(round(percent_vacant[1],2))+'%'

#create base timeseries table
if end_date.month != 1
    df_tm_base[str(end_date)] = pd.Series(df_tm_this_month[str(end_date)])
else:
    df_tm_base = df_tm_this_month
    



#Saving final result
#export dataframes to excel using separate excel tabs
def dfs_tabs(df_list, sheet_list, file_name):
    writer = pd.ExcelWriter(file_name,engine='xlsxwriter')   
    for dataframe, sheet in zip(df_list, sheet_list):
        dataframe.to_excel(writer, sheet_name=sheet, startrow=0 , startcol=0)   
    writer.save()

# list of dataframes and sheet names
dfs = [df_base, df_tm_base]
sheets = ['Raw_Aggregated_Data','Timeseries_Analysis']    

# run function
dfs_tabs(dfs, sheets, filename,'lajeunesse_'+str(end_date.month)+str(end_date.year)+'TD'+'.xlsx')



