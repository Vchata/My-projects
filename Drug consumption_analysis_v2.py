# ENVIRONMENT
# ===============================
# Import modules
import os
import pandas as pd
import numpy as np
import openpyxl
import datetime
# Get current working directory
cwd = os.getcwd()
print(cwd)
# Remove SettingWithCopyWarning
pd.options.mode.chained_assignment = None


# FILE IMPORT PARAMETERS
# ===============================
# Create path/file variable
def file(name, extension, location):
    
    if location == 'data':
        file = os.path.join(cwd.replace('/Scripts/Consumption Analysis',''), 'data/' + name + extension)
    elif location == 'archive':
        file = os.path.join(cwd, 'archive/' + name + extension)
    elif location == 'output':
        file = os.path.join(cwd, 'output/' + name + extension)
       
    return file

# ********************************************************
# ********************************************************
# Specify group-by variables during computations (merging key)
group_by_value = 'Facility_Code'
#group_by_value = 'District'
#group_by_value = 'Province'
# Specify file names
history_file = 'Aug 2023 Dispensed Consumption Cumulative'
test_file = 'Sep 2023 Dispensed Consumption'
# Specify comparison variables
compare_tag = '202308'
compare_file = 'Anomaly Reports 202308 By Facility v1'
#compare_file = 'Anomaly Reports 202308 By District v1'
# Specify the year and month of the test data. This will drive what data is kept in the training and test datasets.
test_year = 2023
test_month = 9
# ********************************************************
# ********************************************************

# Specify the product list to analyze
product_list = []
#product_list = ['ARV0016','ARV0018','ARV0020','ARV0050','ARV0063','HTK0002','HTK0007','LAB7735','ARV0074','ARV0077','LAB7736']
# Specify the column names during read-in
df_vars_history = ['Facility_Code','Product_Code','Consumption_Year','MonthID','Consumption_Quantity']
df_vars_test = ['Province','District','Program_Area','Facility_Code','Facility_Type','Facility','Product_Code','Product','Consumption_Quantity','Consumption_Month','MonthID','Consumption_Year']
# Specify the variables to save before concatenation (facility code is lowest in hierarchy)
df_vars_keep = ['Facility_Code','Product_Code','Consumption_Year','MonthID','Consumption_Quantity','Facility','Product','Program_Area','Province','District']
# Specify the column names for the comparison files to keep
df_vars_compare = ['Invalid', 'Valid', 'Total', '% Valid', 'R', 'X', 'XR', 'Total Anomalies']


# STATIC GLOBAL VARIABLES
# ===============================
# Specify group-by variables during computations (merging key)
group_by_list = [group_by_value, 'Product_Code']
# Path and name of exemptions file
exempt_file = os.path.join(cwd, 'Output/' + 'Exemptions.xlsx')
# Create month count constant such that the test month is always month 25
month_count_constant = test_year*12 + test_month - 25
# Minimum sample sizes for lat 24 months and 6 months
samples_last_24M = 12
samples_last_6M = 5
samples_seasonality = 3


# EXPORT MODULE
# ===============================
def export_excel(df, outfile, sheet):
    
    # Specify file location
    file = os.path.join(cwd, 'Output/' + outfile)
    # If file already exists, add another sheet
    if os.path.exists(file):
        writer = pd.ExcelWriter(file, engine='openpyxl')
        book = openpyxl.load_workbook(file)
        writer.book = book
        # Check if the sheet exists. If it does, then delete.
        if sheet in book.sheetnames:
            book.remove(book[sheet])
        df.to_excel(writer, sheet_name=sheet, index=False)
        # Save and close the workbook
        writer.save()
        writer.close()
    # If file does not exist, create new workbook
    else:
        with pd.ExcelWriter(file) as writer:
            df.to_excel(writer, sheet_name=sheet, index=False)
    

# IMPORT MODULE
# ===============================
def import_data(data_source, dl, source):
    
    # Import the data
    # Ensure columns are the correct type - this ensures that sorting is done correctly
    if source == 'test':
        df_unsorted = pd.read_excel(data_source, names=df_vars_test, \
                                    dtype={'Province': str, \
                                           'District': str, \
                                           'Program_Area': str, \
                                           'Facility_Code': str, \
                                           'Facility_Type': str, \
                                           'Facility': str, \
                                           'Product_Code': str, \
                                           'Product': str, \
                                           'Consumption_Quantity': float, \
                                           'Consumption_Month': str, \
                                           'MonthID': float, \
                                           'Consumption_Year': float})
    
        # Clean variables of whitespace
        df_unsorted['Province'] = df_unsorted['Province'].str.strip()
        df_unsorted['District'] = df_unsorted['District'].str.strip()
        df_unsorted['Facility_Code'] = df_unsorted['Facility_Code'].str.strip()
        df_unsorted['Product_Code'] = df_unsorted['Product_Code'].str.strip()
        # Keep only certain variables
        df_unsorted_clean = df_unsorted[df_vars_keep]
        
        return df_unsorted_clean
    
    elif source == 'history':
        df_unsorted = pd.read_csv(data_source, delimiter=dl, names=df_vars_history, skiprows=1, encoding="ISO-8859-1", \
                                  dtype={'Facility_Code': str, \
                                         'Product_Code': str, \
                                         'Consumption_Year': float, \
                                         'MonthID': float, \
                                         'Consumption_Quantity': float})
        # Clean variables of whitespace
        df_unsorted['Facility_Code'] = df_unsorted['Facility_Code'].str.strip()
        df_unsorted['Product_Code'] = df_unsorted['Product_Code'].str.strip()
        # Add details from test data
        df_unsorted_with_details = pd.merge(df_unsorted, df_test_details_facility, how='left', left_on=['Facility_Code','Product_Code'], right_on=['Facility_Code','Product_Code'])
    
        return df_unsorted_with_details
   
    
# MODIFY DATA MODULE
# ===============================
def modify_data(data_source):
     
    def transform(df, column_in, column_out, column_out_value, filter_column, filter_values):
    
    #Create new data source
     data_source_new = df.copy()
     # Rename different Artemether + Lumefantrine products to be one product in the test file
     data_source_new[column_out] = np.where(data_source_new[filter_column].isin(filter_values), column_out_value, data_source_new[column_in])
     # Drop original column and rename new column
     data_source_new = data_source_new.drop(columns=[column_in])
     data_source_new = data_source_new.rename(columns={column_out: column_in})
            
     return data_source_new
    
    # Rename different Artemether + Lumefantrine products to be one product in the modified dataset
    #df_modified = transform(data_source, 'Product', 'Product_Modified', 'Artemether + Lumefantrine', 'Product_Code', ['MAL0001','MAL0002','MAL0003','MAL0004'])
    
    # Update consumption quantity for different Artemether + Lumefantrine products since there are many tablets in each pack size
    #df_modified = transform(df_modified, 'Consumption_Quantity', 'Consumption_Quantity_Modified', df_modified['Consumption_Quantity']*180, 'Product_Code', ['MAL0001'])
    #df_modified = transform(df_modified, 'Consumption_Quantity', 'Consumption_Quantity_Modified', df_modified['Consumption_Quantity']*360, 'Product_Code', ['MAL0002'])
    #df_modified = transform(df_modified, 'Consumption_Quantity', 'Consumption_Quantity_Modified', df_modified['Consumption_Quantity']*540, 'Product_Code', ['MAL0003'])
    #df_modified = transform(df_modified, 'Consumption_Quantity', 'Consumption_Quantity_Modified', df_modified['Consumption_Quantity']*720, 'Product_Code', ['MAL0004'])

    # Rename different Artemether + Lumefantrine products to be one product code in both the history and test files
    #df_modified = transform(df_modified, 'Product_Code', 'Product_Code_Modified', 'MAL0000', 'Product_Code', ['MAL0001','MAL0002','MAL0003','MAL0004'])
    
    # This dataset will only contained modified products
    #df_modified = df_modified.loc[df_modified['Product_Code'] == 'MAL0000']
    # Keep only certain variables
    #df_modified_clean = df_modified[df_vars_keep]
    
    #return df_modified_clean


# CONCATENATE MODULE
# ===============================
def concatenate_and_aggregate(df_list, source):
    # Concatenate the datasets

    df_concat = pd.concat(df_list, ignore_index=True)
        
    # Pull test data details if source is test data    
    if source == 'test':
        # Get distinct details for each facility code/product code combination to be merged with the history file
        # NOTE: Some combinations have more than one program area so keep the one with highest consumption
        df_details_by_facility_sorted = df_concat.sort_values(by=['Facility_Code','Product_Code','Consumption_Quantity'], kind='mergesort', ascending=[True, True, False])
        df_details_by_facility = df_details_by_facility_sorted[['Facility_Code','Product_Code','Facility','Product','Program_Area','Province','District']].drop_duplicates(['Facility_Code','Product_Code'], keep='first')        
        # Because details are only by facility-product combinations, they need be re-sorted to remove duplicates
        # if the group-by value is not facility code
        df_details_by_other_sorted = df_concat.sort_values(by=[group_by_value,'Product_Code','Consumption_Quantity'], kind='mergesort', ascending=[True, True, False])
        df_details_by_other = df_details_by_other_sorted[[group_by_value,'Product_Code','Product','Program_Area']].drop_duplicates(group_by_list, keep='first')       
        
        return df_concat, df_details_by_facility, df_details_by_other
    
    # Otherwise, aggregate all datasets
    elif source == 'all':
        # Create a cumulative file
        df_cumulative = df_concat[['Facility_Code','Product_Code','Consumption_Year','MonthID','Consumption_Quantity']]
        # Specify how the dataset will be grouped
        df_group_vars = group_by_list + ['Consumption_Year','MonthID']
        # Aggregate consumption quantity by group-by list and month
        df_concat_unsorted_agg = df_concat.groupby(df_group_vars)['Consumption_Quantity'].sum().reset_index()   
        # Sort the dataset
        df_concat_sorted_agg = df_concat_unsorted_agg.sort_values(by=df_group_vars, kind='mergesort')
    
        return df_cumulative, df_concat_sorted_agg


# DATA SPLIT MODULE
# ===============================
def split_data(df):
    ###Clare Update###
    df = df[df['Consumption_Quantity'] > 0]
    ##################

    # Get the lag of consumption quantity for computations
    df['Consumption_Quantity_Lag1'] = df.groupby(group_by_list)['Consumption_Quantity'].shift(1)        
    # Add a month count variable
    df['Month_Count'] = df['Consumption_Year']*12 + df['MonthID'] - month_count_constant 
    # Get the lag of month count for computations
    df['Month_Count_Lag1'] = df.groupby(group_by_list)['Month_Count'].shift(1)
    print('> New Variables Created')

    # Get the training dataset
    consumption_train = df.loc[(df['Month_Count'] > 0) & (df['Month_Count'] < 25)]
    # Get the test dataset
    consumption_test = df.loc[(df['Month_Count'] == 25)]
    print('> New Datasets Created')
    
    return consumption_train, consumption_test


# OUTLIERS MODULE
# ===============================
def compute_outliers(train):
    
    # OUTLIERS
    # ====================
    # Keep only the records in the specified product list; otherwise, use entire datase
    if not product_list:
        df_subset = train.copy()
    else:
        df_subset = train[train['Product_Code'].isin(product_list)]
    # Calculate 1st quantile
    df_subset_q1 = df_subset.groupby(group_by_list)['Consumption_Quantity'].quantile(0.25).reset_index().rename(columns={'Consumption_Quantity': 'Q1'})
    print('> Q1 Complete')
    # Calculate 3rd quantile
    df_subset_q3 = df_subset.groupby(group_by_list)['Consumption_Quantity'].quantile(0.75).reset_index().rename(columns={'Consumption_Quantity': 'Q3'})
    print('> Q3 Complete')
    # Merge 1st and 3rd quantiles
    df_quantile = pd.merge(df_subset_q1, df_subset_q3, how='inner', left_on=group_by_list, right_on=group_by_list)
    print('> Quantile Merge Complete')  
    # Calculate outlier upper and lower limits
    df_quantile['Outlier_Upper'] = df_quantile['Q3'] + 1.5*(df_quantile['Q3'] - df_quantile['Q1'])
    df_quantile['Outlier_Lower'] = np.maximum(0, df_quantile['Q1'] - 1.5*(df_quantile['Q3'] - df_quantile['Q1']))
    print('> Limits Complete')
    # Merge outlier information with original dataset
    df_subset_outlier = pd.merge(df_subset, df_quantile, how='left', left_on=group_by_list, right_on=group_by_list)
    print('> Outliers Merge Complete')
    # Add outlier flags
    df_subset_outlier['Outlier'] = np.where((df_subset_outlier['Consumption_Quantity'] > df_subset_outlier['Outlier_Upper']) | (df_subset_outlier['Consumption_Quantity'] < df_subset_outlier['Outlier_Lower']), 1, 0)
    df_subset_outlier['Outlier_Lag1'] = df_subset_outlier.groupby(group_by_list)['Outlier'].shift(1)
    print('> Subset Complete')
    
    return df_subset_outlier
 
    
# PARAMETERS MODULE
# ===============================
def compute_parameters(df_subset_outlier):
    
    # I-MR CHART PARAMETERS
    # ====================
    # If records are not in consecutive months or if the current/previous month consumption is an outlier, a range cannot be calculated
    ### Clare Update ###
    '''
    df_subset_outlier['Range'] = np.where(((df_subset_outlier['Month_Count'] - df_subset_outlier['Month_Count_Lag1']) > 1) | (df_subset_outlier['Outlier'] == 1) | (df_subset_outlier['Outlier_Lag1'] == 1), \
        np.NaN, abs(df_subset_outlier['Consumption_Quantity'] - df_subset_outlier['Consumption_Quantity_Lag1']))
    '''
    df_subset_outlier['Range'] = np.where((df_subset_outlier['Outlier'] == 1) | (df_subset_outlier['Outlier_Lag1'] == 1), \
        np.NaN, abs(df_subset_outlier['Consumption_Quantity'] - df_subset_outlier['Consumption_Quantity_Lag1']))
    #####################
    # Calculate x-bar parameters for non-outliers only
    df_subset_outlier_xbar = df_subset_outlier.loc[(df_subset_outlier['Outlier'] == 0)].groupby(group_by_list)['Consumption_Quantity'].mean().reset_index().rename(columns={'Consumption_Quantity': 'X_Bar'})
    # Calculate r-bar parameters for non-outliers only    
    df_subset_outlier_rbar = df_subset_outlier.loc[(df_subset_outlier['Outlier'] == 0)].groupby(group_by_list)['Range'].mean().reset_index().rename(columns={'Range': 'R_Bar'})
    # Merge x-bar and r-bar parameters
    df_subset_outlier_bars = pd.merge(df_subset_outlier_xbar, df_subset_outlier_rbar, how='inner', left_on=group_by_list, right_on=group_by_list)
    # Calculate x-bar and r-bar upper and lower limits
    df_subset_outlier_bars['X_UCL'] = df_subset_outlier_bars['X_Bar'] + (3*df_subset_outlier_bars['R_Bar'])/1.128
    df_subset_outlier_bars['X_LCL'] = np.maximum(0, df_subset_outlier_bars['X_Bar'] - (3*df_subset_outlier_bars['R_Bar'])/1.128)
    df_subset_outlier_bars['R_UCL'] = 3.267*df_subset_outlier_bars['R_Bar']
    # Merge x-bar and r-bar information with original dataset
    df_subset_temp = pd.merge(df_subset_outlier, df_subset_outlier_bars, how='left', left_on=group_by_list, right_on=group_by_list)
    # Get non-zero dataset
    df_subset_temp_non_zero = df_subset_temp[df_subset_temp['Consumption_Quantity'] > 0]
    print('> Parameters Complete')
    
    # QC PARAMETERS
    # ====================
    # Each product combination should only have one unique x-bar and r-bar
    df_subset_temp_qc = df_subset_temp.groupby(group_by_list)[['X_Bar','R_Bar']].nunique().reset_index()
    # Output any combination with more than 1 x-bar or 1 r-bar
    df_subset_temp_qc_counts = df_subset_temp_qc.loc[(df_subset_temp_qc['X_Bar'] > 1) | (df_subset_temp_qc['R_Bar'] > 1)]
    # Check if there are any duplicate product combinations
    df_subset_temp_qc['Key'] = df_subset_temp_qc.drop(['X_Bar','R_Bar'], axis=1).sum(axis=1)
    # Find duplicate keys
    df_subset_temp_qc_combos_temp = df_subset_temp_qc['Key']
    df_subset_temp_qc_combos = df_subset_temp_qc_combos_temp[df_subset_temp_qc_combos_temp.duplicated(keep=False)]
    print('> QC Complete')

    # COUNTS
    # ====================
    # Calculate number of valid samples (including zeroes)
    zero_obs = df_subset_temp.loc[(df_subset_temp['Outlier'] == 0)].groupby(group_by_list)['Month_Count'].nunique().reset_index().rename(columns={'Month_Count': 'Total_Month_Obs'})
    # Calculate number of valid samples (excluding zeroes)
    non_zero_obs = df_subset_temp_non_zero.loc[(df_subset_temp['Outlier'] == 0)].groupby(group_by_list)['Month_Count'].nunique().reset_index().rename(columns={'Month_Count': 'Total_Non_Zero_Month_Obs'})
    # Number of valid samples (including zeroes) in last 6 months
    ### Clare Update ###
    critical_obs = df_subset_temp.loc[((df_subset_temp['Month_Count'] >= 12) & (df_subset_temp['Month_Count'] <= 14)) & (df_subset_temp['Outlier'] == 0)] \
        .groupby(group_by_list)['Month_Count'].nunique().reset_index().rename(columns={'Month_Count': 'Total_Valid_Obs_Year_Prior'})
    ####################
    last_6_months_obs = df_subset_temp.loc[(df_subset_temp['Month_Count'] > 18) & (df_subset_temp['Outlier'] == 0)] \
        .groupby(group_by_list)['Month_Count'].nunique().reset_index().rename(columns={'Month_Count': 'Total_Valid_Obs_Last_6M'})
    # Merge obs datasets
    obs_1 = pd.merge(zero_obs, non_zero_obs, how='outer', left_on=group_by_list, right_on=group_by_list)
    obs_2 = pd.merge(obs_1, last_6_months_obs, how='outer', left_on=group_by_list, right_on=group_by_list)
    ###Clare Update###
    obs_3 = pd.merge(obs_2, critical_obs, how='outer', left_on=group_by_list, right_on=group_by_list)
    ##################
    # Create final dataset
    df_subset_24M = pd.merge(df_subset_temp, obs_3, how='left', left_on=group_by_list, right_on=group_by_list)
    print('> Counts Complete')
    
    # UNIQUE COMBINATIONS
    # ====================    
    # List of variables to drop in train data
    drop_list_subset_24M = ['Consumption_Year','MonthID','Consumption_Quantity','Consumption_Quantity_Lag1','Month_Count','Month_Count_Lag1','Outlier','Outlier_Lag1','Range']
    # Get list of product combinations and their associated parameters
    df_subset_unique = df_subset_24M.drop(drop_list_subset_24M, axis=1).drop_duplicates()
    print('> Train Data Complete')
    
    return df_subset_24M, df_subset_unique, df_subset_temp_qc_counts, df_subset_temp_qc_combos


# SCORING MODULE
# ===============================
def score_test_data(train, test, exempt):
    
    # TEST DATA
    # ====================
    # List of variables to keep in test data
    keep_list = group_by_list + ['Consumption_Quantity','Consumption_Quantity_Lag1','Month_Count','Month_Count_Lag1']
    # Keep only specified variables and products in test data
    if not product_list:
        df_test_reduced = test.copy()
    else:
        df_test_reduced = test[test['Product_Code'].isin(product_list)][keep_list]
    # Merge test data with unique combination parameters for scoring
    df_test = pd.merge(df_test_reduced, train, how='left', left_on=group_by_list, right_on=group_by_list)
    # Add investigation month    
    df_test['Investigation_Month'] = str(test_year) + '-' + str(test_month) + '-1'
    df_test['Investigation_Month'] = pd.to_datetime(df_test['Investigation_Month'])
    # Calculate outlier information and range for test data
    df_test['Outlier'] = np.where((df_test['Consumption_Quantity'] > df_test['Outlier_Upper']) | (df_test['Consumption_Quantity'] < df_test['Outlier_Lower']), 1, 0)
    df_test['Outlier_Lag1'] = np.where((df_test['Consumption_Quantity_Lag1'] > df_test['Outlier_Upper']) | (df_test['Consumption_Quantity_Lag1'] < df_test['Outlier_Lower']), 1, 0)
    ### Clare Update ###
    '''
    df_test['Range'] = np.where(((df_test['Month_Count'] - df_test['Month_Count_Lag1']) > 1) | (df_test['Outlier_Lag1'] == 1), \
        np.NaN, abs(df_test['Consumption_Quantity'] - df_test['Consumption_Quantity_Lag1']))
        '''
    df_test['Range'] = np.where((df_test['Outlier_Lag1'] == 1), \
        np.NaN, abs(df_test['Consumption_Quantity'] - df_test['Consumption_Quantity_Lag1']))
    ####################
    print('> Test Data Complete')
   
    # ADD EXEMPTIONS
    # ====================    
    # Because exemptions are only by facility-product combinations, a new variable needs to be created
    # if the group-by value is not facility code
    if group_by_value != 'Facility_Code':
        exempt[group_by_value] = np.NaN
        # Keep only the group-by variables and exemption variables (5 total)
        df_keep_vars = group_by_list + ['Exemption_Period_Start','Exemption_Period_End','Exemption_Code']
        exempt = exempt[df_keep_vars]
        
    # Add exemptions to test dataset
    df_test_exempt = pd.merge(df_test, exempt, how='left', left_on=group_by_list, right_on=group_by_list)
        
    print('> Exemptions Complete')
   
    # SCORE DATA
    # ====================
    # Add exemption variable
    df_test_exempt['Exemption_Status'] = np.where((df_test_exempt['Exemption_Period_Start'] <= df_test_exempt['Investigation_Month']) & (df_test_exempt['Investigation_Month'] <= df_test_exempt['Exemption_Period_End']), 'E', 'N')
    # Calculate if test data breaches I-MR charts
    df_test_exempt['X_Breach'] = np.where((df_test_exempt['Consumption_Quantity'] > df_test_exempt['X_UCL']) | (df_test_exempt['Consumption_Quantity'] < df_test_exempt['X_LCL']), 1, 0)
    df_test_exempt['R_Breach'] = np.where((df_test_exempt['Range'] > df_test_exempt['R_UCL']), 1, 0)
    # Create anomaly code
    def anomaly_code(df):
        if ((df['X_Breach'] == 1) and (df['R_Breach'] == 1)):
            return 'XR'
        elif ((df['X_Breach'] == 1) and (df['R_Breach'] == 0)):
            return 'X'
        elif ((df['X_Breach'] == 0) and (df['R_Breach'] == 1)):
            return 'R'
        elif ((df['X_Breach'] == 0) and (df['R_Breach'] == 0)):
            return np.NaN
    df_test_exempt['Anomaly_Code'] = df_test_exempt.apply(anomaly_code, axis=1)
    print('> Score Complete')
    
    # FILTER DATA
    # ====================
    # Flag valid combinations
    df_test_exempt['Valid_Combo'] = np.where((df_test_exempt['Total_Month_Obs'] >= samples_last_24M) & (df_test_exempt['Total_Valid_Obs_Last_6M'] >= samples_last_6M) & (df_test_exempt['Total_Valid_Obs_Year_Prior'] == samples_seasonality), 1, 0)
    # Get valid combinations only
    df_valid_obs = df_test_exempt.loc[(df_test_exempt['Valid_Combo'] == 1)]
    # Keep only breaching test data
    df_valid_breach = df_valid_obs.loc[df_valid_obs['Anomaly_Code'].isin(['X','R','XR'])]
    print('> Filter Complete')
    
    # RANK DATA
    # ====================
    # Calculate anomaly deviations;
    def calc_x_delta(df):   
        if (df['Consumption_Quantity'] > df['X_UCL']):
            if (df['X_Bar'] == 0):
                return 1 
            else:
                return (df['Consumption_Quantity'] - df['X_UCL'])/df['X_Bar']
        elif (df['Consumption_Quantity'] < df['X_LCL']):
            if (df['X_Bar'] == 0):
                return 1
            else:
                return (df['X_LCL'] - df['Consumption_Quantity'])/df['X_Bar']
        else:
            return np.NaN
    def calc_r_delta(df): 
        if (df['Range'] > df['R_UCL']):
            if (df['R_Bar'] == 0):
                return 1
            else:
                return (df['Range'] - df['R_UCL'])/df['R_Bar']
        else:
            return np.NaN
    df_valid_breach['X_Delta'] = df_valid_breach.apply(calc_x_delta, axis=1)
    df_valid_breach['R_Delta'] = df_valid_breach.apply(calc_r_delta, axis=1)
    # Calculate anomaly rankings
    df_valid_breach['X_Rank'] = df_valid_breach['X_Delta'].rank(ascending=False)
    df_valid_breach['R_Rank'] = df_valid_breach['R_Delta'].rank(ascending=False)
    
    # MERGE DETAILS
    # ====================
    # Main output variables
    output_variables_main = ['Consumption_Quantity','Range','Anomaly_Code','X_Rank','R_Rank', \
                            'Exemption_Status','Exemption_Code','Exemption_Period_Start','Exemption_Period_End', \
                            'Outlier','Q1','Q3','Outlier_Upper','Outlier_Lower','X_Bar','R_Bar','X_UCL','X_LCL','R_UCL', \
                            'Total_Month_Obs','Total_Non_Zero_Month_Obs','Total_Valid_Obs_Last_6M','Total_Valid_Obs_Year_Prior']
    
    # Merge details onto valid dataset
    if group_by_value != 'Facility_Code':
        df_valid_breach_merge = pd.merge(df_valid_breach, df_test_details_other, how='left', left_on=group_by_list, right_on=group_by_list)
        output_variables_all = ['Investigation_Month',group_by_value,'Product_Code','Product','Program_Area'] + output_variables_main
    else:
        df_valid_breach_merge = pd.merge(df_valid_breach, df_test_details_facility, how='left', left_on=group_by_list, right_on=group_by_list)
        output_variables_all = ['Investigation_Month','Facility_Code','Product_Code','Facility','Product','Program_Area'] + output_variables_main
   
    # Drop variables
    df_valid_final = df_valid_breach_merge[output_variables_all]
    
    # Check for duplicate combinations
    df_qc_combos_final = df_valid_final[df_valid_final.duplicated(group_by_list, keep=False)]
    
    print('> Rank Complete')
            
    # FREQUENCIES
    # ====================
    # Find number of valid and invalid product combinations by the group-by list
    df_valid_product_counts = pd.crosstab(index=df_test_exempt['Product_Code'], columns=df_test_exempt['Valid_Combo']).rename(columns={0: 'Invalid', 1:'Valid'}).reset_index()
    df_valid_group_by_counts = pd.crosstab(index=df_test_exempt[group_by_value], columns=df_test_exempt['Valid_Combo']).rename(columns={0: 'Invalid', 1:'Valid'}).reset_index()   
    # Get total combinations
    df_valid_product_counts['Total'] = df_valid_product_counts['Invalid'] + df_valid_product_counts['Valid']
    df_valid_group_by_counts['Total'] = df_valid_group_by_counts['Invalid'] + df_valid_group_by_counts['Valid']
    # Get percent valid combinations
    df_valid_product_counts['% Valid'] = df_valid_product_counts['Valid']/df_valid_product_counts['Total']
    df_valid_group_by_counts['% Valid'] = df_valid_group_by_counts['Valid']/df_valid_group_by_counts['Total']
    # Find number of anomalies by product code and type of anomaly
    df_anomaly_product_counts = pd.crosstab(index=df_valid_final['Product_Code'], columns=df_valid_final['Anomaly_Code']).reset_index()
    df_anomaly_group_by_counts = pd.crosstab(index=df_valid_final[group_by_value], columns=df_valid_final['Anomaly_Code']).reset_index()
    # Merge all counts
    df_product_counts_final = pd.merge(df_valid_product_counts, df_anomaly_product_counts, how='left', left_on=['Product_Code'], right_on=['Product_Code'])
    df_group_by_counts_final = pd.merge(df_valid_group_by_counts, df_anomaly_group_by_counts, how='left', left_on=[group_by_value], right_on=[group_by_value])    
    # Get total anomalies
    df_product_counts_final['Total Anomalies'] = df_product_counts_final['X'] + df_product_counts_final['R'] + df_product_counts_final['XR']
    df_group_by_counts_final['Total Anomalies'] = df_group_by_counts_final['X'] + df_group_by_counts_final['R'] + df_group_by_counts_final['XR']
    print('> Frequencies Complete')
    
    return df_valid_final, df_qc_combos_final, df_product_counts_final, df_group_by_counts_final


# COMPARISON MODULE
# ===============================
def compare_results(new_file, old_file, old_sheet, old_tag, merge_key):
    
    # RENAME FUNCTION
    # ====================
    def rename_cols(dataset):
        rename_list = {}
        name_list = list(dataset)
        for i in name_list:
            if i not in ('Product_Code', merge_key):
                rename_list.update({i: i + ' ' + old_tag})
            
        return rename_list
    
    # IMPORT & RENAME
    # ====================
    # Import frequencies by product      
    df_old_import = pd.read_excel(old_file, sheet_name=old_sheet)
    # Keep only certain columns
    df_old_import_cleaned = df_old_import[[merge_key] + df_vars_compare]
    # Get list of column names
    df_old_rename = rename_cols(df_old_import_cleaned)
    # Rename columns with appended tag
    df_old = df_old_import_cleaned.rename(columns=df_old_rename)
    print(list(df_old))
    print('> Old File Complete')

    # CREATE COMPARISON TABLE
    # ====================  
    # Merge old and new frequency tables
    df_compare = pd.merge(new_file, df_old, how='left', left_on=[merge_key], right_on=[merge_key])
    for new, old in df_old_rename.items():
        # Calculate percent change
        change_in_new = '% Chg In ' + new
        df_compare[change_in_new] = (df_compare[new] - df_compare[old])/df_compare[old]
        # Replace inf with 0
        df_compare[change_in_new].replace(-np.inf, np.nan, inplace=True)
        df_compare[change_in_new].replace(np.inf, np.nan, inplace=True)
    print('> Compare Complete')
        
    return df_compare

def watch_list(new_file, old_file, old_sheet, sequence):
    
    # For new analysis, months on list will equal one
    if sequence == 'new':
        df_watch_list = df_investigation.copy()
        df_watch_list['Months_On_List'] = 1
        
        return df_watch_list

    # If continuation, pull months on list from old file and add one
    else:
        # Import old investigation report  
        df_old_import = pd.read_excel(old_file, sheet_name=old_sheet)
        df_old_import_reduced = df_old_import[group_by_list + ['Months_On_List']]
        
        # Merge old file with new file
        df_watch_list = pd.merge(new_file, df_old_import_reduced, how='left', left_on=group_by_list, right_on=group_by_list)
        df_watch_list['Months_On_List'].fillna(0, inplace=True)
        df_watch_list['Months_On_List'] += 1
        print('> Watch List Complete')

    return df_watch_list


# ANOMALY HISTORY
# ===============================
def anomaly_history(anomaly, main, parameters):
    
    # List of variables to keep in investigation report
    keep_list_anomaly = group_by_list + ['Investigation_Month']
    # Create reduced list of valid combinations to investigate
    df_valid_final_reduced = anomaly[keep_list_anomaly]
    # List of variables to keep in main datasets
    keep_list_main = group_by_list + ['Consumption_Year','MonthID','Consumption_Quantity']
    keep_list_parameters = group_by_list + ['Consumption_Year','MonthID','Range','Outlier']
    # Merge with main dataset
    df_anomaly_history_main = pd.merge(df_valid_final_reduced, main[keep_list_main], how='left', left_on=group_by_list, right_on=group_by_list)
    # Merge with parameters dataset
    df_anomaly_history = pd.merge(df_anomaly_history_main, parameters[keep_list_parameters], how='left', left_on=group_by_list + ['Consumption_Year','MonthID'], right_on=group_by_list + ['Consumption_Year','MonthID'])
    # Create consumption month
    df_anomaly_history['year'] = df_anomaly_history['Consumption_Year']
    df_anomaly_history['month'] = df_anomaly_history['MonthID']
    df_anomaly_history['day'] = 1
    df_anomaly_history['Consumption_Month'] = pd.to_datetime(df_anomaly_history[['year', 'month', 'day']])
    
    # MERGE FACILITY AND PRODUCT DETAILS
    # ====================
    # Merge details onto valid dataset
    if group_by_value != 'Facility_Code':
        df_anomaly_history_merge = pd.merge(df_anomaly_history, df_test_details_other, how='left', left_on=group_by_list, right_on=group_by_list)
    else:
        df_anomaly_history_merge = pd.merge(df_anomaly_history, df_test_details_facility, how='left', left_on=group_by_list, right_on=group_by_list)
   
    # Keep only specific variables
    df_anomaly_history_reduced = df_anomaly_history_merge[['Investigation_Month',group_by_value,'Product_Code','Program_Area','Consumption_Month','Consumption_Quantity','Range','Outlier']]
    
    return df_anomaly_history_reduced




# =================== #
#      EXECUTION      #
# =================== #

print('STEP 1: Importing data...')
print('> Running Test Data')
# ====================
# TEST FILE
# ====================
# Import test file
df_test_raw = import_data(file(test_file, '.xlsx', 'data'), '' , 'test')
# Check data and data types
df_test_raw.dtypes
print(df_test_raw)
# Modify test data
df_test_raw_modified = modify_data(df_test_raw)
# Combine test files and get details by facility/other
df_test_concat, df_test_details_facility, df_test_details_other = concatenate_and_aggregate([df_test_raw, df_test_raw_modified], 'test')

print('> Running History Data')
# ====================
# HISTORY FILE
# ====================
# Import history file with details added
df_history_raw = import_data(file(history_file, '.txt', 'data'), '\t', 'history')
# Check data and data types
df_history_raw.dtypes
print(df_history_raw)
# Modify history data (if necessary)
#df_history_raw_modified = modify_data(df_history_raw)

# Combine history and test files and aggregate based on group-by list
print('> Concatenating & Aggregating Data')
df_main, df_main_aggregated = concatenate_and_aggregate([df_history_raw, df_test_concat], 'all')
print(df_main_aggregated)

# Output concatenated file
print('> Saving Concatenated Data')
df_main.to_csv(file(test_file + ' Cumulative', '.txt', 'data'), sep='\t', index=False)
print('-- FINISHED --')


## EXTRA ANALYSIS
## Output limited concatenated file
#df_main_limited = df_main_aggregated[df_main_aggregated['Facility'].str.contains('Kanyama 1st Level Hospital')]
#df_main_limited['year'] = df_main_limited['Consumption_Year']
#df_main_limited['month'] = df_main_limited['MonthID']
#df_main_limited['day'] = 1
#df_main_limited['Consumption_Month'] = pd.to_datetime(df_main_limited[['year', 'month', 'day']])
#df_main_limited_output = pd.pivot_table(df_main_limited[['Product','Consumption_Month','Consumption_Quantity']], values='Consumption_Quantity', index='Product', columns='Consumption_Month')
#print(df_main_limited_output)
#df_main_limited_output.to_csv(file('Kanyama', '.csv', 'data'), sep=',', index=True)


# Import exceptions file
print('STEP 2: Importing exemptions...')
df_exempt = pd.read_excel(exempt_file)
print('-- FINISHED --')

# Split main dataset into training and testing data
print('STEP 3: Splitting data...')
c_train, c_test = split_data(df_main_aggregated)
print('-- FINISHED --')

# Get 24-month history
print('STEP 4: Creating 24-month history...')
df_subset_outlier = compute_outliers(c_train)
df_subset_24M, df_subset_unique, df_qc_counts, df_qc_combos = compute_parameters(df_subset_outlier)
print('-- FINISHED --')

# Export to excel
print('STEP 5: Exporting QC Check...')
#Changed to_excel(df_qc_counts, 'QC Checks.xlsx')
# Cahnged to_excel(df_qc_combos, 'QC Checks.xlsx')

df_qc_counts.to_excel('QC Checks.xlsx',index = False, header=True)
df_qc_combos.to_excel('QC Checks.xlsx',index = False, header=True)
print('-- FINISHED --')
#Creating workbook
import xlsxwriter
Anomaly_Reports = pd.ExcelWriter('Anomaly_Reports.xlsx', engine='xlsxwriter')



# Get investigation report
print('STEP 6a: Creating investigation data...')
df_investigation, df_qc_combos_final, df_freq_by_product, df_freq_by_group_by = score_test_data(df_subset_unique, c_test, df_exempt)
df_watch_list = watch_list(df_investigation, file(compare_file, '.xlsx', 'archive'), 'Anomaly Investigation Report', 'continue')
print('STEP 6b: Exporting QC Check...')
#changed export_excel(df_qc_combos_final, 'QC Checks.xlsx', 'Duplicate Combos Final')

df_qc_combos_final.to_excel('QC Checks.xlsx',index = False, header=True)
print('-- FINISHED --')

# Compare frequencies from previous month
print('STEP 7: Comparing frequency data...')
df_freq_by_product_compare = compare_results(df_freq_by_product, file(compare_file, '.xlsx', 'archive'), 'Frequencies (By Product)', compare_tag, 'Product_Code')
df_freq_by_group_by_compare = compare_results(df_freq_by_group_by, file(compare_file, '.xlsx', 'archive'), 'Frequencies (By ' + group_by_value + ')', compare_tag, group_by_value)

# Export to excel
print('STEP 8: Exporting investigation data...')
# Changed export_excel(df_watch_list, 'Anomaly Reports.xlsx', 'Anomaly Investigation Report')
#Changed export_excel(df_freq_by_product_compare, 'Anomaly Reports.xlsx', 'Frequencies (By Product)')
#changed export_excel(df_freq_by_group_by_compare, 'Anomaly Reports.xlsx', 'Frequencies (By ' + group_by_value + ')')
df_watch_list.to_excel(Anomaly_Reports,sheet_name='Anomaly Investigation Report')
df_freq_by_product_compare.to_excel(Anomaly_Reports,sheet_name='Frequencies (By Product)') #,index = False, header=True)
df_freq_by_group_by_compare.to_excel(Anomaly_Reports,sheet_name='Frequencies (By ' + group_by_value + ')')

print('-- FINISHED --')

# Get anomaly history report
print('STEP 9: Creating anomaly history report...')
df_anomaly = anomaly_history(df_investigation, df_main_aggregated, df_subset_24M)
print('-- FINISHED --')

# Export to excel
print('STEP 10: Exporting anomaly history report...')
#Changed export_excel(df_anomaly, 'Anomaly Reports.xlsx', 'Anomaly History Report')
df_anomaly.to_excel(Anomaly_Reports,sheet_name='Anomaly History Report')
Anomaly_Reports.close()
#Anomaly_Reports.to_excel('C:\Zambia Consumption Threshold Analysis (2)\Zambia Consumption Threshold Analysis\archive', index=False)
print('-- FINISHED --')











# ================= #
#      TESTING      #
# ================= #

#def filter_data(dataset, product):
#    df = dataset.loc[(dataset['Product_Code'] == product)]
#    df.to_csv(os.path.join(cwd, 'Output/' + 'Test_1.csv'))
#filter_data(df_subset_unique, 'ARV0050')
#filter_data(df_main_aggregated[['Facility','Product_Code','Consumption_Year','MonthID','Consumption_Quantity']], 'ARV0050')
#filter_data(df_main_aggregated, 'ARV0050')

#def month_counts(dataset):
#    df = dataset.loc[(dataset['Product_Code'] == 'HTK0002')]
#    df_counts = df.groupby(['Consumption_Year','MonthID'])['Consumption_Quantity'].count().reset_index()
#    df_counts.to_csv(os.path.join(cwd, 'Output/' + 'Test_2.csv'))
#month_counts(df_main_aggregated)

#df_MAL_products = df_test_raw[df_test_raw['Product_Code'].str.contains('MAL0001')][['Product','Consumption_Quantity','Consumption_Quantity_Modified']]
#df_MAL_products = df_test_raw[df_test_raw['Product'].str.contains('Artemether')][['Product_Code','Product']]
#df_MAL_products_unique = df_MAL_products.drop_duplicates()
#df_MAL_products = df_test_raw[df_test_raw['Product_Code'].str.contains('MAL')][['Product_Code','Product']]
#df_MAL_products_unique = df_MAL_products.drop_duplicates()
#df_MAL_products = df_history_raw[df_history_raw['Product_Code'].str.contains('MAL')][['Product_Code', 'Consumption_Quantity']]
#df_MAL_products_unique = df_MAL_products.drop_duplicates()

#df_unique_districts = df_main_aggregated[['District']].drop_duplicates()
#df_unique_provinces = df_main_aggregated[['Province']].drop_duplicates()
#df = df_test_raw.loc[(df_test_raw['Product_Code'] == 'ARV0018') & (df_test_raw['Province'] == 'Central')]
#df = df_history_raw.loc[(df_history_raw['Product_Code'] == 'ARV0018') & (df_history_raw['Province'] == 'Central')]