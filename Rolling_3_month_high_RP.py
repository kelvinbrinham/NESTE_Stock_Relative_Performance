import pandas as pd
import numpy as np
import openpyxl as xl
import datetime
from datetime import datetime as dt
import xlsxwriter as xlw
import matplotlib.pyplot as plt
import dateutil.relativedelta as reldel

#Set how much of a dataframe is displayed when running the scipt
pd.set_option('display.max_rows', 10)
pd.set_option('display.max_columns', 6)

#<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>

#Load in the stock and index data
Neste_and_SXXP_price_df = pd.read_excel('Neste_and_SXXP price.xlsx', skiprows=[x for x in range(6)])

#Drop the first 2 empty columns from the excel sheet
Neste_and_SXXP_price_df.dropna(axis=1, how='all', inplace=True)

#Rename columns for convenience
Neste_and_SXXP_price_df.rename(columns={'Dates': 'Date', 'PX_LAST': 'NESTE_FH_last_PX', 'PX_LAST.1': 'SXXP_last_PX'}, inplace=True)

#Add column for relative perfomance (RP = Stock price / Index price = NESTE_FH_last_PX/SXXP_last_PX)
Neste_and_SXXP_price_df['RP'] = Neste_and_SXXP_price_df['NESTE_FH_last_PX'] / Neste_and_SXXP_price_df['SXXP_last_PX']


'''
Calculating rolling 3 month high RP and adding to the dataframe as a new column

I.e. Calulate the highest RP for each 3 month rolling period with 1 month DEFINED 1 MONTH AS 22 TRADING DAYS
'''
one_month_days = 22


#List of RP from the data
RP_lst = list(Neste_and_SXXP_price_df['RP'])
#List of Dates from the data
date_lst = list(Neste_and_SXXP_price_df['Date'])
#Set of dates in the data
date_st = set(date_lst) #Set of trading dates

#First date in the data
start_date = date_lst[0]
#Last date in the data
end_date = date_lst[-1]

#List of ALL dates within the data window (inc. non trading dates)
full_date_lst = list(pd.date_range(start_date, end_date, freq='d'))

#List of the RP data but with 0's for days it wasn't traded
#(This is convenient for indexing)
full_RP_lst = []

j = 0
for i in range(len(full_date_lst)):
    if full_date_lst[i] in date_st:
        full_RP_lst.append(RP_lst[j])
        j += 1

    else:
        full_RP_lst.append(0)



#List of ALL dates with the non trading days NTD masked with 'NTD'
full_date_lst = [x if x in date_st else 'NTD' for x in full_date_lst]


no_days_in_three_months = 3 * one_month_days
rolling_3_month_high_RP_lst = ([np.nan] * no_days_in_three_months)
rolling_3_month_high_RP_date_lst = ([np.nan] * no_days_in_three_months)

for i in range(no_days_in_three_months, len(date_lst)):
    three_month_RP_lst = RP_lst[(i-no_days_in_three_months):i+1]
    three_month_high_RP = max(three_month_RP_lst)
    rolling_3_month_high_RP_lst.append(three_month_high_RP)

    #Index of highest RP in three_month_RP_lst
    three_month_high_RP_index = np.argmax(three_month_RP_lst) + i - no_days_in_three_months
    #Date of the current 3 month high
    three_month_high_RP_date = date_lst[three_month_high_RP_index]
    #Add date to list
    rolling_3_month_high_RP_date_lst.append(three_month_high_RP_date)


Neste_and_SXXP_price_df['RP_3_month_high'] = rolling_3_month_high_RP_lst
Neste_and_SXXP_price_df['RP_3_month_high_date'] = rolling_3_month_high_RP_date_lst

# Neste_and_SXXP_price_df.to_excel('OUTPUT/TEST.xlsx', sheet_name='Sheet1', startrow = 2, index=False)

# plt.plot(date_lst, RP_lst, 'x', label = 'RP')
# plt.plot(date_lst, rolling_3_month_high_RP_lst, '+', label = 'RP 3 month high')
# plt.legend()
# plt.title('NESTE FH RP')
# plt.show()
