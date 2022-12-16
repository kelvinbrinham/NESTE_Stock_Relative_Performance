import pandas as pd
import numpy as np
import openpyxl as xl
import datetime
from datetime import datetime as dt
import xlsxwriter as xlw
import matplotlib.pyplot as plt
import dateutil.relativedelta as reldel
import statistics as stats
import functions as func

#Set how much of a dataframe is displayed when running the scipt
pd.set_option('display.max_rows', 10)
pd.set_option('display.max_columns', 7)

#<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>

from Rolling_3_month_high_RP import Neste_and_SXXP_price_df
from Rolling_3_month_high_RP import one_month_days


#<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
#Adding yes/no column for when there is a 3 month high
#I do this by checking whether the date of the high matches the date
#In the dates (first) column

RP_3_month_high_lst = list(Neste_and_SXXP_price_df['RP_3_month_high'])
RP_3_month_high_date_lst = list(Neste_and_SXXP_price_df['RP_3_month_high_date'])
date_lst = list(Neste_and_SXXP_price_df['Date'])
RP_lst = list(Neste_and_SXXP_price_df['RP'])

RP_3_month_high_date_st = set(RP_3_month_high_date_lst)
RP_3_month_high_y_or_n = []

for i in range(len(date_lst)):
    if date_lst[i] == RP_3_month_high_date_lst[i]:
        RP_3_month_high_y_or_n.append('Yes')

    else:
        RP_3_month_high_y_or_n.append('No')


Neste_and_SXXP_price_df['RP_3_month_high_y_or_n'] = RP_3_month_high_y_or_n


#<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
#Statistical data follwing a 3 month high RP for every date (inc. non trading days)
#Dates that do not correspond to a 3 month high RP are left empty
outperformance_dict = {'one_month': [], 'three_months': [], 'six_months': []}
RP_after_high_dict = {'one_month': [], 'three_months': [], 'six_months': []}

three_month_days = 3 * one_month_days
six_month_days = 6 * one_month_days


#<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
# 1 month outperformance

for i in range(len(date_lst) - one_month_days):
    if date_lst[i] not in RP_3_month_high_date_st:
        outperformance_dict['one_month'].append(np.nan)

    else:
        RP_in_one_month = RP_lst[i + one_month_days]
        RP_after_high_dict['one_month'].append(RP_in_one_month)
        RP_current = RP_lst[i]
        one_month_outperformance_current = RP_in_one_month / RP_current
        outperformance_dict['one_month'].append(one_month_outperformance_current)

#<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
# 3 month outperformance

for i in range(len(date_lst) - three_month_days):
    if date_lst[i] not in RP_3_month_high_date_st:
        outperformance_dict['three_months'].append(np.nan)

    else:
        RP_in_three_month = RP_lst[i + three_month_days]
        RP_after_high_dict['three_months'].append(RP_in_three_month)
        RP_current = RP_lst[i]
        three_month_outperformance_current = RP_in_three_month / RP_current
        outperformance_dict['three_months'].append(three_month_outperformance_current)

#<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
# 6 month outperformance

for i in range(len(date_lst) - six_month_days):
    if date_lst[i] not in RP_3_month_high_date_st:
        outperformance_dict['six_months'].append(np.nan)

    else:
        RP_in_six_month = RP_lst[i + six_month_days]
        RP_current = RP_lst[i]
        RP_after_high_dict['six_months'].append(RP_in_six_month)
        six_month_outperformance_current = RP_in_six_month / RP_current
        outperformance_dict['six_months'].append(six_month_outperformance_current)

#<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>


outperformance_dict['one_month'] = outperformance_dict['one_month'] + ([np.nan] * one_month_days)
outperformance_dict['three_months'] = outperformance_dict['three_months'] + ([np.nan] * three_month_days)
outperformance_dict['six_months'] = outperformance_dict['six_months'] + ([np.nan] * six_month_days)

#Wipe values of one month mean that are not inline with a 3 month high
for i in range(len(RP_3_month_high_y_or_n)):
    if RP_3_month_high_y_or_n[i] != 'Yes':
        outperformance_dict['one_month'][i] = np.nan
        outperformance_dict['three_months'][i] = np.nan
        outperformance_dict['six_months'][i] = np.nan



Neste_and_SXXP_price_df['1_month_outperformance'] = outperformance_dict['one_month']
Neste_and_SXXP_price_df['3_month_outperformance'] = outperformance_dict['three_months']
Neste_and_SXXP_price_df['6_month_outperformance'] = outperformance_dict['six_months']

# Neste_and_SXXP_price_df.to_excel('OUTPUT/TEST_2.xlsx', sheet_name='Sheet1', startrow = 2, index=False)
