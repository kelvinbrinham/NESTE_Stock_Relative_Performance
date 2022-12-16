import pandas as pd
import numpy as np
import openpyxl as xl
import datetime
from datetime import datetime as dt
import xlsxwriter as xlw
import matplotlib.pyplot as plt
import dateutil.relativedelta as reldel
import statistics as stats
from openpyxl import Workbook
from openpyxl.styles import PatternFill
import functions as func

#Set how much of a dataframe is displayed when running the scipt
pd.set_option('display.max_rows', 10)
pd.set_option('display.max_columns', 10)

#<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>

from Rolling_3_month_high_RP import Neste_and_SXXP_price_df
from Rolling_3_month_high_RP import one_month_days
from Calculating_stats import Neste_and_SXXP_HIGH_DATA_df
from Calculating_outperformance import RP_after_high_dict


#<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>

date_lst = Neste_and_SXXP_price_df['Date']
rolling_3_month_high_RP_lst = Neste_and_SXXP_price_df['RP_3_month_high']
rolling_3_month_high_RP_date_lst = Neste_and_SXXP_price_df['RP_3_month_high_date']
RP_lst = Neste_and_SXXP_price_df['RP']

rolling_3_month_high_RP_date_st = set(rolling_3_month_high_RP_date_lst)

rolling_3_month_high_RP_date_no_dupli_lst = list(rolling_3_month_high_RP_date_st)
rolling_3_month_high_RP_date_no_dupli_lst.sort()


date_one_month_forward_lst = []
date_three_months_forward_lst = []
date_six_months_forward_lst = []


for i in range(len(date_lst)):
    if date_lst[i] in rolling_3_month_high_RP_date_st:
        date_one_month_forward_lst.append(date_lst[i + one_month_days])
        if i < len(date_lst) - (3 * one_month_days):
            date_three_months_forward_lst.append(date_lst[i + (3 * one_month_days)])
        if i < len(date_lst) - (6 * one_month_days):
            date_six_months_forward_lst.append(date_lst[i + (6 * one_month_days)])


fig, ax = plt.subplots(figsize=(10,7))

RP, = ax.plot(date_lst, RP_lst, label = 'RP')
RP_3_month_high, = ax.plot(date_lst, rolling_3_month_high_RP_lst, label = 'RP 3-month High')

RP_1_month_after_3_month_high, = ax.plot(date_one_month_forward_lst, RP_after_high_dict['one_month'], label = 'RP 1 Month After 3-month High')
RP_3_month_after_3_month_high, = ax.plot(date_three_months_forward_lst, RP_after_high_dict['three_months'], label = 'RP 3 Months After 3-month High')
RP_6_month_after_3_month_high, = ax.plot(date_six_months_forward_lst, RP_after_high_dict['six_months'], label = 'RP 6 Months After 3-month High')

# RP_1_month_after_3_month_high, = ax.plot(rolling_3_month_high_RP_date_no_dupli_lst[1:], RP_after_high_dict['one_month'], label = '1 Month RP')
# RP_3_month_after_3_month_high, = ax.plot(rolling_3_month_high_RP_date_no_dupli_lst[2:], RP_after_high_dict['three_months'], label = '3 Month RP')
# RP_6_month_after_3_month_high, = ax.plot(rolling_3_month_high_RP_date_no_dupli_lst[14:], RP_after_high_dict['six_months'], label = '6 Month RP')


legend = plt.legend(loc='upper left')

RP_leg, RP_3_month_high_leg, RP_1_month_after_3_month_high_leg, RP_3_month_after_3_month_high_leg, RP_6_month_after_3_month_high_leg = legend.get_lines()

RP_leg.set_picker(True)
RP_leg.set_pickradius(10)

RP_3_month_high_leg.set_picker(True)
RP_3_month_high_leg.set_pickradius(10)

RP_1_month_after_3_month_high_leg.set_picker(True)
RP_1_month_after_3_month_high_leg.set_pickradius(10)

RP_3_month_after_3_month_high_leg.set_picker(True)
RP_3_month_after_3_month_high_leg.set_pickradius(10)

RP_6_month_after_3_month_high_leg.set_picker(True)
RP_6_month_after_3_month_high_leg.set_pickradius(10)

graphs_dict = {}
graphs_dict[RP_leg] = RP
graphs_dict[RP_3_month_high_leg] = RP_3_month_high
graphs_dict[RP_1_month_after_3_month_high_leg] = RP_1_month_after_3_month_high
graphs_dict[RP_3_month_after_3_month_high_leg] = RP_3_month_after_3_month_high
graphs_dict[RP_6_month_after_3_month_high_leg] = RP_6_month_after_3_month_high

def on_pick(event):
    legend_ = event.artist
    isVisible = legend_.get_visible()

    graphs_dict[legend_].set_visible(not isVisible)
    legend_.set_visible(not isVisible)

    fig.canvas.draw()


plt.connect('pick_event', on_pick)

ax.set_title('Time Series Data')


# plt.savefig('Plot_examples/Time_Series_Data.png', dpi = 800)


plt.show()
