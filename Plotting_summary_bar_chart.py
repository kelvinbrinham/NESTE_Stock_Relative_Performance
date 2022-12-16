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

from Calculating_stats import Neste_and_SXXP_summary_df
from Rolling_3_month_high_RP import start_date
from Rolling_3_month_high_RP import end_date

Observations_lst = list(Neste_and_SXXP_summary_df['Observations'])
Avg_performance_lst = list(Neste_and_SXXP_summary_df['Average Performance'])
Median_performance_lst = list(Neste_and_SXXP_summary_df['Median Performance'])
Hit_Ratio_lst = list(Neste_and_SXXP_summary_df['Hit Ratio'])
Max_performance_lst = list(Neste_and_SXXP_summary_df['Maximum Performance'])
Min_performance_lst = list(Neste_and_SXXP_summary_df['Minimum Performance'])

Data_lst = [Observations_lst, Avg_performance_lst, Median_performance_lst,
Hit_Ratio_lst, Max_performance_lst, Min_performance_lst]

#<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
#Bar chart of final observations
#Manipulate data for bar bar chart
one_month_list = []
three_month_list = []
six_month_list = []

for list_ in Data_lst:
    one_month_list.append(list_[0])
    three_month_list.append(list_[1])
    six_month_list.append(list_[2])

one_month_list = one_month_list[1:]
three_month_list = three_month_list[1:]
six_month_list = six_month_list[1:]

round_to_decimal_places = 1
one_month_list_perc = [round(100*x, round_to_decimal_places) for x in one_month_list]
three_month_list_perc = [round(100*x, round_to_decimal_places) for x in three_month_list]
six_month_list_perc = [round(100*x, round_to_decimal_places) for x in six_month_list]

labels = Neste_and_SXXP_summary_df.columns.values[1:]


for i in range(len(labels)):
    dummy_list = labels
    dummy_list[i] = dummy_list[i].replace(' ', '\n')

labels = dummy_list

x = np.arange(len(labels))  # the label locations

width = 0.3  # the width of the bars

fig, ax = plt.subplots()
fig.set_size_inches(7, 6)
rects1 = ax.bar(x - width, one_month_list_perc, width, label='1 Month')
rects2 = ax.bar(x, three_month_list_perc, width, label='3 Months')
rects3 = ax.bar(x + width, six_month_list_perc, width, label='6 Months', color = 'purple')

comment1 = str(start_date.date()) + ' - ' + str(end_date.date()) + '\n1 Month Observations = ' + str(259) + '\n3 Month Observations = ' + str(259) + '\n6 Month Observations = ' + str(247)
ax.annotate(comment1, fontsize = 10, xy=(0.05, 0.07), xycoords='axes fraction', bbox = dict(facecolor = 'white', edgecolor = 'black', alpha = 1))

# comment2 = str(start_date) + ' - ' + str(end_date)

# Add some text for labels, title and custom x-axis tick labels, etc.
ax.set_ylabel('%')
ax.set_title('Summary Data')
ax.set_xticks(x, labels)
ax.legend()

ax.bar_label(rects1, padding=3)
ax.bar_label(rects2, padding=3)
ax.bar_label(rects3, padding=3)

fig.tight_layout()

# plt.savefig('Plot_examples/Summary_Data_Bar_Chart.png', dpi = 800)

plt.show()


#<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
