# Aperture_NESTE_Relative_Performance_Task

Summary:
Starting from a spreadsheet containing the NESTE FH and SXXP Index prices for the last ten years, I calculate the stock’s relative performance (RP) 
as a time series. I then collect various statistics relating to the stock’s performance 1, 3 and 6 months after each 3-month high RP. My code outputs 
various formatted spreadsheets at each stage of the exercise to highlight a 3-month high, outperformance/underperformance, among other things. Finally, 
the code produces two plots summarising the data produced.

NB: 1_month_(out)performance = RP_1_month_after_3_month_high / RP_at_3_month_high (3 and 6 month equivilants defined similarly)
NB: I define outperformance (underperformance) as performance >= 0 (< 0)

This repository contains 6 scripts, 1 input spreadsheet 'Neste_and_SXXP price.xlsx' and an OUTPUT folder (to hold output spreadsheets).

The scripts (in the order they run) are:

- functions.py: Holds useful functions for calculating statistics and colouring excel cells based on conditions.

- Rolling_3_month_high_RP.py: Calculates the rolling 3-month RP high and the dates they occur. 
- Calculating_outperformance.py: Calculates the performance and RP 1, 3 and 6 months after each 3-month high.
- Calculating_stats.py: Calculates various statistics relating to performance (max, hit ratio etc.) and produces 3 output spreadsheets at various
stages of the calculation process.

- Plotting_time_series.py: Plots an interactive time series of RP, rolling 3-month high RP, 1, 3 and 6 month performance after a 3 month high.
- Plotting_summary_bar_chart.py: Produces a bar chart containing the averages of various performance statistics at 1, 3 and 6 months after a 3 month 
high RP.


Instructions:

1. Run Calculating_stats.py. This script will call the earlier scripts it needs. The spreadsheets this script produces are:
Bla
Bla
Bla
2. Run Plotting_time_series.py or/and Plotting_summary_bar_chart.py. Clicking on the line segments in the legend on the time series plot will hide
or reveal lines to make the plot clearer as desired. 

IMPORTANT: If the hiding line feature on the time series plot is not working try updating matplotlib on your machine. On mac, this is achieved by running
the following command within the terminal:

pip install --upgrade matplotlib

If other packages do not work, ensure all packages imported at the start of the Plotting_time_series.py script are installed on your machine.



