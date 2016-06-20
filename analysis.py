import openpyxl
from openpyxl import Workbook
import csv
import os, os.path
from openpyxl.cell import get_column_letter
from datetime import datetime, timedelta, date
import re
from datetime import datetime
from dateutil import parser
import pandas as pd
import matplotlib as plt
import Tkinter
import numpy as np
import FileDialog

datapath = os.getcwd()
rowlist = []
each_row = []
values_list = []
data = []
dates = []
current36_mo = []
current48_mo = []
current60_mo = []
late_30_36_mo = []
late_30_48_mo = []
late_30_60_mo = []
late_60_36_mo = []
late_60_48_mo = []
late_60_60_mo = []
late_90np_36_mo = []
late_90np_48_mo = []
late_90np_60_mo = []
late_90p_36_mo = []
late_90p_48_mo = []
late_90p_60_mo = []
done_con = []
files=len(os.listdir(datapath))
finish_count=0

for filename in os.listdir(datapath):
    filename = os.path.join(datapath, filename)
    print "Reading:", filename
    finish_count = finish_count + 1
    m_36=0
    m_48=0
    m_60=0
    unknown=0
    total=0
    complete=0
    rowlist = []
    each_row = []
    if(filename.find('.py')>-1 or filename.find('.exe')>-1):
        pass;
    else:
        wb = openpyxl.load_workbook(filename)
        sheet = wb.get_sheet_by_name('Sheet1')
        due = sheet['A2640'].value
        for row in sheet['A2':'AK2639']:
            for item in row:
                each_row.append(item.value)
            rowlist.append(list(each_row))
            each_row[:] = []
        i=0
        j=0
        late=0
        late_recpmt_36=0
        late_recpmt_48=0
        late_recpmt_60=0
        late90_36=0
        late90_48=0
        late90_60=0
        late60_36=0
        late60_48=0
        late60_60=0
        late30_36=0
        late30_48=0
        late30_60=0
        current_36=0
        current_48=0
        current_60=0
        current = timedelta(days=30)
        late_30 = timedelta(days=60)
        late_90 = timedelta(days=90)
        u_error=0
        test_count = 0
        for client in rowlist:
            status = str(rowlist[i][4])
            if(re.search('=IF', status) is not None):
                try:
                    if(int(rowlist[i][0])<10000 and int(rowlist[i][0])>1000):
                        m_36 = m_36+1
                    elif(rowlist[i][0]<200000 and rowlist[i][0]>9999):
                        m_48 = m_48+1
                    elif(rowlist[i][0]>199999 and rowlist[i][0]<299999):
                        m_60 = m_60 + 1
                    else:
                        unknown = unknown + 1
                except ValueError:
                    pass
            elif(re.search('[a-zA-Z]', status) is not None):
                complete = complete + 1
            else:
                pass
            i=i+1
            total=m_36+m_48+m_60+complete
            status = str(rowlist[j][4])
            last_pmt = rowlist[j][8]
            duedate = rowlist[j][7]
            if(isinstance(last_pmt, unicode) is True):
                u_error = u_error+1
            elif(last_pmt is str or last_pmt is None or last_pmt is int or duedate is str or duedate is None or duedate is int):
                pass
            elif(rowlist[j][0]<10000 and rowlist[j][0]>1 and re.search('=IF', status) is not None):            #36 months
                late = due - duedate
                test_count = test_count+1
                if(late <= current):
                    current_36 = current_36 + 1
                elif(late > current and late <= late_30):
                    late30_36 = late30_36 + 1
                elif(late > late_30 and late < late_90):
                    late60_36 = late60_36 + 1
                elif(late >= late_90 and (due-last_pmt) <= current):
                    late_recpmt_36 = late_recpmt_36 + 1
                elif(late >= late_90 and (due-last_pmt) > current):
                    late90_36 = late90_36 + 1
                else:
                    print late
            elif(rowlist[j][0]<200000 and rowlist[j][0]>9999 and re.search('=IF', status) is not None):      #48 months
                duedate = rowlist[j][7]
                late = due - duedate
                if(late <= current):
                    current_48 = current_48 + 1
                elif(late > current and late <= late_30):
                    late30_48 = late30_48 + 1
                elif(late > late_30 and late < late_90):
                    late60_48 = late60_48 + 1
                elif(late >= late_90 and (due-last_pmt) <= current):
                    late_recpmt_48 = late_recpmt_48 + 1
                elif(late >= late_90 and (due-last_pmt) > current):
                    late90_48 = late90_48 + 1
                else:
                    print late
            elif(rowlist[j][0]>199999 and rowlist[j][0]<299999 and re.search('=IF', status) is not None):    #60 months
                duedate = rowlist[j][7]
                late = due - duedate
                if(late <= current):
                    current_60 = current_60 + 1
                elif(late > current and late <= late_30):
                    late30_60 = late30_60 + 1
                elif(late > late_30 and late < late_90):
                    late60_60 = late60_60 + 1
                elif(late >= late_90 and (due-last_pmt) <= current):
                    late_recpmt_60 = late_recpmt_60 + 1
                elif(late >= late_90 and (due-last_pmt) > current):
                    late90_60 = late90_60 + 1
                else:
                    print late
            else:
                pass
            j=j+1
        date_str = due.date()
# List Order: (date, 36 mo, 48 mo, 60 mo, unknown, sold/repo/etc, 90+ recent pmt 36mo, 90+ recent pmt 48mo, 90+ recent pmt 60mo, 90+ no pmt 36mo, 90+ no pmt 48mo, 90+ no pmt 60mo,
#                  60-90 36mo, 60-90 48mo, 60-90 60mo, 30-60 36mo, 30-60 48mo, 30-60 60mo, current 36mo, current 48mo, current 60mo, complete %)
        if(m_36<1):
            m_36=m_36+1
        if(m_48<1):
            m_48=m_48+1
        if(m_60<1):
            m_60=m_60+1
        info = [date_str, m_36, m_48, m_60, unknown, complete, (late_recpmt_36/float(m_36))*100, (late_recpmt_48/float(m_48))*100, (late_recpmt_60/float(m_60))*100, (late90_36/float(m_36))*100,
            (late90_48/float(m_48))*100, (late90_60/float(m_60))*100, (late60_36/float(m_36))*100, (late60_48/float(m_48))*100, (late60_60/float(m_60))*100, (late30_36/float(m_36))*100, (late30_48/float(m_48))*100,
            (late30_60/float(m_60))*100, (current_36/float(m_36))*100, (current_48/float(m_48))*100, (current_60/float(m_60))*100, (complete/float(total))*100]
        progress = (finish_count/float(files))*100
        print ("%.1f"%progress), "% Complete"
        data.append(info)
k=0
for thing in data:
    data[k][0] = str(data[k][0])
    k = k+1
data.sort(key = lambda x: datetime.strptime(x[0], '%Y-%m-%d'))
m=0
for info in data:
    dates.append(data[m][0])
    current36_mo.append(data[m][18])
    current48_mo.append(data[m][19])
    current60_mo.append(data[m][20])
    late_30_36_mo.append(data[m][15])
    late_30_48_mo.append(data[m][16])
    late_30_60_mo.append(data[m][17])
    late_60_36_mo.append(data[m][12])
    late_60_48_mo.append(data[m][13])
    late_60_60_mo.append(data[m][14])
    late_90np_36_mo.append(data[m][9])
    late_90np_48_mo.append(data[m][10])
    late_90np_60_mo.append(data[m][11])
    late_90p_36_mo.append(data[m][6])
    late_90p_48_mo.append(data[m][7])
    late_90p_60_mo.append(data[m][8])
    done_con.append(data[m][21])
    m = m+1

print "Could not process due to Unicode error: ", u_error

font = {'family': 'normal', 'weight': 'bold', 'size': 60}
plt.rc('font', **font)

dates_dr = pd.date_range(dates[0], dates[len(dates)-1], freq='M')
report36_df = pd.DataFrame({'Months': dates_dr, 'Current (< 30 Days) (36 Mo)': current36_mo, '30-60 Days Late (36 Mo)': late_30_36_mo, '60-90 Days Late (36 Mo)': late_60_36_mo, '90+ Days (No Recent Payment) (36 Mo)': late_90np_36_mo, '90+ Days (Recent Payment) (36 Mo)': late_90p_36_mo})
report36_df['Months'] = pd.to_datetime(report36_df['Months']).apply(lambda y: y.date())
report36_df = report36_df.set_index('Months')
print report36_df
ax = report36_df.plot(kind='line', rot=40, title="Contract Statuses Per Term", colormap='autumn', figsize=(75,25), linewidth=3, marker='None')
ax.tick_params(labelsize=60)
ax.set_xlabel("Months")
ax.set_ylabel("Percentage")
ax.xaxis.grid()
ax.yaxis.grid()
#fig = ax.get_figure()
#fig.savefig('36mo.png')

report48_df = pd.DataFrame({'Months': dates_dr, 'Current (< 30 Days) (48 Mo)': current48_mo, '30-60 Days Late (48 Mo)': late_30_48_mo, '60-90 Days Late (48 Mo)': late_60_48_mo, '90+ Days (No Recent Payment) (48 Mo)': late_90np_48_mo, '90+ Days (Recent Payment) (48 Mo)': late_90p_48_mo})
report48_df['Months'] = pd.to_datetime(report48_df['Months']).apply(lambda y: y.date())
report48_df = report48_df.set_index('Months')
print report48_df
#ax2 = report48_df.plot(kind='line', rot=40, title="48 Month Contracts", figsize=(20,12), color=['lime', 'yellow', 'red', 'orange', 'blue'], linewidth=3, marker='o', markersize=12)
ax_ax = report48_df.plot(ax=ax, colormap='cool', linewidth=3, marker='None', figsize=(75,25))
ax_ax.tick_params(labelsize=60)
ax_ax.set_xlabel("Months")
ax_ax.set_ylabel("Percentage")
ax_ax.xaxis.grid()
ax_ax.yaxis.grid()
#fig2 = ax_ax.get_figure()
#fig2.savefig('48mo.png')
custom_colormap = ['springgreen', 'greenyellow', 'lightgreen', 'teal', 'darkgreen']
report60_df = pd.DataFrame({'Months': dates_dr, 'Current (< 30 Days) (60 Mo)': current60_mo, '30-60 Days Late (60 Mo)': late_30_60_mo, '60-90 Days Late (60 Mo)': late_60_60_mo, '90+ Days (No Recent Payment) (60 Mo)': late_90np_60_mo, '90+ Days (Recent Payment) (60 Mo)': late_90p_60_mo})
report60_df['Months'] = pd.to_datetime(report60_df['Months']).apply(lambda y: y.date())
report60_df = report60_df.set_index('Months')
print report60_df
#ax3 = report60_df.plot(kind='line', rot=40, title="60 Month Contracts", figsize=(20,12), color=['lime', 'yellow', 'red', 'orange', 'blue'], linewidth=3, marker='o', markersize=12)
ax = ax_ax
ax3 = report60_df.plot(ax=ax, color=custom_colormap, linewidth=3, marker='None', figsize=(75,25), rot=40)
ax3.tick_params(labelsize=60)
ax3.set_xlabel("Months")
ax3.set_ylabel("Percentage")
ax3.xaxis.grid()
ax3.yaxis.grid()
#ax3.set_position([1,1,0.5,0.8])
lgd = ax3.legend(loc='center left', bbox_to_anchor=(1.0,0.5), fancybox=True, shadow=True, prop={'size':40})
fig3 = ax3.get_figure()
fig3.savefig('chart.png', bbox_extra_artists=(lgd,), bbox_inches='tight')
