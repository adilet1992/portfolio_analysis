import datetime
import numpy as np
import pandas as pd
from matplotlib import pyplot as plt
import seaborn as sns

def vintage_amount(from_month, prosdays, df):
    df2 = df
    df2 = df2[(df2.OPENDATE >= datetime.date(2016, 12, 1))]
    df_open = df2
    df_open = df_open.sort_values(by='OPENDATE_YM')
    df_open = df_open.drop_duplicates('OPENDATE_YM', keep='last')
    open_ym = df_open.OPENDATE_YM.tolist()
    months_max = max(df2.MONTHS_SINCE_OPEN)
    months = [str(i) + ' Month' for i in range(from_month, months_max+1)]
    table = pd.DataFrame()
    table['Month Issue'] = open_ym
    table['Amount Loan'] = 0
    table['Quantity'] = 0
    table = table.reindex(['Month Issue', 'Amount Loan', 'Quantity'] + months, axis=1)
    for i in range(len(table)):
        table.at[table.index[i], 'Amount Loan'] = df2[df2.OPENDATE_YM == open_ym[i]].drop_duplicates('CONTRNUM', keep='last').AMOUNTLOANKZT.sum()
        table.at[table.index[i], 'Quantity'] = len(df2[df2.OPENDATE_YM == open_ym[i]].drop_duplicates('CONTRNUM', keep='last'))
    for i in range(len(table.index)):
        for j in range(3, len(table.columns)):
            df3 = df2[(df2.MONTHS_SINCE_OPEN == int(table.columns[j].split(' ')[0]))&(df2.OPENDATE_YM == table.iloc[i]['Month Issue'])]
            if len(df3):
                table.at[table.index[i], '{} Month'.format(j + from_month - 3)] = round((df3[df3.PROSDAYS > prosdays].LOANRESTKZT.sum()/table.iloc[i]['Amount Loan'])*100, 2)
    table = table.replace(np.nan, '')
    dic = {'01':'January', '02':'February', '03':'March', '04':'April', '05':'May', '06':'June', '07':'July', '08':'August', '09':'September', '10':'October', '11':'November', '12':'December'}
    month_names = [dic[str(issued)[4:6]] + ' ' + str(issued)[0:4] for issued in table['Month Issue']]
    table['Month Issue'] = month_names
    return table

def vintage_count(from_month, prosdays, df):
    df2 = df
    df2 = df2[(df2.OPENDATE >= datetime.date(2016, 12, 1))]
    df_open = df2
    df_open = df_open.sort_values(by='OPENDATE_YM')
    df_open = df_open.drop_duplicates('OPENDATE_YM', keep='last')
    open_ym = df_open.OPENDATE_YM.tolist()
    months_max = max(df2.MONTHS_SINCE_OPEN)
    months = [str(i) + ' Month' for i in range(from_month, months_max+1)]
    table = pd.DataFrame()
    table['Month Issue'] = open_ym
    table['Amount Loan'] = 0
    table['Quantity'] = 0
    table = table.reindex(['Month Issue', 'Amount Loan', 'Quantity'] + months, axis=1)
    for i in range(len(table)):
        table.at[table.index[i], 'Amount Loan'] = df2[df2.OPENDATE_YM == open_ym[i]].drop_duplicates('CONTRNUM', keep='last').AMOUNTLOANKZT.sum()
        table.at[table.index[i], 'Quantity'] = len(df2[df2.OPENDATE_YM == open_ym[i]].drop_duplicates('CONTRNUM', keep='last'))
    for i in range(len(table.index)):
        for j in range(3, len(table.columns)):
            df3 = df2[(df2.MONTHS_SINCE_OPEN == int(table.columns[j].split(' ')[0]))&(df2.OPENDATE_YM == table.iloc[i]['Month Issue'])]
            if len(df3):
                table.at[table.index[i], '{} Month'.format(j + from_month - 3)] = len(df3[df3.PROSDAYS > prosdays])
    table = table.replace(np.nan, '')
    dic = {'01':'January', '02':'February', '03':'March', '04':'April', '05':'May', '06':'June', '07':'July', '08':'August', '09':'September', '10':'October', '11':'November', '12':'December'}
    month_names = [dic[str(issued)[4:6]] + ' ' + str(issued)[0:4] for issued in table['Month Issue']]
    table['Month Issue'] = month_names
    return table
	
def vintage_graph(from_month, df):
    a = []
    fig, ax = plt.subplots()
    plt.xticks(rotation = 90)
    plt.tick_params(labelsize=10)
    months = []
    for i in range(from_month, len(df)):
        months.append({'Months':'{} Month'.format(i)})

    d1 = pd.DataFrame(months)
    month_iss = []
    for i in range(len(df)-from_month):
        month_iss.append(df.iloc[i]['Month Issue'])

    for i in range(len(df)-from_month):
        for j in range(from_month, len(df)):
            if df.iloc[i]['{} Month'.format(j)] != '':
                a.append(df.iloc[i]['{} Month'.format(j)])
        ax = sns.pointplot(ax=ax, x=d1['Months'], y=a, scale = 0.4, color = (np.random.rand(), np.random.rand(), np.random.rand()))
        a = []
    ax.legend(handles=ax.lines[::len(df) - from_month + 1], labels=month_iss, ncol=8, loc=9, fontsize=10)
    fig.set_size_inches(20, 10)
    plt.xlabel('');
    return fig