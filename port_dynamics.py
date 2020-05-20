from matplotlib import pyplot as plt
import seaborn as sns
import pandas as pd
import datetime

def createTable(df):
	tab = []
	df2 = df
	df2 = df2[df2.OPENDATE >= datetime.date(2016, 12, 1)]
	operdates = list(set(df2.OPERDATE1.tolist()))
	operdates.sort()
	for operdate in operdates:
		amount = df2[df2.OPERDATE1 == operdate].LOANRESTKZT.sum()
		one_plus_loan = df2[(df2.OPERDATE1 == operdate)&(df2.PROSDAYS > 0)].LOANRESTKZT.sum()
		thirty_plus_loan = df2[(df2.OPERDATE1 == operdate)&(df2.PROSDAYS > 30)].LOANRESTKZT.sum()
		ninety_plus_loan = df2[(df2.OPERDATE1 == operdate)&(df2.PROSDAYS > 90)].LOANRESTKZT.sum()
		hundredeighty_plus_loan = df2[(df2.OPERDATE1 == operdate)&(df2.PROSDAYS > 180)].LOANRESTKZT.sum()
		tab.append({'Operdate':operdate, 'Portfolio':amount, '1+DPD':round((one_plus_loan/amount)*100, 2), '30+DPD':round((thirty_plus_loan/amount)*100, 2),
				   '90+DPD':round((ninety_plus_loan/amount)*100, 2), '180+DPD':round((hundredeighty_plus_loan/amount)*100, 2)})
	df_tab = pd.DataFrame(tab)
	df_tab = df_tab.reindex(['Operdate', 'Portfolio', '1+DPD', '30+DPD', '90+DPD', '180+DPD'], axis=1)
	return df_tab
	
def viewTable(df):
	tab = []
	df2 = df
	df2 = df2[df2.OPENDATE >= datetime.date(2016, 12, 1)]
	operdates = list(set(df2.OPERDATE1.tolist()))
	operdates.sort()
	for operdate in operdates:
		amount = df2[df2.OPERDATE1 == operdate].LOANRESTKZT.sum()
		one_plus_loan = df2[(df2.OPERDATE1 == operdate)&(df2.PROSDAYS > 0)].LOANRESTKZT.sum()
		thirty_plus_loan = df2[(df2.OPERDATE1 == operdate)&(df2.PROSDAYS > 30)].LOANRESTKZT.sum()
		ninety_plus_loan = df2[(df2.OPERDATE1 == operdate)&(df2.PROSDAYS > 90)].LOANRESTKZT.sum()
		hundredeighty_plus_loan = df2[(df2.OPERDATE1 == operdate)&(df2.PROSDAYS > 180)].LOANRESTKZT.sum()
		tab.append({'Operdate':operdate, 'Portfolio':'{:,}'.format(int(amount)).replace(',', ' '), '1+DPD':str(round((one_plus_loan/amount)*100, 2))+'%', '30+DPD':str(round((thirty_plus_loan/amount)*100, 2))+'%',
				   '90+DPD':str(round((ninety_plus_loan/amount)*100, 2))+'%', '180+DPD':str(round((hundredeighty_plus_loan/amount)*100, 2))+'%'})
	df_tab = pd.DataFrame(tab)
	df_tab = df_tab.reindex(['Operdate', 'Portfolio', '1+DPD', '30+DPD', '90+DPD', '180+DPD'], axis=1)
	df_tab.Operdate = pd.to_datetime(df_tab.Operdate, format="%d%m%Y%").dt.strftime('%d.%m.%Y')
	return df_tab
	
def createGraph(df_tab):
	data = df_tab
	data.Operdate = pd.to_datetime(data.Operdate, format="%d%m%Y%").dt.strftime('%d.%m.%Y')
	f, ax = plt.subplots()
	plt.xticks(rotation=90)
	plt.tick_params(labelsize=7)
	ax = sns.barplot(x=data['Operdate'], y=data['Portfolio'], data=data, color='darkturquoise')
	ax.tick_params(axis='y')
	ax1 = ax.twinx()
	ax1=sns.pointplot(ax=ax1,x=data['Operdate'],y=data['1+DPD'],data=data,color='blue', scale=0.4)
	ax1=sns.pointplot(ax=ax1,x=data['Operdate'],y=data['30+DPD'],data=data,color='orange', scale=0.4)
	ax1=sns.pointplot(ax=ax1,x=data['Operdate'],y=data['90+DPD'],data=data,color='red', scale=0.4)
	ax1=sns.pointplot(ax=ax1,x=data['Operdate'],y=data['180+DPD'],data=data,color='black', scale=0.4)
	ax1.tick_params(axis='y')
	f.tight_layout()
	ax1.legend(handles=ax1.lines[::len(data)+1], labels=['1+DPD','30+DPD','90+DPD', '180+DPD'], ncol=4, loc=9, fontsize=7)
	f.set_size_inches(12,7)
	plt.xlabel('')
	plt.ylabel('');
	return f