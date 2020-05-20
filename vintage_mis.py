import pandas as pd
import datetime
from matplotlib import pyplot as plt
import seaborn as sns

def createTable(base):
	tab = []
	base = base[base.OPENDATE >= datetime.date(2016, 12, 1)]
	open_yms = base.drop_duplicates('OPENDATE_YM', keep='last').sort_values(by='OPENDATE_YM')['OPENDATE_YM'].tolist()
	for open_ym in open_yms:
		count = len(base[base.OPENDATE_YM == open_ym].drop_duplicates('CONTRNUM', keep='last'))
		amount = base[base.OPENDATE_YM == open_ym].drop_duplicates('CONTRNUM', keep='last').AMOUNTLOANKZT.sum()
		one_plus = base[(base.PROSDAYS > 0)&(base.OPENDATE_YM == open_ym)&(base.MONTHS_SINCE_OPEN == 1)].LOANRESTKZT.sum()
		thirty_plus = base[(base.PROSDAYS > 30)&(base.OPENDATE_YM == open_ym)&(base.MONTHS_SINCE_OPEN == 3)].LOANRESTKZT.sum()
		ninety_plus = base[(base.PROSDAYS > 90)&(base.OPENDATE_YM == open_ym)&(base.MONTHS_SINCE_OPEN == 6)].LOANRESTKZT.sum()
		tab.append({'Month of Issue':open_ym, 'Quantity of Issue':count, 'Amount Loan':amount, '1+@1MOB':round((one_plus/amount)*100, 1), 
				   '30+@3MOB':round((thirty_plus/amount)*100, 1), '90+@6MOB':round((ninety_plus/amount)*100, 1)})
	df_tab = pd.DataFrame(tab)
	df_tab = df_tab.reindex(['Month of Issue', 'Quantity of Issue', 'Amount Loan', '1+@1MOB', '30+@3MOB', '90+@6MOB'], axis=1)
	months_name = {'01':'January', '02':'February', '03':'March', '04':'April', '05':'May', '06':'June', '07':'July', '08':'August',
                          '09':'September', '10':'October', '11':'November', '12':'December'}
	df_tab['Month of Issue'] = [months_name[str(month)[4:6]] + ' ' + str(month)[0:4] for month in df_tab['Month of Issue']]
	return df_tab
	
def viewTable(base):
	tab = []
	base = base[base.OPENDATE >= datetime.date(2016, 12, 1)]
	open_yms = base.drop_duplicates('OPENDATE_YM', keep='last').sort_values(by='OPENDATE_YM')['OPENDATE_YM'].tolist()
	for open_ym in open_yms:
		count = len(base[base.OPENDATE_YM == open_ym].drop_duplicates('CONTRNUM', keep='last'))
		amount = base[base.OPENDATE_YM == open_ym].drop_duplicates('CONTRNUM', keep='last').AMOUNTLOANKZT.sum()
		one_plus = base[(base.PROSDAYS > 0)&(base.OPENDATE_YM == open_ym)&(base.MONTHS_SINCE_OPEN == 1)].LOANRESTKZT.sum()
		thirty_plus = base[(base.PROSDAYS > 30)&(base.OPENDATE_YM == open_ym)&(base.MONTHS_SINCE_OPEN == 3)].LOANRESTKZT.sum()
		ninety_plus = base[(base.PROSDAYS > 90)&(base.OPENDATE_YM == open_ym)&(base.MONTHS_SINCE_OPEN == 6)].LOANRESTKZT.sum()
		tab.append({'Month of Issue':open_ym, 'Quantity of Issue':count, 'Amount Loan':'{:,}'.format(int(amount)).replace(',', ' '), '1+@1MOB':str(round((one_plus/amount)*100, 1)) + '%', 
				   '30+@3MOB':str(round((thirty_plus/amount)*100, 1)) + '%', '90+@6MOB':str(round((ninety_plus/amount)*100, 1)) + '%'})
	df_tab = pd.DataFrame(tab)
	df_tab = df_tab.reindex(['Month of Issue', 'Quantity of Issue', 'Amount Loan', '1+@1MOB', '30+@3MOB', '90+@6MOB'], axis=1)
	months_name = {'01':'January', '02':'February', '03':'March', '04':'April', '05':'May', '06':'June', '07':'July', '08':'August',
                          '09':'September', '10':'October', '11':'November', '12':'December'}
	df_tab['Month of Issue'] = [months_name[str(month)[4:6]] + ' ' + str(month)[0:4] for month in df_tab['Month of Issue']]
	return df_tab
	
def createGraph(df_tab):
	f, ax = plt.subplots()
	plt.xticks(rotation=90)
	plt.tick_params(labelsize=7)
	ax = sns.barplot(x=df_tab['Month of Issue'], y=df_tab['Amount Loan'], data=df_tab, color='darkturquoise')
	ax.tick_params(axis='y')
	ax1 = ax.twinx()
	ax1=sns.pointplot(ax=ax1,x=df_tab['Month of Issue'],y=df_tab['1+@1MOB'],data=df_tab,color='blue', scale=0.3)
	ax1=sns.pointplot(ax=ax1,x=df_tab['Month of Issue'],y=df_tab['30+@3MOB'],data=df_tab,color='orange', scale=0.3)
	ax1=sns.pointplot(ax=ax1,x=df_tab['Month of Issue'],y=df_tab['90+@6MOB'],data=df_tab,color='red', scale=0.3)
	ax1.tick_params(axis='y')
	f.tight_layout()
	ax1.legend(handles=ax1.lines[::len(df_tab)+1], labels=["1+@1MOB","30+@3MOB","90+@6MOB"], ncol=4, loc=9, fontsize=7)
	f.set_size_inches(17,7)
	plt.xlabel('')
	plt.ylabel('')
	return f
			