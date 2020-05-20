import pandas as pd
import datetime

def createReport(base, excel_file):
	base = base[base.PRODUCT == 'Экспресс кредит']
	max_date = max(base.OPERDATE1)
	d1 = max_date
	d2 = d1 - datetime.timedelta(days=28)
	a = d2.day-1
	if a > 0:
		d2 = d2 - datetime.timedelta(days=a)
	if a == 0:
		d2 = d2
	prev_date = d2

	base_dec = base[base.OPERDATE1 == max_date]
	base_dec = base_dec[(base_dec.PROSDAYS > 0)&(base_dec.PROSDAYS <= 30)]
	con_dec = []
	for i in range(len(base_dec)):
		con_dec.append(base_dec.iloc[i]['CONTRNUM'])
	base_nov = base[base.OPERDATE1 == prev_date]
	base_nov = base_nov[(base_nov.PROSDAYS == 0)]
	base_nov = base_nov[base_nov.CONTRNUM.isin(con_dec)]
	con_nov = []
	for i in range(len(base_nov)):
		con_nov.append(base_nov.iloc[i]['CONTRNUM'])
	base_dec1 = base_dec[base_dec.CONTRNUM.isin(con_nov)]
	base_dec1['BASKET'] = '[Current] -> [1-30]'
	base_dec3 = base_dec1[['OPERDATE1', 'CONTRNUM', 'FILIAL', 'CLIENTLABEL', 'IIN', 'AMOUNTLOANKZT', 'PROSAMOUNT', 'LOANRESTKZT', 'OPENDATE', 'CLOSEDATE', 'PROSDAYS', 'PRODUCT', 'BASKET']]
	open_ym = []
	tab = []
	tab2 = []
	fils = []
	amount = 0
	loan = 0
	base_dec3_1_30 = base_dec3

	base_dec = base[base.OPERDATE1 == max_date]
	base_dec = base_dec[(base_dec.PROSDAYS > 30)&(base_dec.PROSDAYS <= 60)]
	con_dec = []
	for i in range(len(base_dec)):
		con_dec.append(base_dec.iloc[i]['CONTRNUM'])
	base_nov = base[base.OPERDATE1 == prev_date]
	base_nov = base_nov[(base_nov.PROSDAYS > 0)&(base_nov.PROSDAYS <= 30)]
	base_nov = base_nov[base_nov.CONTRNUM.isin(con_dec)]
	con_nov = []
	for i in range(len(base_nov)):
		con_nov.append(base_nov.iloc[i]['CONTRNUM'])
	base_dec1 = base_dec[base_dec.CONTRNUM.isin(con_nov)]
	base_dec1['BASKET'] = '[1-30] -> [31-60]'
	base_dec3 = base_dec1[['OPERDATE1', 'CONTRNUM', 'FILIAL', 'CLIENTLABEL', 'IIN', 'AMOUNTLOANKZT', 'PROSAMOUNT', 'LOANRESTKZT', 'OPENDATE', 'CLOSEDATE', 'PROSDAYS', 'PRODUCT', 'BASKET']]
	open_ym = []
	tab = []
	tab2 = []
	fils = []
	amount = 0
	loan = 0
	base_dec3_31_60 = base_dec3

	base_dec = base[base.OPERDATE1 == max_date]
	base_dec = base_dec[(base_dec.PROSDAYS > 60)&(base_dec.PROSDAYS <= 90)]
	con_dec = []
	for i in range(len(base_dec)):
		con_dec.append(base_dec.iloc[i]['CONTRNUM'])
	base_nov = base[base.OPERDATE1 == prev_date]
	base_nov = base_nov[(base_nov.PROSDAYS > 30)&(base_nov.PROSDAYS <= 60)]
	base_nov = base_nov[base_nov.CONTRNUM.isin(con_dec)]
	con_nov = []
	for i in range(len(base_nov)):
		con_nov.append(base_nov.iloc[i]['CONTRNUM'])
	base_dec1 = base_dec[base_dec.CONTRNUM.isin(con_nov)]
	base_dec1['BASKET'] = '[31-60] -> [61-90]'
	base_dec3 = base_dec1[['OPERDATE1', 'CONTRNUM', 'FILIAL', 'CLIENTLABEL', 'IIN', 'AMOUNTLOANKZT', 'PROSAMOUNT', 'LOANRESTKZT', 'OPENDATE', 'CLOSEDATE', 'PROSDAYS', 'PRODUCT', 'BASKET']]
	open_ym = []
	tab = []
	tab2 = []
	fils = []
	amount = 0
	loan = 0
	base_dec3_61_90 = base_dec3

	base_dec = base[base.OPERDATE1 == max_date]
	base_dec = base_dec[(base_dec.PROSDAYS > 90)&(base_dec.PROSDAYS <= 120)]
	con_dec = []
	for i in range(len(base_dec)):
		con_dec.append(base_dec.iloc[i]['CONTRNUM'])
	base_nov = base[base.OPERDATE1 == prev_date]
	base_nov = base_nov[(base_nov.PROSDAYS > 60)&(base_nov.PROSDAYS <= 90)]
	base_nov = base_nov[base_nov.CONTRNUM.isin(con_dec)]
	con_nov = []
	for i in range(len(base_nov)):
		con_nov.append(base_nov.iloc[i]['CONTRNUM'])
	base_dec1 = base_dec[base_dec.CONTRNUM.isin(con_nov)]
	base_dec1['BASKET'] = '[61-90] -> [91-120]'
	base_dec3 = base_dec1[['OPERDATE1', 'CONTRNUM', 'FILIAL', 'CLIENTLABEL', 'IIN', 'AMOUNTLOANKZT', 'PROSAMOUNT', 'LOANRESTKZT', 'OPENDATE', 'CLOSEDATE', 'PROSDAYS', 'PRODUCT', 'BASKET']]
	open_ym = []
	tab = []
	tab2 = []
	fils = []
	amount = 0
	loan = 0
	base_dec3_91_120 = base_dec3

	base_all = pd.concat([base_dec3_1_30, base_dec3_31_60, base_dec3_61_90, base_dec3_91_120], axis=0)
	base_all.OPERDATE1 = pd.to_datetime(base_all.OPERDATE1, format="%d%m%Y%").dt.strftime('%d.%m.%Y')
	base_all.OPENDATE = pd.to_datetime(base_all.OPENDATE, format="%d%m%Y%").dt.strftime('%d.%m.%Y')
	base_all.CLOSEDATE = pd.to_datetime(base_all.CLOSEDATE, format="%d%m%Y%").dt.strftime('%d.%m.%Y')
	base_all.to_excel(excel_file, index=False)