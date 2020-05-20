import datetime
import pandas as pd
import xlsxwriter

def createReport(all_rp, excel_file):
	base = all_rp
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
	base_dec3 = base_dec1[['OPERDATE1', 'CONTRNUM', 'FILIAL', 'CLIENTLABEL', 'IIN', 'OPENDATE', 'AMOUNTLOANKZT', 'LOANRESTKZT', 'PROSDAYS', 'PRODUCT']]
	open_ym = []
	tab = []
	tab2 = []
	fils = []
	amount = 0
	loan = 0

	base_dec2 = base_dec1
	base_dec2 = base_dec2.sort_values(by = 'OPENDATE_YM')
	base_dec2 = base_dec2.drop_duplicates('OPENDATE_YM', keep = 'first')

	base_dec4 = base_dec1
	base_dec4 = base_dec4.drop_duplicates('FILIAL', keep = 'first')
	fils = base_dec4.FILIAL.tolist()

	all_loan = base_dec3.LOANRESTKZT.sum()
	all_count = len(base_dec3)

	for i in range(len(base_dec2)):
		open_ym.append(base_dec2.iloc[i]['OPENDATE_YM'])
	for i in range(len(open_ym)):
		open_ym2 = base_dec1[base_dec1.OPENDATE_YM == open_ym[i]]
		loan  = open_ym2['LOANRESTKZT'].sum()
		amount = open_ym2['AMOUNTLOANKZT'].sum()
		if str(open_ym[i])[4:6] == '01':
			month = 'Январь'
		if str(open_ym[i])[4:6] == '02':
			month = 'Февраль'
		if str(open_ym[i])[4:6] == '03':
			month = 'Март'
		if str(open_ym[i])[4:6] == '04':
			month = 'Апрель'
		if str(open_ym[i])[4:6] == '05':
			month = 'Май'
		if str(open_ym[i])[4:6] == '06':
			month = 'Июнь'
		if str(open_ym[i])[4:6] == '07':
			month = 'Июль'
		if str(open_ym[i])[4:6] == '08':
			month = 'Август'
		if str(open_ym[i])[4:6] == '09':
			month = 'Сентябрь'
		if str(open_ym[i])[4:6] == '10':
			month = 'Октябрь'
		if str(open_ym[i])[4:6] == '11':
			month = 'Ноябрь'
		if str(open_ym[i])[4:6] == '12':
			month = 'Декабрь'
		tab.append({'Месяц выдачи': month + ' ' + str(open_ym[i])[0:4], 'Сумма ОД':loan, 'Сумма займа':amount, 'Кол-во':len(open_ym2), 
					'Доля по ОД':loan/all_loan, 'Доля по кол-ву':len(open_ym2)/all_count})
		loan = 0
		amount = 0
		open_ym2.iloc[0:0]

	for i in range(len(fils)):
		fils2 = base_dec1[base_dec1.FILIAL == fils[i]]
		loan = fils2['LOANRESTKZT'].sum()
		amount = fils2['AMOUNTLOANKZT'].sum()
		tab2.append({'Филиал':fils[i], 'Сумма займа':amount, 'Сумма ОД':loan, 'Кол-во':len(fils2),
					'Доля по ОД':loan/all_loan, 'Доля по кол-ву':len(fils2)/all_count})
		loan = 0
		amount = 0
		fils2.iloc[0:0]

	base_dec3.columns = ['Дата отчета', 'Договор', 'Филиал', 'ФИО', 'ИИН', 'Дата выдачи', 'Сумма выдачи', 'ОД', 'Просрочка', 'Продукт']
	df_month = pd.DataFrame(data=tab)
	df_month = df_month.reindex(['Месяц выдачи', 'Сумма займа', 'Сумма ОД', 'Доля по ОД', 'Кол-во', 'Доля по кол-ву'], axis=1)
	df_fil = pd.DataFrame(data=tab2)
	df_fil = df_fil.reindex(['Филиал', 'Сумма займа', 'Сумма ОД', 'Доля по ОД', 'Кол-во', 'Доля по кол-ву'], axis=1)


	wb = xlsxwriter.Workbook(excel_file)

	ws = wb.add_worksheet('Current -> 1-30')
	columns1 = ['Дата отчета', 'Договор', 'Филиал', 'ФИО', 'ИИН', 'Дата выдачи', 'Сумма выдачи', 'ОД', 'Просрочка', 'Продукт']
	columns2 = ['Месяц выдачи', 'Сумма займа', 'Сумма ОД', 'Доля по ОД', 'Кол-во', 'Доля по кол-ву']
	columns3 = ['Филиал', 'Сумма займа', 'Сумма ОД', 'Доля по ОД', 'Кол-во', 'Доля по кол-ву']

	df_month_top3 = []
	df_fil_top3 = []

	df_month_sort = df_month
	df_month_sort = df_month_sort.sort_values(by = 'Доля по ОД', ascending=False)
	for i in range(3):        
		df_month_top3.append(df_month_sort.iloc[i]['Месяц выдачи'])
		
	df_fil_sort = df_fil
	df_fil_sort = df_fil_sort.sort_values(by = 'Доля по ОД', ascending=False)
	for i in range(3):
		df_fil_top3.append(df_fil_sort.iloc[i]['Филиал'])

	date_form = wb.add_format()
	date_form.set_num_format('dd.mm.yyyy')
	date_form.set_border()

	num_form = wb.add_format()
	num_form.set_num_format('###')
	num_form.set_border()

	num_form2 = wb.add_format()
	num_form2.set_border()
	num_form2.set_num_format('###,#.00')

	bold_form = wb.add_format()
	bold_form.set_bold()
	bold_form.set_align('center')
	bold_form.set_border()

	percent_form = wb.add_format()
	percent_form.set_num_format('0.00%')
	percent_form.set_border()

	simple_form = wb.add_format()
	simple_form.set_border()
		  
	date_form_red = wb.add_format()
	date_form_red.set_num_format('dd.mm.yyyy')
	date_form_red.set_border()
	date_form_red.set_font_color('red')
	date_form_red.set_bold()

	num_form_red = wb.add_format()
	num_form_red.set_num_format('###')
	num_form_red.set_border()
	num_form_red.set_font_color('red')
	num_form_red.set_bold()

	num_form2_red = wb.add_format()
	num_form2_red.set_border()
	num_form2_red.set_num_format('###,#.00')
	num_form2_red.set_font_color('red')
	num_form2_red.set_bold()

	bold_form = wb.add_format()
	bold_form.set_bold()
	bold_form.set_align('center')
	bold_form.set_border()

	percent_form_red = wb.add_format()
	percent_form_red.set_num_format('0.00%')
	percent_form_red.set_border()
	percent_form_red.set_font_color('red')
	percent_form_red.set_bold()

	simple_form_red = wb.add_format()
	simple_form_red.set_border()
	simple_form_red.set_bold()
	simple_form_red.set_font_color('red')

	date_form_yel = wb.add_format()
	date_form_yel.set_num_format('dd.mm.yyyy')
	date_form_yel.set_border()
	date_form_yel.set_bg_color('yellow')

	num_form_yel = wb.add_format()
	num_form_yel.set_num_format('###')
	num_form_yel.set_border()
	num_form_yel.set_bg_color('yellow')

	num_form2_yel = wb.add_format()
	num_form2_yel.set_border()
	num_form2_yel.set_num_format('###,#.00')
	num_form2_yel.set_bg_color('yellow')

	bold_form = wb.add_format()
	bold_form.set_bold()
	bold_form.set_align('center')
	bold_form.set_border()

	percent_form_yel = wb.add_format()
	percent_form_yel.set_num_format('0.00%')
	percent_form_yel.set_border()
	percent_form_yel.set_bg_color('yellow')

	simple_form_yel = wb.add_format()
	simple_form_yel.set_border()
	simple_form_yel.set_bg_color('yellow')


	for i in range(len(columns1)):
		ws.write(0, i, columns1[i], bold_form)
		
	for i in range(len(columns2)):
		ws.write(0, i+11, columns2[i], bold_form)
		
	for i in range(len(columns3)):
		ws.write(0, i+18, columns3[i], bold_form)
		
	for i in range(len(base_dec3)):
		if base_dec3.iloc[i]['Дата выдачи'].date() >= datetime.date(2018, 7, 1):
			for j in range(len(base_dec3.columns)):
				if (j == 0) | (j == 5):
					ws.write(i+1, j, base_dec3.iloc[i][j], date_form_yel)
				elif (j == 4):
					ws.write(i+1, j, base_dec3.iloc[i][j], num_form_yel)
				elif (j == 6) | (j == 7):
					ws.write(i+1, j, base_dec3.iloc[i][j], num_form2_yel)
				else:
					ws.write(i+1, j, base_dec3.iloc[i][j], simple_form_yel)
		else:
			for j in range(len(base_dec3.columns)):
				if (j == 0) | (j == 5):
					ws.write(i+1, j, base_dec3.iloc[i][j], date_form)
				elif (j == 4):
					ws.write(i+1, j, base_dec3.iloc[i][j], num_form)
				elif (j == 6) | (j == 7):
					ws.write(i+1, j, base_dec3.iloc[i][j], num_form2)
				else:
					ws.write(i+1, j, base_dec3.iloc[i][j], simple_form)
				
	for i in range(len(df_month)):
		if df_month.iloc[i]['Месяц выдачи'] in df_month_top3:
			for j in range(len(df_month.columns)):
				if (j == 1) | (j == 2):
					ws.write(i+1, j+11, df_month.iloc[i][j], num_form2_red)
				elif (j == 3) | (j == 5):
					ws.write(i+1, j+11, df_month.iloc[i][j], percent_form_red)
				else:
					ws.write(i+1, j+11, df_month.iloc[i][j], simple_form_red)
		else:
			for j in range(len(df_month.columns)):
				if (j == 1) | (j == 2):
					ws.write(i+1, j+11, df_month.iloc[i][j], num_form2)
				elif (j == 3) | (j == 5):
					ws.write(i+1, j+11, df_month.iloc[i][j], percent_form)
				else:
					ws.write(i+1, j+11, df_month.iloc[i][j], simple_form)

	for i in range(len(df_fil)):
		if df_fil.iloc[i]['Филиал'] in df_fil_top3:
			for j in range(len(df_fil.columns)):
				if (j == 1) | (j == 2):
					ws.write(i+1, j+18, df_fil.iloc[i][j], num_form2_red)
				elif (j == 3) | (j == 5):
					ws.write(i+1, j+18, df_fil.iloc[i][j], percent_form_red)
				else:
					ws.write(i+1, j+18, df_fil.iloc[i][j], simple_form_red)
		else:
			for j in range(len(df_fil.columns)):
				if (j == 1) | (j == 2):
					ws.write(i+1, j+18, df_fil.iloc[i][j], num_form2)
				elif (j == 3) | (j == 5):
					ws.write(i+1, j+18, df_fil.iloc[i][j], percent_form)
				else:
					ws.write(i+1, j+18, df_fil.iloc[i][j], simple_form)
				
	######################################   [1-30] -> [31-60]   ########################################### 
	base = all_rp
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
	base_dec3 = base_dec1[['OPERDATE1', 'CONTRNUM', 'FILIAL', 'CLIENTLABEL', 'IIN', 'OPENDATE', 'AMOUNTLOANKZT', 'LOANRESTKZT', 'PROSDAYS', 'PRODUCT']]
	open_ym = []
	tab = []
	tab2 = []
	fils = []
	amount = 0
	loan = 0

	base_dec2 = base_dec1
	base_dec2 = base_dec2.sort_values(by = 'OPENDATE_YM')
	base_dec2 = base_dec2.drop_duplicates('OPENDATE_YM', keep = 'first')

	base_dec4 = base_dec1
	base_dec4 = base_dec4.drop_duplicates('FILIAL', keep = 'first')
	fils = base_dec4.FILIAL.tolist()

	all_loan = base_dec3.LOANRESTKZT.sum()
	all_count = len(base_dec3)

	for i in range(len(base_dec2)):
		open_ym.append(base_dec2.iloc[i]['OPENDATE_YM'])
	for i in range(len(open_ym)):
		open_ym2 = base_dec1[base_dec1.OPENDATE_YM == open_ym[i]]
		loan  = open_ym2['LOANRESTKZT'].sum()
		amount = open_ym2['AMOUNTLOANKZT'].sum()
		if str(open_ym[i])[4:6] == '01':
			month = 'Январь'
		if str(open_ym[i])[4:6] == '02':
			month = 'Февраль'
		if str(open_ym[i])[4:6] == '03':
			month = 'Март'
		if str(open_ym[i])[4:6] == '04':
			month = 'Апрель'
		if str(open_ym[i])[4:6] == '05':
			month = 'Май'
		if str(open_ym[i])[4:6] == '06':
			month = 'Июнь'
		if str(open_ym[i])[4:6] == '07':
			month = 'Июль'
		if str(open_ym[i])[4:6] == '08':
			month = 'Август'
		if str(open_ym[i])[4:6] == '09':
			month = 'Сентябрь'
		if str(open_ym[i])[4:6] == '10':
			month = 'Октябрь'
		if str(open_ym[i])[4:6] == '11':
			month = 'Ноябрь'
		if str(open_ym[i])[4:6] == '12':
			month = 'Декабрь'
		tab.append({'Месяц выдачи': month + ' ' + str(open_ym[i])[0:4], 'Сумма ОД':loan, 'Сумма займа':amount, 'Кол-во':len(open_ym2), 
					'Доля по ОД':loan/all_loan, 'Доля по кол-ву':len(open_ym2)/all_count})
		loan = 0
		amount = 0
		open_ym2.iloc[0:0]

	for i in range(len(fils)):
		fils2 = base_dec1[base_dec1.FILIAL == fils[i]]
		loan = fils2['LOANRESTKZT'].sum()
		amount = fils2['AMOUNTLOANKZT'].sum()
		tab2.append({'Филиал':fils[i], 'Сумма займа':amount, 'Сумма ОД':loan, 'Кол-во':len(fils2),
					'Доля по ОД':loan/all_loan, 'Доля по кол-ву':len(fils2)/all_count})
		loan = 0
		amount = 0
		fils2.iloc[0:0]

	base_dec3.columns = ['Дата отчета', 'Договор', 'Филиал', 'ФИО', 'ИИН', 'Дата выдачи', 'Сумма выдачи', 'ОД', 'Просрочка', 'Продукт']
	df_month = pd.DataFrame(data=tab)
	df_month = df_month.reindex(['Месяц выдачи', 'Сумма займа', 'Сумма ОД', 'Доля по ОД', 'Кол-во', 'Доля по кол-ву'], axis=1)
	df_fil = pd.DataFrame(data=tab2)
	df_fil = df_fil.reindex(['Филиал', 'Сумма займа', 'Сумма ОД', 'Доля по ОД', 'Кол-во', 'Доля по кол-ву'], axis=1)


	ws = wb.add_worksheet('1-30 -> 31-60')
	columns1 = ['Дата отчета', 'Договор', 'Филиал', 'ФИО', 'ИИН', 'Дата выдачи', 'Сумма выдачи', 'ОД', 'Просрочка', 'Продукт']
	columns2 = ['Месяц выдачи', 'Сумма займа', 'Сумма ОД', 'Доля по ОД', 'Кол-во', 'Доля по кол-ву']
	columns3 = ['Филиал', 'Сумма займа', 'Сумма ОД', 'Доля по ОД', 'Кол-во', 'Доля по кол-ву']

	df_month_top3 = []
	df_fil_top3 = []

	df_month_sort = df_month
	df_month_sort = df_month_sort.sort_values(by = 'Доля по ОД', ascending=False)
	for i in range(3):        
		df_month_top3.append(df_month_sort.iloc[i]['Месяц выдачи'])
		
	df_fil_sort = df_fil
	df_fil_sort = df_fil_sort.sort_values(by = 'Доля по ОД', ascending=False)
	for i in range(3):
		df_fil_top3.append(df_fil_sort.iloc[i]['Филиал'])

	date_form = wb.add_format()
	date_form.set_num_format('dd.mm.yyyy')
	date_form.set_border()

	num_form = wb.add_format()
	num_form.set_num_format('###')
	num_form.set_border()

	num_form2 = wb.add_format()
	num_form2.set_border()
	num_form2.set_num_format('###,#.00')

	bold_form = wb.add_format()
	bold_form.set_bold()
	bold_form.set_align('center')
	bold_form.set_border()

	percent_form = wb.add_format()
	percent_form.set_num_format('0.00%')
	percent_form.set_border()

	simple_form = wb.add_format()
	simple_form.set_border()
		  
	date_form_red = wb.add_format()
	date_form_red.set_num_format('dd.mm.yyyy')
	date_form_red.set_border()
	date_form_red.set_font_color('red')
	date_form_red.set_bold()

	num_form_red = wb.add_format()
	num_form_red.set_num_format('###')
	num_form_red.set_border()
	num_form_red.set_font_color('red')
	num_form_red.set_bold()

	num_form2_red = wb.add_format()
	num_form2_red.set_border()
	num_form2_red.set_num_format('###,#.00')
	num_form2_red.set_font_color('red')
	num_form2_red.set_bold()

	bold_form = wb.add_format()
	bold_form.set_bold()
	bold_form.set_align('center')
	bold_form.set_border()

	percent_form_red = wb.add_format()
	percent_form_red.set_num_format('0.00%')
	percent_form_red.set_border()
	percent_form_red.set_font_color('red')
	percent_form_red.set_bold()

	simple_form_red = wb.add_format()
	simple_form_red.set_border()
	simple_form_red.set_bold()
	simple_form_red.set_font_color('red')

	date_form_yel = wb.add_format()
	date_form_yel.set_num_format('dd.mm.yyyy')
	date_form_yel.set_border()
	date_form_yel.set_bg_color('yellow')

	num_form_yel = wb.add_format()
	num_form_yel.set_num_format('###')
	num_form_yel.set_border()
	num_form_yel.set_bg_color('yellow')

	num_form2_yel = wb.add_format()
	num_form2_yel.set_border()
	num_form2_yel.set_num_format('###,#.00')
	num_form2_yel.set_bg_color('yellow')

	bold_form = wb.add_format()
	bold_form.set_bold()
	bold_form.set_align('center')
	bold_form.set_border()

	percent_form_yel = wb.add_format()
	percent_form_yel.set_num_format('0.00%')
	percent_form_yel.set_border()
	percent_form_yel.set_bg_color('yellow')

	simple_form_yel = wb.add_format()
	simple_form_yel.set_border()
	simple_form_yel.set_bg_color('yellow')


	for i in range(len(columns1)):
		ws.write(0, i, columns1[i], bold_form)
		
	for i in range(len(columns2)):
		ws.write(0, i+11, columns2[i], bold_form)
		
	for i in range(len(columns3)):
		ws.write(0, i+18, columns3[i], bold_form)
		
	for i in range(len(base_dec3)):
		if base_dec3.iloc[i]['Дата выдачи'].date() >= datetime.date(2018, 7, 1):
			for j in range(len(base_dec3.columns)):
				if (j == 0) | (j == 5):
					ws.write(i+1, j, base_dec3.iloc[i][j], date_form_yel)
				elif (j == 4):
					ws.write(i+1, j, base_dec3.iloc[i][j], num_form_yel)
				elif (j == 6) | (j == 7):
					ws.write(i+1, j, base_dec3.iloc[i][j], num_form2_yel)
				else:
					ws.write(i+1, j, base_dec3.iloc[i][j], simple_form_yel)
		else:
			for j in range(len(base_dec3.columns)):
				if (j == 0) | (j == 5):
					ws.write(i+1, j, base_dec3.iloc[i][j], date_form)
				elif (j == 4):
					ws.write(i+1, j, base_dec3.iloc[i][j], num_form)
				elif (j == 6) | (j == 7):
					ws.write(i+1, j, base_dec3.iloc[i][j], num_form2)
				else:
					ws.write(i+1, j, base_dec3.iloc[i][j], simple_form)
				
	for i in range(len(df_month)):
		if df_month.iloc[i]['Месяц выдачи'] in df_month_top3:
			for j in range(len(df_month.columns)):
				if (j == 1) | (j == 2):
					ws.write(i+1, j+11, df_month.iloc[i][j], num_form2_red)
				elif (j == 3) | (j == 5):
					ws.write(i+1, j+11, df_month.iloc[i][j], percent_form_red)
				else:
					ws.write(i+1, j+11, df_month.iloc[i][j], simple_form_red)
		else:
			for j in range(len(df_month.columns)):
				if (j == 1) | (j == 2):
					ws.write(i+1, j+11, df_month.iloc[i][j], num_form2)
				elif (j == 3) | (j == 5):
					ws.write(i+1, j+11, df_month.iloc[i][j], percent_form)
				else:
					ws.write(i+1, j+11, df_month.iloc[i][j], simple_form)

	for i in range(len(df_fil)):
		if df_fil.iloc[i]['Филиал'] in df_fil_top3:
			for j in range(len(df_fil.columns)):
				if (j == 1) | (j == 2):
					ws.write(i+1, j+18, df_fil.iloc[i][j], num_form2_red)
				elif (j == 3) | (j == 5):
					ws.write(i+1, j+18, df_fil.iloc[i][j], percent_form_red)
				else:
					ws.write(i+1, j+18, df_fil.iloc[i][j], simple_form_red)
		else:
			for j in range(len(df_fil.columns)):
				if (j == 1) | (j == 2):
					ws.write(i+1, j+18, df_fil.iloc[i][j], num_form2)
				elif (j == 3) | (j == 5):
					ws.write(i+1, j+18, df_fil.iloc[i][j], percent_form)
				else:
					ws.write(i+1, j+18, df_fil.iloc[i][j], simple_form)

	######################################   [31-60] -> [61-90]   ########################################### 
	base = all_rp
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
	base_dec3 = base_dec1[['OPERDATE1', 'CONTRNUM', 'FILIAL', 'CLIENTLABEL', 'IIN', 'OPENDATE', 'AMOUNTLOANKZT', 'LOANRESTKZT', 'PROSDAYS', 'PRODUCT']]
	open_ym = []
	tab = []
	tab2 = []
	fils = []
	amount = 0
	loan = 0

	base_dec2 = base_dec1
	base_dec2 = base_dec2.sort_values(by = 'OPENDATE_YM')
	base_dec2 = base_dec2.drop_duplicates('OPENDATE_YM', keep = 'first')

	base_dec4 = base_dec1
	base_dec4 = base_dec4.drop_duplicates('FILIAL', keep = 'first')
	fils = base_dec4.FILIAL.tolist()

	all_loan = base_dec3.LOANRESTKZT.sum()
	all_count = len(base_dec3)

	for i in range(len(base_dec2)):
		open_ym.append(base_dec2.iloc[i]['OPENDATE_YM'])
	for i in range(len(open_ym)):
		open_ym2 = base_dec1[base_dec1.OPENDATE_YM == open_ym[i]]
		loan  = open_ym2['LOANRESTKZT'].sum()
		amount = open_ym2['AMOUNTLOANKZT'].sum()
		if str(open_ym[i])[4:6] == '01':
			month = 'Январь'
		if str(open_ym[i])[4:6] == '02':
			month = 'Февраль'
		if str(open_ym[i])[4:6] == '03':
			month = 'Март'
		if str(open_ym[i])[4:6] == '04':
			month = 'Апрель'
		if str(open_ym[i])[4:6] == '05':
			month = 'Май'
		if str(open_ym[i])[4:6] == '06':
			month = 'Июнь'
		if str(open_ym[i])[4:6] == '07':
			month = 'Июль'
		if str(open_ym[i])[4:6] == '08':
			month = 'Август'
		if str(open_ym[i])[4:6] == '09':
			month = 'Сентябрь'
		if str(open_ym[i])[4:6] == '10':
			month = 'Октябрь'
		if str(open_ym[i])[4:6] == '11':
			month = 'Ноябрь'
		if str(open_ym[i])[4:6] == '12':
			month = 'Декабрь'
		tab.append({'Месяц выдачи': month + ' ' + str(open_ym[i])[0:4], 'Сумма ОД':loan, 'Сумма займа':amount, 'Кол-во':len(open_ym2), 
					'Доля по ОД':loan/all_loan, 'Доля по кол-ву':len(open_ym2)/all_count})
		loan = 0
		amount = 0
		open_ym2.iloc[0:0]

	for i in range(len(fils)):
		fils2 = base_dec1[base_dec1.FILIAL == fils[i]]
		loan = fils2['LOANRESTKZT'].sum()
		amount = fils2['AMOUNTLOANKZT'].sum()
		tab2.append({'Филиал':fils[i], 'Сумма займа':amount, 'Сумма ОД':loan, 'Кол-во':len(fils2),
					'Доля по ОД':loan/all_loan, 'Доля по кол-ву':len(fils2)/all_count})
		loan = 0
		amount = 0
		fils2.iloc[0:0]

	base_dec3.columns = ['Дата отчета', 'Договор', 'Филиал', 'ФИО', 'ИИН', 'Дата выдачи', 'Сумма выдачи', 'ОД', 'Просрочка', 'Продукт']
	df_month = pd.DataFrame(data=tab)
	df_month = df_month.reindex(['Месяц выдачи', 'Сумма займа', 'Сумма ОД', 'Доля по ОД', 'Кол-во', 'Доля по кол-ву'], axis=1)
	df_fil = pd.DataFrame(data=tab2)
	df_fil = df_fil.reindex(['Филиал', 'Сумма займа', 'Сумма ОД', 'Доля по ОД', 'Кол-во', 'Доля по кол-ву'], axis=1)


	ws = wb.add_worksheet('31-60 -> 61-90')
	columns1 = ['Дата отчета', 'Договор', 'Филиал', 'ФИО', 'ИИН', 'Дата выдачи', 'Сумма выдачи', 'ОД', 'Просрочка', 'Продукт']
	columns2 = ['Месяц выдачи', 'Сумма займа', 'Сумма ОД', 'Доля по ОД', 'Кол-во', 'Доля по кол-ву']
	columns3 = ['Филиал', 'Сумма займа', 'Сумма ОД', 'Доля по ОД', 'Кол-во', 'Доля по кол-ву']

	df_month_top3 = []
	df_fil_top3 = []

	df_month_sort = df_month
	df_month_sort = df_month_sort.sort_values(by = 'Доля по ОД', ascending=False)
	for i in range(3):        
		df_month_top3.append(df_month_sort.iloc[i]['Месяц выдачи'])
		
	df_fil_sort = df_fil
	df_fil_sort = df_fil_sort.sort_values(by = 'Доля по ОД', ascending=False)
	for i in range(3):
		df_fil_top3.append(df_fil_sort.iloc[i]['Филиал'])

	date_form = wb.add_format()
	date_form.set_num_format('dd.mm.yyyy')
	date_form.set_border()

	num_form = wb.add_format()
	num_form.set_num_format('###')
	num_form.set_border()

	num_form2 = wb.add_format()
	num_form2.set_border()
	num_form2.set_num_format('###,#.00')

	bold_form = wb.add_format()
	bold_form.set_bold()
	bold_form.set_align('center')
	bold_form.set_border()

	percent_form = wb.add_format()
	percent_form.set_num_format('0.00%')
	percent_form.set_border()

	simple_form = wb.add_format()
	simple_form.set_border()
		  
	date_form_red = wb.add_format()
	date_form_red.set_num_format('dd.mm.yyyy')
	date_form_red.set_border()
	date_form_red.set_font_color('red')
	date_form_red.set_bold()

	num_form_red = wb.add_format()
	num_form_red.set_num_format('###')
	num_form_red.set_border()
	num_form_red.set_font_color('red')
	num_form_red.set_bold()

	num_form2_red = wb.add_format()
	num_form2_red.set_border()
	num_form2_red.set_num_format('###,#.00')
	num_form2_red.set_font_color('red')
	num_form2_red.set_bold()

	bold_form = wb.add_format()
	bold_form.set_bold()
	bold_form.set_align('center')
	bold_form.set_border()

	percent_form_red = wb.add_format()
	percent_form_red.set_num_format('0.00%')
	percent_form_red.set_border()
	percent_form_red.set_font_color('red')
	percent_form_red.set_bold()

	simple_form_red = wb.add_format()
	simple_form_red.set_border()
	simple_form_red.set_bold()
	simple_form_red.set_font_color('red')

	date_form_yel = wb.add_format()
	date_form_yel.set_num_format('dd.mm.yyyy')
	date_form_yel.set_border()
	date_form_yel.set_bg_color('yellow')

	num_form_yel = wb.add_format()
	num_form_yel.set_num_format('###')
	num_form_yel.set_border()
	num_form_yel.set_bg_color('yellow')

	num_form2_yel = wb.add_format()
	num_form2_yel.set_border()
	num_form2_yel.set_num_format('###,#.00')
	num_form2_yel.set_bg_color('yellow')

	bold_form = wb.add_format()
	bold_form.set_bold()
	bold_form.set_align('center')
	bold_form.set_border()

	percent_form_yel = wb.add_format()
	percent_form_yel.set_num_format('0.00%')
	percent_form_yel.set_border()
	percent_form_yel.set_bg_color('yellow')

	simple_form_yel = wb.add_format()
	simple_form_yel.set_border()
	simple_form_yel.set_bg_color('yellow')


	for i in range(len(columns1)):
		ws.write(0, i, columns1[i], bold_form)
		
	for i in range(len(columns2)):
		ws.write(0, i+11, columns2[i], bold_form)
		
	for i in range(len(columns3)):
		ws.write(0, i+18, columns3[i], bold_form)
		
	for i in range(len(base_dec3)):
		if base_dec3.iloc[i]['Дата выдачи'].date() >= datetime.date(2018, 7, 1):
			for j in range(len(base_dec3.columns)):
				if (j == 0) | (j == 5):
					ws.write(i+1, j, base_dec3.iloc[i][j], date_form_yel)
				elif (j == 4):
					ws.write(i+1, j, base_dec3.iloc[i][j], num_form_yel)
				elif (j == 6) | (j == 7):
					ws.write(i+1, j, base_dec3.iloc[i][j], num_form2_yel)
				else:
					ws.write(i+1, j, base_dec3.iloc[i][j], simple_form_yel)
		else:
			for j in range(len(base_dec3.columns)):
				if (j == 0) | (j == 5):
					ws.write(i+1, j, base_dec3.iloc[i][j], date_form)
				elif (j == 4):
					ws.write(i+1, j, base_dec3.iloc[i][j], num_form)
				elif (j == 6) | (j == 7):
					ws.write(i+1, j, base_dec3.iloc[i][j], num_form2)
				else:
					ws.write(i+1, j, base_dec3.iloc[i][j], simple_form)
				
	for i in range(len(df_month)):
		if df_month.iloc[i]['Месяц выдачи'] in df_month_top3:
			for j in range(len(df_month.columns)):
				if (j == 1) | (j == 2):
					ws.write(i+1, j+11, df_month.iloc[i][j], num_form2_red)
				elif (j == 3) | (j == 5):
					ws.write(i+1, j+11, df_month.iloc[i][j], percent_form_red)
				else:
					ws.write(i+1, j+11, df_month.iloc[i][j], simple_form_red)
		else:
			for j in range(len(df_month.columns)):
				if (j == 1) | (j == 2):
					ws.write(i+1, j+11, df_month.iloc[i][j], num_form2)
				elif (j == 3) | (j == 5):
					ws.write(i+1, j+11, df_month.iloc[i][j], percent_form)
				else:
					ws.write(i+1, j+11, df_month.iloc[i][j], simple_form)

	for i in range(len(df_fil)):
		if df_fil.iloc[i]['Филиал'] in df_fil_top3:
			for j in range(len(df_fil.columns)):
				if (j == 1) | (j == 2):
					ws.write(i+1, j+18, df_fil.iloc[i][j], num_form2_red)
				elif (j == 3) | (j == 5):
					ws.write(i+1, j+18, df_fil.iloc[i][j], percent_form_red)
				else:
					ws.write(i+1, j+18, df_fil.iloc[i][j], simple_form_red)
		else:
			for j in range(len(df_fil.columns)):
				if (j == 1) | (j == 2):
					ws.write(i+1, j+18, df_fil.iloc[i][j], num_form2)
				elif (j == 3) | (j == 5):
					ws.write(i+1, j+18, df_fil.iloc[i][j], percent_form)
				else:
					ws.write(i+1, j+18, df_fil.iloc[i][j], simple_form)

	######################################   [61-90] -> [91-120]   ########################################### 
	base = all_rp
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
	base_dec3 = base_dec1[['OPERDATE1', 'CONTRNUM', 'FILIAL', 'CLIENTLABEL', 'IIN', 'OPENDATE', 'AMOUNTLOANKZT', 'LOANRESTKZT', 'PROSDAYS', 'PRODUCT']]
	open_ym = []
	tab = []
	tab2 = []
	fils = []
	amount = 0
	loan = 0

	base_dec2 = base_dec1
	base_dec2 = base_dec2.sort_values(by = 'OPENDATE_YM')
	base_dec2 = base_dec2.drop_duplicates('OPENDATE_YM', keep = 'first')

	base_dec4 = base_dec1
	base_dec4 = base_dec4.drop_duplicates('FILIAL', keep = 'first')
	fils = base_dec4.FILIAL.tolist()

	all_loan = base_dec3.LOANRESTKZT.sum()
	all_count = len(base_dec3)

	for i in range(len(base_dec2)):
		open_ym.append(base_dec2.iloc[i]['OPENDATE_YM'])
	for i in range(len(open_ym)):
		open_ym2 = base_dec1[base_dec1.OPENDATE_YM == open_ym[i]]
		loan  = open_ym2['LOANRESTKZT'].sum()
		amount = open_ym2['AMOUNTLOANKZT'].sum()
		if str(open_ym[i])[4:6] == '01':
			month = 'Январь'
		if str(open_ym[i])[4:6] == '02':
			month = 'Февраль'
		if str(open_ym[i])[4:6] == '03':
			month = 'Март'
		if str(open_ym[i])[4:6] == '04':
			month = 'Апрель'
		if str(open_ym[i])[4:6] == '05':
			month = 'Май'
		if str(open_ym[i])[4:6] == '06':
			month = 'Июнь'
		if str(open_ym[i])[4:6] == '07':
			month = 'Июль'
		if str(open_ym[i])[4:6] == '08':
			month = 'Август'
		if str(open_ym[i])[4:6] == '09':
			month = 'Сентябрь'
		if str(open_ym[i])[4:6] == '10':
			month = 'Октябрь'
		if str(open_ym[i])[4:6] == '11':
			month = 'Ноябрь'
		if str(open_ym[i])[4:6] == '12':
			month = 'Декабрь'
		tab.append({'Месяц выдачи': month + ' ' + str(open_ym[i])[0:4], 'Сумма ОД':loan, 'Сумма займа':amount, 'Кол-во':len(open_ym2), 
					'Доля по ОД':loan/all_loan, 'Доля по кол-ву':len(open_ym2)/all_count})
		loan = 0
		amount = 0
		open_ym2.iloc[0:0]

	for i in range(len(fils)):
		fils2 = base_dec1[base_dec1.FILIAL == fils[i]]
		loan = fils2['LOANRESTKZT'].sum()
		amount = fils2['AMOUNTLOANKZT'].sum()
		tab2.append({'Филиал':fils[i], 'Сумма займа':amount, 'Сумма ОД':loan, 'Кол-во':len(fils2),
					'Доля по ОД':loan/all_loan, 'Доля по кол-ву':len(fils2)/all_count})
		loan = 0
		amount = 0
		fils2.iloc[0:0]

	base_dec3.columns = ['Дата отчета', 'Договор', 'Филиал', 'ФИО', 'ИИН', 'Дата выдачи', 'Сумма выдачи', 'ОД', 'Просрочка', 'Продукт']
	df_month = pd.DataFrame(data=tab)
	df_month = df_month.reindex(['Месяц выдачи', 'Сумма займа', 'Сумма ОД', 'Доля по ОД', 'Кол-во', 'Доля по кол-ву'], axis=1)
	df_fil = pd.DataFrame(data=tab2)
	df_fil = df_fil.reindex(['Филиал', 'Сумма займа', 'Сумма ОД', 'Доля по ОД', 'Кол-во', 'Доля по кол-ву'], axis=1)


	ws = wb.add_worksheet('61-90 -> 91-120')
	columns1 = ['Дата отчета', 'Договор', 'Филиал', 'ФИО', 'ИИН', 'Дата выдачи', 'Сумма выдачи', 'ОД', 'Просрочка', 'Продукт']
	columns2 = ['Месяц выдачи', 'Сумма займа', 'Сумма ОД', 'Доля по ОД', 'Кол-во', 'Доля по кол-ву']
	columns3 = ['Филиал', 'Сумма займа', 'Сумма ОД', 'Доля по ОД', 'Кол-во', 'Доля по кол-ву']

	df_month_top3 = []
	df_fil_top3 = []

	df_month_sort = df_month
	df_month_sort = df_month_sort.sort_values(by = 'Доля по ОД', ascending=False)
	for i in range(3):        
		df_month_top3.append(df_month_sort.iloc[i]['Месяц выдачи'])
		
	df_fil_sort = df_fil
	df_fil_sort = df_fil_sort.sort_values(by = 'Доля по ОД', ascending=False)
	for i in range(3):
		df_fil_top3.append(df_fil_sort.iloc[i]['Филиал'])

	date_form = wb.add_format()
	date_form.set_num_format('dd.mm.yyyy')
	date_form.set_border()

	num_form = wb.add_format()
	num_form.set_num_format('###')
	num_form.set_border()

	num_form2 = wb.add_format()
	num_form2.set_border()
	num_form2.set_num_format('###,#.00')

	bold_form = wb.add_format()
	bold_form.set_bold()
	bold_form.set_align('center')
	bold_form.set_border()

	percent_form = wb.add_format()
	percent_form.set_num_format('0.00%')
	percent_form.set_border()

	simple_form = wb.add_format()
	simple_form.set_border()
		  
	date_form_red = wb.add_format()
	date_form_red.set_num_format('dd.mm.yyyy')
	date_form_red.set_border()
	date_form_red.set_font_color('red')
	date_form_red.set_bold()

	num_form_red = wb.add_format()
	num_form_red.set_num_format('###')
	num_form_red.set_border()
	num_form_red.set_font_color('red')
	num_form_red.set_bold()

	num_form2_red = wb.add_format()
	num_form2_red.set_border()
	num_form2_red.set_num_format('###,#.00')
	num_form2_red.set_font_color('red')
	num_form2_red.set_bold()

	bold_form = wb.add_format()
	bold_form.set_bold()
	bold_form.set_align('center')
	bold_form.set_border()

	percent_form_red = wb.add_format()
	percent_form_red.set_num_format('0.00%')
	percent_form_red.set_border()
	percent_form_red.set_font_color('red')
	percent_form_red.set_bold()

	simple_form_red = wb.add_format()
	simple_form_red.set_border()
	simple_form_red.set_bold()
	simple_form_red.set_font_color('red')

	date_form_yel = wb.add_format()
	date_form_yel.set_num_format('dd.mm.yyyy')
	date_form_yel.set_border()
	date_form_yel.set_bg_color('yellow')

	num_form_yel = wb.add_format()
	num_form_yel.set_num_format('###')
	num_form_yel.set_border()
	num_form_yel.set_bg_color('yellow')

	num_form2_yel = wb.add_format()
	num_form2_yel.set_border()
	num_form2_yel.set_num_format('###,#.00')
	num_form2_yel.set_bg_color('yellow')

	bold_form = wb.add_format()
	bold_form.set_bold()
	bold_form.set_align('center')
	bold_form.set_border()

	percent_form_yel = wb.add_format()
	percent_form_yel.set_num_format('0.00%')
	percent_form_yel.set_border()
	percent_form_yel.set_bg_color('yellow')

	simple_form_yel = wb.add_format()
	simple_form_yel.set_border()
	simple_form_yel.set_bg_color('yellow')


	for i in range(len(columns1)):
		ws.write(0, i, columns1[i], bold_form)
		
	for i in range(len(columns2)):
		ws.write(0, i+11, columns2[i], bold_form)
		
	for i in range(len(columns3)):
		ws.write(0, i+18, columns3[i], bold_form)
		
	for i in range(len(base_dec3)):
		if base_dec3.iloc[i]['Дата выдачи'].date() >= datetime.date(2018, 7, 1):
			for j in range(len(base_dec3.columns)):
				if (j == 0) | (j == 5):
					ws.write(i+1, j, base_dec3.iloc[i][j], date_form_yel)
				elif (j == 4):
					ws.write(i+1, j, base_dec3.iloc[i][j], num_form_yel)
				elif (j == 6) | (j == 7):
					ws.write(i+1, j, base_dec3.iloc[i][j], num_form2_yel)
				else:
					ws.write(i+1, j, base_dec3.iloc[i][j], simple_form_yel)
		else:
			for j in range(len(base_dec3.columns)):
				if (j == 0) | (j == 5):
					ws.write(i+1, j, base_dec3.iloc[i][j], date_form)
				elif (j == 4):
					ws.write(i+1, j, base_dec3.iloc[i][j], num_form)
				elif (j == 6) | (j == 7):
					ws.write(i+1, j, base_dec3.iloc[i][j], num_form2)
				else:
					ws.write(i+1, j, base_dec3.iloc[i][j], simple_form)
				
	for i in range(len(df_month)):
		if df_month.iloc[i]['Месяц выдачи'] in df_month_top3:
			for j in range(len(df_month.columns)):
				if (j == 1) | (j == 2):
					ws.write(i+1, j+11, df_month.iloc[i][j], num_form2_red)
				elif (j == 3) | (j == 5):
					ws.write(i+1, j+11, df_month.iloc[i][j], percent_form_red)
				else:
					ws.write(i+1, j+11, df_month.iloc[i][j], simple_form_red)
		else:
			for j in range(len(df_month.columns)):
				if (j == 1) | (j == 2):
					ws.write(i+1, j+11, df_month.iloc[i][j], num_form2)
				elif (j == 3) | (j == 5):
					ws.write(i+1, j+11, df_month.iloc[i][j], percent_form)
				else:
					ws.write(i+1, j+11, df_month.iloc[i][j], simple_form)

	for i in range(len(df_fil)):
		if df_fil.iloc[i]['Филиал'] in df_fil_top3:
			for j in range(len(df_fil.columns)):
				if (j == 1) | (j == 2):
					ws.write(i+1, j+18, df_fil.iloc[i][j], num_form2_red)
				elif (j == 3) | (j == 5):
					ws.write(i+1, j+18, df_fil.iloc[i][j], percent_form_red)
				else:
					ws.write(i+1, j+18, df_fil.iloc[i][j], simple_form_red)
		else:
			for j in range(len(df_fil.columns)):
				if (j == 1) | (j == 2):
					ws.write(i+1, j+18, df_fil.iloc[i][j], num_form2)
				elif (j == 3) | (j == 5):
					ws.write(i+1, j+18, df_fil.iloc[i][j], percent_form)
				else:
					ws.write(i+1, j+18, df_fil.iloc[i][j], simple_form)
				
	wb.close()

