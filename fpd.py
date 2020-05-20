import pandas as pd
import datetime
import seaborn as sns
from matplotlib import pyplot as plt

def fpd_data(base):
    df2 = base
    tab = []
    df2 = df2[df2.OPENDATE >= datetime.date(2016, 12, 1)]
    df2 = df2.sort_values(by = ['CONTRNUM', 'OPERDATE1'])
    df2['SHIFT_CON'] = df2.CONTRNUM.shift()
    df2['SHIFT_AMOUNT'] = df2.AMOUNTLOANKZT.shift()
    df2['SHIFT_LOAN'] = df2.LOANRESTKZT.shift()
    df2['SHIFT_DEL'] = df2.DEL_DAYS.shift()
    df2['FPD'] = df2['SPD'] = df2['TPD'] = df2['FDEF'] = df2['F120'] = df2['F150'] = df2['LOSS'] = 0
    #####################FPD########################
    df3 = df2[(df2.MONTHS_SINCE_OPEN == 0)&(df2.DEL_1P == 1)]
    df4 = df2[df2.CONTRNUM == df2.SHIFT_CON]
    df4 = df4[(df4.MONTHS_SINCE_OPEN == 1)&(df4.DEL_1P == 1)]
    df5 = df2[df2.CONTRNUM == df2.SHIFT_CON]
    df5 = df5[(df5.SHIFT_AMOUNT == df5.SHIFT_LOAN)&(df5.MONTHS_SINCE_OPEN == 2)&(df5.DEL_DAYS <= 3)&(df5.DEL_1P == 1)]
    df6 = pd.concat([df3, df4, df5], axis=0)
    for i in range(len(df6)):
        df2.at[df6.index[i], 'FPD'] = 1 

    df2['SHIFT_FPD'] = df2.FPD.shift()

    df7 = df2[df2.FPD == 1]
    df7 = df7[df7.FPD == df7.SHIFT_FPD]
    for i in range(len(df7)):
        df2.at[df7.index[i], 'FPD'] = 0

    df2['SHIFT_FPD'] = df2.FPD.shift()
    df7 = df2[df2.FPD == 1]
    #####################SPD########################
    df8 = df2[df2.SHIFT_FPD == 1]
    df8 = df8[(df8.AMOUNTLOANKZT == df8.LOANRESTKZT)&(df8.CONTRNUM == df8.SHIFT_CON)&(df8.DEL_30P == 1)]

    df9 = df2[df2.SHIFT_FPD == 1]
    df9 = df9[(df9.LOANRESTKZT == df9.SHIFT_LOAN)&(df9.CONTRNUM == df9.SHIFT_CON)&(df9.DEL_30P == 1)]

    df10 = df2[df2.SHIFT_FPD == 1]
    df10 = df10[(df10.DEL_DAYS > df10.SHIFT_DEL)&(df10.CONTRNUM == df10.SHIFT_CON)&(df10.DEL_30P == 1)]

    df11 = pd.concat([df8, df9, df10], axis=0)
    df11 = df11.drop_duplicates('CONTRNUM', keep = 'first')

    for i in range(len(df11)):
        df2.at[df11.index[i], 'SPD'] = 1

    df2['SHIFT_SPD'] = df2.SPD.shift()
    #####################TPD########################
    df12 = df2[df2.SHIFT_SPD == 1]
    df12 = df12[(df12.AMOUNTLOANKZT == df12.LOANRESTKZT)&(df12.CONTRNUM == df12.SHIFT_CON)&(df12.DEL_60P == 1)]

    df13 = df2[df2.SHIFT_SPD == 1]
    df13 = df13[(df13.LOANRESTKZT == df13.SHIFT_LOAN)&(df13.CONTRNUM == df13.SHIFT_CON)&(df13.DEL_60P == 1)]

    df14 = df2[df2.SHIFT_SPD == 1]
    df14 = df14[(df14.DEL_DAYS > df14.SHIFT_DEL)&(df14.CONTRNUM == df14.SHIFT_CON)&(df14.DEL_60P == 1)]

    df15 = pd.concat([df12, df13, df14], axis=0)
    df15 = df15.drop_duplicates('CONTRNUM', keep = 'first')

    for i in range(len(df15)):
        df2.at[df15.index[i], 'TPD'] = 1

    df2['SHIFT_TPD'] = df2.TPD.shift()
    #####################FDEF########################
    df16 = df2[df2.SHIFT_TPD == 1]
    df16 = df16[(df16.AMOUNTLOANKZT == df16.LOANRESTKZT)&(df16.CONTRNUM == df16.SHIFT_CON)&(df16.DEL_90P == 1)]

    df17 = df2[df2.SHIFT_TPD == 1]
    df17 = df17[(df17.LOANRESTKZT == df17.SHIFT_LOAN)&(df17.CONTRNUM == df17.SHIFT_CON)&(df17.DEL_90P == 1)]

    df18 = df2[df2.SHIFT_TPD == 1]
    df18 = df18[(df18.DEL_DAYS > df18.SHIFT_DEL)&(df18.CONTRNUM == df18.SHIFT_CON)&(df18.DEL_90P == 1)]

    df19 = pd.concat([df16, df17, df18], axis=0)
    df19 = df19.drop_duplicates('CONTRNUM', keep = 'first')

    for i in range(len(df19)):
        df2.at[df19.index[i], 'FDEF'] = 1

    df2['SHIFT_FDEF'] = df2.FDEF.shift()
    #####################F120########################
    df20 = df2[df2.SHIFT_FDEF == 1]
    df20 = df20[(df20.AMOUNTLOANKZT == df20.LOANRESTKZT)&(df20.CONTRNUM == df20.SHIFT_CON)&(df20.DEL_120P == 1)]

    df21 = df2[df2.SHIFT_FDEF == 1]
    df21 = df21[(df21.LOANRESTKZT == df21.SHIFT_LOAN)&(df21.CONTRNUM == df21.SHIFT_CON)&(df21.DEL_120P == 1)]

    df22 = df2[df2.SHIFT_FDEF == 1]
    df22 = df22[(df22.DEL_DAYS > df22.SHIFT_DEL)&(df22.CONTRNUM == df22.SHIFT_CON)&(df22.DEL_120P == 1)]

    df23 = pd.concat([df20, df21, df22], axis=0)
    df23 = df23.drop_duplicates('CONTRNUM', keep = 'first')

    for i in range(len(df23)):
        df2.at[df23.index[i], 'F120'] = 1

    df2['SHIFT_F120'] = df2.F120.shift()
    #####################F150########################
    df24 = df2[df2.SHIFT_F120 == 1]
    df24 = df24[(df24.AMOUNTLOANKZT == df24.LOANRESTKZT)&(df24.CONTRNUM == df24.SHIFT_CON)&(df24.DEL_150P == 1)]

    df25 = df2[df2.SHIFT_F120 == 1]
    df25 = df25[(df25.LOANRESTKZT == df25.SHIFT_LOAN)&(df25.CONTRNUM == df25.SHIFT_CON)&(df25.DEL_150P == 1)]

    df26 = df2[df2.SHIFT_F120 == 1]
    df26 = df26[(df26.DEL_DAYS > df26.SHIFT_DEL)&(df26.CONTRNUM == df26.SHIFT_CON)&(df26.DEL_150P == 1)]

    df27 = pd.concat([df24, df25, df26], axis=0)
    df27 = df27.drop_duplicates('CONTRNUM', keep = 'first')

    for i in range(len(df27)):
        df2.at[df27.index[i], 'F150'] = 1

    df2['SHIFT_F150'] = df2.F150.shift()
    #####################LOSS########################
    df28 = df2[df2.SHIFT_F150 == 1]
    df28 = df28[(df28.AMOUNTLOANKZT == df28.LOANRESTKZT)&(df28.CONTRNUM == df28.SHIFT_CON)&(df28.DEL_180P == 1)]

    df29 = df2[df2.SHIFT_F150 == 1]
    df29 = df29[(df29.LOANRESTKZT == df29.SHIFT_LOAN)&(df29.CONTRNUM == df29.SHIFT_CON)&(df29.DEL_180P == 1)]

    df30 = df2[df2.SHIFT_F150 == 1]
    df30 = df30[(df30.DEL_DAYS > df30.SHIFT_DEL)&(df30.CONTRNUM == df30.SHIFT_CON)&(df30.DEL_180P == 1)]

    df31 = pd.concat([df28, df29, df30], axis=0)
    df31 = df31.drop_duplicates('CONTRNUM', keep = 'first')

    for i in range(len(df31)):
        df2.at[df31.index[i], 'LOSS'] = 1
    #####################TABLE########################
    df_open = df2
    df_open = df_open.sort_values(by = 'OPENDATE_YM')
    df_open = df_open.drop_duplicates('OPENDATE_YM', keep = 'first')
    opens = df_open.OPENDATE_YM.tolist()

    dic = {'01':'January', '02':'February', '03':'March', '04':'April', '05':'May', '06':'June', '07':'July', '08':'August', '09':'September', '10':'October', '11':'November', '12':'December'}
    for i in range(len(opens)):
        df_fpd = df7[df7.OPENDATE_YM == opens[i]]
        fpd_sum = df_fpd.LOANRESTKZT.sum()
        df_spd = df11[df11.OPENDATE_YM == opens[i]]
        spd_sum = df_spd.LOANRESTKZT.sum()
        df_tpd = df15[df15.OPENDATE_YM == opens[i]]
        tpd_sum = df_tpd.LOANRESTKZT.sum()
        df_fdef = df19[df19.OPENDATE_YM == opens[i]]
        fdef_sum = df_fdef.LOANRESTKZT.sum()
        df_loss = df31[df31.OPENDATE_YM == opens[i]]
        loss_sum = df_loss.LOANRESTKZT.sum()

        all_ = df2[df2.OPENDATE_YM == opens[i]]
        all_ = all_[all_.MONTHS_SINCE_OPEN == 0]
        all_sum = all_.AMOUNTLOANKZT.sum()

        open_ym = dic[str(opens[i])[4:6]] + ' ' + str(opens[i])[0:4]
        tab.append({'Месяц выдачи':open_ym, 'Кол-во':len(all_), 'Сумма выдачи, млн, KZT':round(all_sum/1000000, 1), 'Кол-во FPD':len(df_fpd), 'Доля FPD, %':round((fpd_sum/all_sum)*100, 1),
                   'Кол-во SPD':len(df_spd), 'Доля SPD, %':round((spd_sum/all_sum)*100, 1),'Кол-во TPD':len(df_tpd), 'Доля TPD, %':round((tpd_sum/all_sum)*100, 1),
                   'Кол-во FDef':len(df_fdef), 'Доля FDef, %':round((fdef_sum/all_sum)*100, 1),'Кол-во Loss':len(df_loss), 'Доля Loss, %':round((loss_sum/all_sum)*100, 1)})

    df_tab = pd.DataFrame(tab)
    df_tab = df_tab.reindex(['Месяц выдачи', 'Кол-во', 'Сумма выдачи, млн, KZT', 'Кол-во FPD', 'Доля FPD, %', 'Кол-во SPD', 'Доля SPD, %', 'Кол-во TPD', 'Доля TPD, %', 'Кол-во FDef', 'Доля FDef, %', 'Кол-во Loss', 'Доля Loss, %'], axis=1)
    
    return df_tab

def graph_fpd(df_tab):
    f, ax = plt.subplots()
    plt.xticks(rotation=90)
    plt.style.use('default')
    plt.style.use('bmh')
    plt.tick_params(labelsize=7)
    ax = sns.barplot(x=df_tab['Месяц выдачи'], y=df_tab['Сумма выдачи, млн, KZT'], data=df_tab, color='darkturquoise')
    ax.tick_params(axis='y')
    ax1 = ax.twinx()
    ax1=sns.pointplot(ax=ax1,x=df_tab['Месяц выдачи'],y=df_tab['Доля FPD, %'],data=df_tab,color='blue', scale=0.3)
    ax1=sns.pointplot(ax=ax1,x=df_tab['Месяц выдачи'],y=df_tab['Доля SPD, %'],data=df_tab,color='orange', scale=0.3)
    ax1=sns.pointplot(ax=ax1,x=df_tab['Месяц выдачи'],y=df_tab['Доля TPD, %'],data=df_tab,color='red', scale=0.3)
    ax1=sns.pointplot(ax=ax1,x=df_tab['Месяц выдачи'],y=df_tab['Доля FDef, %'],data=df_tab,color='green', scale=0.3)
    ax1=sns.pointplot(ax=ax1,x=df_tab['Месяц выдачи'],y=df_tab['Доля Loss, %'],data=df_tab,color='black', scale=0.3)
    ax1.tick_params(axis='y')
    f.tight_layout()
    ax1.legend(handles=ax1.lines[::len(df_tab)+1], labels=["Доля FPD, %","Доля SPD, %","Доля TPD, %", "Доля FDef, %", "Доля Loss, %"], ncol=5, loc=9, fontsize=7)
    #ax.legend(handles=ax.lines[::len(df)+1], labels=["Основной долг (млн.)"])
    f.set_size_inches(17,7)
    plt.xlabel('')
    plt.ylabel('')
    return f