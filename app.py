import sys
from PyQt5.QtWidgets import QMainWindow, QApplication, QWidget, QAction, QTableWidget,QTableWidgetItem, QVBoxLayout, QPushButton, QLabel, QFileDialog, QComboBox, QMessageBox, QGridLayout, QGroupBox, QDialog, QMenu, QLineEdit
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import pyqtSlot, Qt, QCoreApplication
import pandas as pd
import datetime
import dateutil.relativedelta as delta
import numpy as np
import vintage
import lists
import basket_an
import dpk_report
import port_dynamics
import vintage_mis
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import plotly.offline as plo
import calcscore
import krisha_model
from sklearn.externals import joblib
import fpd

class App(QMainWindow):
 
    def __init__(self):
        super().__init__()
        self.initUI()
 
    def initUI(self):
        self.df = pd.read_excel('C:/Users/a.moldabekov/Desktop/ALL_RP_0112.xlsx')
        self.dwh = pd.read_excel('C:/Users/a.moldabekov/Desktop/DWH_03.12.2019.xlsx')
        self.dwh = self.dwh.dropna(subset = ['ABIS_CONTRACT_NUM'])
        self.dwh = self.dwh.set_index('ABIS_CONTRACT_NUM')
        self.model = joblib.load('C:/Users/a.moldabekov/Desktop/krisha.pkl')
        self.dfs = [pd.DataFrame(columns = ['Отчетная дата', 'Филиал', 'Номер договора', 'ФИО', 'ИИН', 'Дата выдачи', 'Дата закрытия', 'Сумма выдачи', 'Остаток долга', 'Продукт', 'Просрочка'])]
        self.setAutoFillBackground(True)
        p = self.palette()
        p.setColor(self.backgroundRole(), Qt.white)
        self.setPalette(p)
        self.setGeometry(300, 10, 950, 1000)
        
        self.menu_main = QMenu()
        self.menu_main.setTitle('Главная')
        self.menu_reports = QMenu()
        self.menu_reports.setTitle('Отчеты')
        self.menu_models = QMenu()
        self.menu_models.setTitle('Модели')
        self.menuBar().addMenu(self.menu_main)
        self.menuBar().addMenu(self.menu_reports)
        self.menuBar().addMenu(self.menu_models)
        
        self.act_korzina = QAction('Корзина...', self)
        df = pd.concat(self.dfs, axis=0)
        self.act_korzina.triggered.connect(self.createBucketDialog)
        self.act_about = QAction('О программе', self)
        self.act_about.triggered.connect(self.aboutDialog)
        self.act_exit = QAction('Выход', self)
        self.act_exit.triggered.connect(self.close)
        
        self.act_basket_an = QAction('Анализ по корзинам', self)
        self.act_basket_an.triggered.connect(self.basket_report)
        self.act_dpk = QAction('Файл для ДПК', self)
        self.act_dpk.triggered.connect(self.dpk_report_)
        
        self.act_score = QAction('Скоринговая модель', self)
        self.act_score.triggered.connect(self.calcScore)
        self.act_krisha = QAction('Модель krisha.kz', self)
        self.act_krisha.triggered.connect(self.predictModel)
        
        self.menu_main.addAction(self.act_korzina)
        self.menu_main.addAction(self.act_about)
        self.menu_main.addAction(self.act_exit)
        
        self.menu_reports.addAction(self.act_basket_an)
        self.menu_reports.addAction(self.act_dpk)
        
        self.menu_models.addAction(self.act_score)
        self.menu_models.addAction(self.act_krisha)
        
        self.btn_show = QPushButton('Анализ', self)
        self.btn_show.setGeometry(10, 40, 80, 20)
        self.btn_show.clicked.connect(self.createTable)
        self.btn_show.setStyleSheet('border: 1px solid #00B0F0;border-radius:4px;background-color:#FFFFFF;color:#7030A0;')
        
        self.com_type = QComboBox(self)
        self.com_type.setGeometry(100, 40, 150, 20)
        self.com_type.addItems(['Portfolio Dynamics', 'Portfolio Dynamics Filial', 'Vintage MIS', 'FPD SPD TPD', 'Vintage Amount 1+', 'Vintage Amount 30+', 'Vintage Amount 90+', 'Vintage Amount 180+',
                               'Vintage Count 1+', 'Vintage Count 30+', 'Vintage Count 90+', 'Vintage Count 180+', 'Quarter Vintage Amount 1+', 'Quarter Vintage Amount 90+', 'Bucket Analysis', 'Lag 30+, 60+, 90+', 'Roll Rates'])
        self.com_type.setStyleSheet('border: 1px solid #00B0F0;border-radius:4px;background-color:#FFFFFF;color:#7030A0;')
        
        self.com_prod = QComboBox(self)
        self.com_prod.setGeometry(260, 40, 150, 20)
        self.com_prod.addItems(lists.products())
        self.com_prod.setStyleSheet('border: 1px solid #00B0F0;border-radius:4px;background-color:#FFFFFF;color:#7030A0;')
        
        self.com_fil = QComboBox(self)
        self.com_fil.setGeometry(420, 40, 100, 20)
        self.com_fil.addItems(lists.branches())
        self.com_fil.setStyleSheet('border: 1px solid #00B0F0;border-radius:4px;background-color:#FFFFFF;color:#7030A0;')
        
        self.btn_graph = QPushButton('График', self)
        self.btn_graph.setGeometry(540, 40, 100, 20)
        self.btn_graph.clicked.connect(self.createGraph)
        self.btn_graph.setStyleSheet('border: 1px solid #00B0F0;border-radius:4px;background-color:#FFFFFF;color:#7030A0;')
        self.show()
    
    def createTable(self):
        if self.com_type.currentText() == 'Vintage MIS':
            self.nd = NewDialog(self)
            self.nd.setGeometry(10, 70, 900, 900)
            base = self.df
            if self.com_prod.currentText() != 'Все продукты':
                base = base[base.PRODUCT == self.com_prod.currentText()]

            if self.com_fil.currentText() != 'Все филиалы':
                base = base[base.FILIAL == self.com_fil.currentText()]

            df_tab = vintage_mis.viewTable(base)
            self.createNewDialog(df_tab)
            self.nd.tableWidget.cellDoubleClicked.connect(self.show_cell_pos_vintage_mis)

        if self.com_type.currentText() == 'Portfolio Dynamics':
            self.nd = NewDialog(self)
            self.nd.setGeometry(10, 70, 900, 900)
            base = self.df
            
            if self.com_prod.currentText() != 'Все продукты':
                base = base[base.PRODUCT == self.com_prod.currentText()]
            if self.com_fil.currentText() != 'Все филиалы':
                base = base[base.FILIAL == self.com_fil.currentText()]
            
            df_tab = port_dynamics.viewTable(base)
            self.createNewDialog(df_tab)
            self.nd.tableWidget.cellDoubleClicked.connect(self.show_cell_port_dynamics)
            ##########Portfolio Dynamics Filial###########
        #if type_text == 'Portfolio Dynamics Filial':

        if self.com_type.currentText() == 'Vintage Amount 1+':
            self.nd = NewDialog(self)
            self.nd.setGeometry(10, 70, 900, 900)
            base = self.df

            if self.com_prod.currentText() != 'Все продукты':
                base = base[base.PRODUCT == self.com_prod.currentText()]
            if self.com_fil.currentText() != 'Все филиалы':
                base = base[base.FILIAL == self.com_fil.currentText()]
                
            df2 = vintage.vintage_amount(1, 0, base)
            self.createNewDialog(df2) 
            self.nd.tableWidget.cellDoubleClicked.connect(self.show_cell_vintage_1)
        
        if self.com_type.currentText() == 'Vintage Amount 30+':
            self.nd = NewDialog(self)
            self.nd.setGeometry(10, 70, 900, 900)
            base = self.df

            if self.com_prod.currentText() != 'Все продукты':
                base = base[base.PRODUCT == self.com_prod.currentText()]
            if self.com_fil.currentText() != 'Все филиалы':
                base = base[base.FILIAL == self.com_fil.currentText()]
                
            df2 = vintage.vintage_amount(2, 30, base)
            self.createNewDialog(df2)
            self.nd.tableWidget.cellDoubleClicked.connect(self.show_cell_vintage_30)
            
        if self.com_type.currentText() == 'Vintage Amount 90+':
            self.nd = NewDialog(self)
            self.nd.setGeometry(10, 70, 900, 900)
            base = self.df

            if self.com_prod.currentText() != 'Все продукты':
                base = base[base.PRODUCT == self.com_prod.currentText()]
            if self.com_fil.currentText() != 'Все филиалы':
                base = base[base.FILIAL == self.com_fil.currentText()]
                
            df2 = vintage.vintage_amount(4, 90, base)
            self.createNewDialog(df2)
            self.nd.tableWidget.cellDoubleClicked.connect(self.show_cell_vintage_90)
            
        if self.com_type.currentText() == 'Vintage Amount 180+':
            self.nd = NewDialog(self)
            self.nd.setGeometry(10, 70, 900, 900)
            base = self.df

            if self.com_prod.currentText() != 'Все продукты':
                base = base[base.PRODUCT == self.com_prod.currentText()]
            if self.com_fil.currentText() != 'Все филиалы':
                base = base[base.FILIAL == self.com_fil.currentText()]
                
            df2 = vintage.vintage_amount(7, 180, base)
            self.createNewDialog(df2)
            self.nd.tableWidget.cellDoubleClicked.connect(self.show_cell_vintage_180)
            
        if self.com_type.currentText() == 'Vintage Count 1+':
            self.nd = NewDialog(self)
            self.nd.setGeometry(10, 70, 900, 900)
            base = self.df

            if self.com_prod.currentText() != 'Все продукты':
                base = base[base.PRODUCT == self.com_prod.currentText()]
            if self.com_fil.currentText() != 'Все филиалы':
                base = base[base.FILIAL == self.com_fil.currentText()]
                
            df2 = vintage.vintage_count(1, 0, base)
            self.createNewDialog(df2)
            self.nd.tableWidget.cellDoubleClicked.connect(self.show_cell_vintage_1)
        
        if self.com_type.currentText() == 'Vintage Count 30+':
            self.nd = NewDialog(self)
            self.nd.setGeometry(10, 70, 900, 900)
            base = self.df

            if self.com_prod.currentText() != 'Все продукты':
                base = base[base.PRODUCT == self.com_prod.currentText()]
            if self.com_fil.currentText() != 'Все филиалы':
                base = base[base.FILIAL == self.com_fil.currentText()]
                
            df2 = vintage.vintage_count(2, 30, base)
            self.createNewDialog(df2)  
            self.nd.tableWidget.cellDoubleClicked.connect(self.show_cell_vintage_30)
            
        if self.com_type.currentText() == 'Vintage Count 90+':
            self.nd = NewDialog(self)
            self.nd.setGeometry(10, 70, 900, 900)
            base = self.df

            if self.com_prod.currentText() != 'Все продукты':
                base = base[base.PRODUCT == self.com_prod.currentText()]
            if self.com_fil.currentText() != 'Все филиалы':
                base = base[base.FILIAL == self.com_fil.currentText()]
                
            df2 = vintage.vintage_count(4, 90, base)
            self.createNewDialog(df2)
            self.nd.tableWidget.cellDoubleClicked.connect(self.show_cell_vintage_90)
            
        if self.com_type.currentText() == 'Vintage Count 180+':
            self.nd = NewDialog(self)
            self.nd.setGeometry(10, 70, 900, 900)
            base = self.df

            if self.com_prod.currentText() != 'Все продукты':
                base = base[base.PRODUCT == self.com_prod.currentText()]
            if self.com_fil.currentText() != 'Все филиалы':
                base = base[base.FILIAL == self.com_fil.currentText()]
                
            df2 = vintage.vintage_count(7, 180, base)
            self.createNewDialog(df2)
            self.nd.tableWidget.cellDoubleClicked.connect(self.show_cell_vintage_180)
        
        if self.com_type.currentText() == 'FPD SPD TPD':
            self.nd = NewDialog(self)
            self.nd.setGeometry(10, 70, 900, 900)
            base = self.df

            if self.com_prod.currentText() != 'Все продукты':
                base = base[base.PRODUCT == self.com_prod.currentText()]
            if self.com_fil.currentText() != 'Все филиалы':
                base = base[base.FILIAL == self.com_fil.currentText()]
                
            df2 = fpd.fpd_data(base)
            self.createNewDialog(df2)
            
    def show_cell_pos_vintage_mis(self):
        itemRow = self.nd.tableWidget.currentItem().row()
        itemCol = self.nd.tableWidget.currentItem().column()
        month = self.nd.tableWidget.item(self.nd.tableWidget.currentRow(), 0).text()
        months = {'January':'01', 'February':'02', 'March':'03', 'April':'04', 'May':'05', 'June':'06', 'July':'07', 'August':'08', 'September':'09', 'October':'10', 'November':'11', 'December':'12'}
        open_ym = int(month.split(' ')[1] + months[month.split(' ')[0]])
        base = self.df
        if self.com_prod.currentText() != 'Все продукты':
            base = base[base.PRODUCT == self.com_prod.currentText()]

        if self.com_fil.currentText() != 'Все филиалы':
            base = base[base.FILIAL == self.com_fil.currentText()]
            
        if itemCol == 3:
            base = base[(base.MONTHS_SINCE_OPEN == 1)&(base.OPENDATE_YM == open_ym)&(base.PROSDAYS > 0)]
        if itemCol == 4:
            base = base[(base.MONTHS_SINCE_OPEN == 3)&(base.OPENDATE_YM == open_ym)&(base.PROSDAYS > 30)]
        if itemCol == 5:
            base = base[(base.MONTHS_SINCE_OPEN == 6)&(base.OPENDATE_YM == open_ym)&(base.PROSDAYS > 90)]
        if (itemCol == 0) or (itemCol == 1) or (itemCol == 2):
            base = base.iloc[0:0]
        
        base = base[['OPERDATE1', 'FILIAL', 'CONTRNUM', 'CLIENTLABEL', 'IIN', 'OPENDATE', 'CLOSEDATE', 'AMOUNTLOANKZT', 'LOANRESTKZT', 'PRODUCT', 'PROSDAYS']]
        base.OPERDATE1 = pd.to_datetime(base.OPERDATE1, format="%d%m%Y%").dt.strftime('%d.%m.%Y')
        base.OPENDATE = pd.to_datetime(base.OPENDATE, format="%d%m%Y%").dt.strftime('%d.%m.%Y')
        base.CLOSEDATE = pd.to_datetime(base.CLOSEDATE, format="%d%m%Y%").dt.strftime('%d.%m.%Y')
        base.AMOUNTLOANKZT = ['{:,}'.format(amount).replace(',', ' ') for amount in base.AMOUNTLOANKZT]
        base.LOANRESTKZT = ['{:,}'.format(loan).replace(',', ' ') for loan in base.LOANRESTKZT]
        self.base_use = base
        self.base_use.columns = ['Отчетная дата', 'Филиал', 'Номер договора', 'ФИО', 'ИИН', 'Дата выдачи', 'Дата закрытия', 'Сумма выдачи', 'Остаток долга', 'Продукт', 'Просрочка']
        
        self.createTabDialog(base)
        
    def show_cell_port_dynamics(self):
        itemCol = self.nd.tableWidget.currentItem().column()
        oper = self.nd.tableWidget.item(self.nd.tableWidget.currentRow(), 0).text()
        operdate = datetime.date(int(oper[6:10]), int(oper[3:5]), int(oper[0:2]))
        base = self.df
        if self.com_prod.currentText() != 'Все продукты':
            base = base[base.PRODUCT == self.com_prod.currentText()]

        if self.com_fil.currentText() != 'Все филиалы':
            base = base[base.FILIAL == self.com_fil.currentText()]
            
        if itemCol == 2:
            base = base[(base.OPERDATE1 == operdate)&(base.PROSDAYS > 0)]
        if itemCol == 3:
            base = base[(base.OPERDATE1 == operdate)&(base.PROSDAYS > 30)]
        if itemCol == 4:
            base = base[(base.OPERDATE1 == operdate)&(base.PROSDAYS > 90)]
        if itemCol == 5:
            base = base[(base.OPERDATE1 == operdate)&(base.PROSDAYS > 180)]    
        if (itemCol == 0) or (itemCol == 1):
            base = base.iloc[0:0]
            
        base = base[['OPERDATE1', 'FILIAL', 'CONTRNUM', 'CLIENTLABEL', 'IIN', 'OPENDATE', 'CLOSEDATE', 'AMOUNTLOANKZT', 'LOANRESTKZT', 'PRODUCT', 'PROSDAYS']]
        base.OPERDATE1 = pd.to_datetime(base.OPERDATE1, format="%d%m%Y%").dt.strftime('%d.%m.%Y')
        base.OPENDATE = pd.to_datetime(base.OPENDATE, format="%d%m%Y%").dt.strftime('%d.%m.%Y')
        base.CLOSEDATE = pd.to_datetime(base.CLOSEDATE, format="%d%m%Y%").dt.strftime('%d.%m.%Y')
        base.AMOUNTLOANKZT = ['{:,}'.format(amount).replace(',', ' ') for amount in base.AMOUNTLOANKZT]
        base.LOANRESTKZT = ['{:,}'.format(loan).replace(',', ' ') for loan in base.LOANRESTKZT]
        self.base_use = base
        self.base_use.columns = ['Отчетная дата', 'Филиал', 'Номер договора', 'ФИО', 'ИИН', 'Дата выдачи', 'Дата закрытия', 'Сумма выдачи', 'Остаток долга', 'Продукт', 'Просрочка']
        
        self.createTabDialog(base)
        
    def show_cell_vintage_1(self):
        itemCol = self.nd.tableWidget.currentItem().column()
        month = self.nd.tableWidget.item(self.nd.tableWidget.currentRow(), 0).text()
        months = {'January':'01', 'February':'02', 'March':'03', 'April':'04', 'May':'05', 'June':'06', 'July':'07', 'August':'08', 'September':'09','October':'10', 'November':'11', 'December':'12'}
        open_ym = int(month.split(' ')[1] + months[month.split(' ')[0]])
        base = self.df
        if self.com_prod.currentText() != 'Все продукты':
            base = base[base.PRODUCT == self.com_prod.currentText()]

        if self.com_fil.currentText() != 'Все филиалы':
            base = base[base.FILIAL == self.com_fil.currentText()]
            
        if itemCol >= 3:
            months_since_open = itemCol - 2
            base = base[(base.MONTHS_SINCE_OPEN == months_since_open)&(base.OPENDATE_YM == open_ym)&(base.PROSDAYS > 0)]
            
        if itemCol < 3:
            base = base.iloc[0:0]
            
        base = base[['OPERDATE1', 'FILIAL', 'CONTRNUM', 'CLIENTLABEL', 'IIN', 'OPENDATE', 'CLOSEDATE', 'AMOUNTLOANKZT', 'LOANRESTKZT', 'PRODUCT', 'PROSDAYS']]
        base.OPERDATE1 = pd.to_datetime(base.OPERDATE1, format="%d%m%Y%").dt.strftime('%d.%m.%Y')
        base.OPENDATE = pd.to_datetime(base.OPENDATE, format="%d%m%Y%").dt.strftime('%d.%m.%Y')
        base.CLOSEDATE = pd.to_datetime(base.CLOSEDATE, format="%d%m%Y%").dt.strftime('%d.%m.%Y')
        base.AMOUNTLOANKZT = ['{:,}'.format(amount).replace(',', ' ') for amount in base.AMOUNTLOANKZT]
        base.LOANRESTKZT = ['{:,}'.format(loan).replace(',', ' ') for loan in base.LOANRESTKZT]
        self.base_use = base
        self.base_use.columns = ['Отчетная дата', 'Филиал', 'Номер договора', 'ФИО', 'ИИН', 'Дата выдачи', 'Дата закрытия', 'Сумма выдачи', 'Остаток долга', 'Продукт', 'Просрочка']
        
        self.createTabDialog(base)
        
    def show_cell_vintage_30(self):
        itemCol = self.nd.tableWidget.currentItem().column()
        month = self.nd.tableWidget.item(self.nd.tableWidget.currentRow(), 0).text()
        months = {'January':'01', 'February':'02', 'March':'03', 'April':'04', 'May':'05', 'June':'06', 'July':'07', 'August':'08', 'September':'09','October':'10', 'November':'11', 'December':'12'}
        open_ym = int(month.split(' ')[1] + months[month.split(' ')[0]])
        base = self.df
        if self.com_prod.currentText() != 'Все продукты':
            base = base[base.PRODUCT == self.com_prod.currentText()]

        if self.com_fil.currentText() != 'Все филиалы':
            base = base[base.FILIAL == self.com_fil.currentText()]
            
        if itemCol >= 3:
            months_since_open = itemCol - 1
            base = base[(base.MONTHS_SINCE_OPEN == months_since_open)&(base.OPENDATE_YM == open_ym)&(base.PROSDAYS > 30)]
            
        if itemCol < 3:
            base = base.iloc[0:0]
            
        base = base[['OPERDATE1', 'FILIAL', 'CONTRNUM', 'CLIENTLABEL', 'IIN', 'OPENDATE', 'CLOSEDATE', 'AMOUNTLOANKZT', 'LOANRESTKZT', 'PRODUCT', 'PROSDAYS']]
        base.OPERDATE1 = pd.to_datetime(base.OPERDATE1, format="%d%m%Y%").dt.strftime('%d.%m.%Y')
        base.OPENDATE = pd.to_datetime(base.OPENDATE, format="%d%m%Y%").dt.strftime('%d.%m.%Y')
        base.CLOSEDATE = pd.to_datetime(base.CLOSEDATE, format="%d%m%Y%").dt.strftime('%d.%m.%Y')
        base.AMOUNTLOANKZT = ['{:,}'.format(amount).replace(',', ' ') for amount in base.AMOUNTLOANKZT]
        base.LOANRESTKZT = ['{:,}'.format(loan).replace(',', ' ') for loan in base.LOANRESTKZT]
        self.base_use = base
        self.base_use.columns = ['Отчетная дата', 'Филиал', 'Номер договора', 'ФИО', 'ИИН', 'Дата выдачи', 'Дата закрытия', 'Сумма выдачи', 'Остаток долга', 'Продукт', 'Просрочка']
        
        self.createTabDialog(base)
        
    def show_cell_vintage_90(self):
        itemCol = self.nd.tableWidget.currentItem().column()
        month = self.nd.tableWidget.item(self.nd.tableWidget.currentRow(), 0).text()
        months = {'January':'01', 'February':'02', 'March':'03', 'April':'04', 'May':'05', 'June':'06', 'July':'07', 'August':'08', 'September':'09','October':'10', 'November':'11', 'December':'12'}
        open_ym = int(month.split(' ')[1] + months[month.split(' ')[0]])
        base = self.df
        if self.com_prod.currentText() != 'Все продукты':
            base = base[base.PRODUCT == self.com_prod.currentText()]

        if self.com_fil.currentText() != 'Все филиалы':
            base = base[base.FILIAL == self.com_fil.currentText()]
            
        if itemCol >= 3:
            months_since_open = itemCol + 1
            base = base[(base.MONTHS_SINCE_OPEN == months_since_open)&(base.OPENDATE_YM == open_ym)&(base.PROSDAYS > 90)]
            
        if itemCol < 3:
            base = base.iloc[0:0]
            
        base = base[['OPERDATE1', 'FILIAL', 'CONTRNUM', 'CLIENTLABEL', 'IIN', 'OPENDATE', 'CLOSEDATE', 'AMOUNTLOANKZT', 'LOANRESTKZT', 'PRODUCT', 'PROSDAYS']]
        base.OPERDATE1 = pd.to_datetime(base.OPERDATE1, format="%d%m%Y%").dt.strftime('%d.%m.%Y')
        base.OPENDATE = pd.to_datetime(base.OPENDATE, format="%d%m%Y%").dt.strftime('%d.%m.%Y')
        base.CLOSEDATE = pd.to_datetime(base.CLOSEDATE, format="%d%m%Y%").dt.strftime('%d.%m.%Y')
        base.AMOUNTLOANKZT = ['{:,}'.format(amount).replace(',', ' ') for amount in base.AMOUNTLOANKZT]
        base.LOANRESTKZT = ['{:,}'.format(loan).replace(',', ' ') for loan in base.LOANRESTKZT]
        self.base_use = base
        self.base_use.columns = ['Отчетная дата', 'Филиал', 'Номер договора', 'ФИО', 'ИИН', 'Дата выдачи', 'Дата закрытия', 'Сумма выдачи', 'Остаток долга', 'Продукт', 'Просрочка']
        
        self.createTabDialog(base)
        
    def show_cell_vintage_180(self):
        itemCol = self.nd.tableWidget.currentItem().column()
        month = self.nd.tableWidget.item(self.nd.tableWidget.currentRow(), 0).text()
        months = {'January':'01', 'February':'02', 'March':'03', 'April':'04', 'May':'05', 'June':'06', 'July':'07', 'August':'08', 'September':'09','October':'10', 'November':'11', 'December':'12'}
        open_ym = int(month.split(' ')[1] + months[month.split(' ')[0]])
        base = self.df
        if self.com_prod.currentText() != 'Все продукты':
            base = base[base.PRODUCT == self.com_prod.currentText()]

        if self.com_fil.currentText() != 'Все филиалы':
            base = base[base.FILIAL == self.com_fil.currentText()]
            
        if itemCol >= 3:
            months_since_open = itemCol + 4
            base = base[(base.MONTHS_SINCE_OPEN == months_since_open)&(base.OPENDATE_YM == open_ym)&(base.PROSDAYS > 180)]
            
        if itemCol < 3:
            base = base.iloc[0:0]
            
        base = base[['OPERDATE1', 'FILIAL', 'CONTRNUM', 'CLIENTLABEL', 'IIN', 'OPENDATE', 'CLOSEDATE', 'AMOUNTLOANKZT', 'LOANRESTKZT', 'PRODUCT', 'PROSDAYS']]
        base.OPERDATE1 = pd.to_datetime(base.OPERDATE1, format="%d%m%Y%").dt.strftime('%d.%m.%Y')
        base.OPENDATE = pd.to_datetime(base.OPENDATE, format="%d%m%Y%").dt.strftime('%d.%m.%Y')
        base.CLOSEDATE = pd.to_datetime(base.CLOSEDATE, format="%d%m%Y%").dt.strftime('%d.%m.%Y')
        base.AMOUNTLOANKZT = ['{:,}'.format(amount).replace(',', ' ') for amount in base.AMOUNTLOANKZT]
        base.LOANRESTKZT = ['{:,}'.format(loan).replace(',', ' ') for loan in base.LOANRESTKZT]
        self.base_use = base
        self.base_use.columns = ['Отчетная дата', 'Филиал', 'Номер договора', 'ФИО', 'ИИН', 'Дата выдачи', 'Дата закрытия', 'Сумма выдачи', 'Остаток долга', 'Продукт', 'Просрочка']
        
        self.createTabDialog(base)

    def show_contracts_vintage_mis(self):
        contract = self.td.tableWidget.item(self.td.tableWidget.currentRow(), 2).text()
        try:
            base_dwh = self.dwh
            s = base_dwh.loc[contract]
            self.dd = DwhDialog(self)
            self.dd.setGeometry(300, 300, 300, 300)
            
            self.dd.lbl_fio = QLabel('ФИО:')
            self.dd.lbl_set_fio = QLabel(s.SURNAME + ' ' + s.FIRST_NAME)
            
            self.dd.lbl_birthday = QLabel('Дата рождения:')
            self.dd.lbl_set_birthday = QLabel(str(s.BIRTHDAY)[:10])
            
            self.dd.lbl_status = QLabel('Семейный статус:')
            self.dd.lbl_set_status = QLabel(s.MARITAL_STATUS_CODE)
            
            self.dd.lbl_birthplace = QLabel('Место рождения:')
            self.dd.lbl_set_birthplace = QLabel(s.BIRTH_ADDRESS_NAME)
            
            self.dd.lbl_edu = QLabel('Образование:')
            self.dd.lbl_set_edu = QLabel(s.ED_TYPE_NAME)
            
            self.dd.lbl_totalwork = QLabel('Общий стаж работы:')
            self.dd.lbl_set_totalwork = QLabel(str(int(s.WA_TOTAL_LENGTH_OF_WORK)) + ' (в мес.)')
            
            self.dd.lbl_lastwork = QLabel('Стаж на посл. месте:')
            self.dd.lbl_set_lastwork = QLabel(str(int(s.WA_LENGTH_OF_WORK)) + ' (в мес.)')
            
            self.dd.lbl_sphere = QLabel('Сфера организации:')
            self.dd.lbl_set_sphere = QLabel(s.WA_ORGANIZATION_ACTIVITY_CODE)
            
            self.dd.lbl_orgname = QLabel('Наименование организации:')
            self.dd.lbl_set_orgname = QLabel(s.WA_ORGANIZATION_NAME)
            
            self.dd.lbl_spec = QLabel('Специальность:')
            self.dd.lbl_set_spec = QLabel(s.WA_POSITION_CODE)
        
            self.dd.lbl_manager = QLabel('Кредитный менеджер:')
            self.dd.lbl_set_manager = QLabel(s.LOAN_MANAGER)
            
            self.dd.lbl_reqamount = QLabel('Запрошенная сумма:')
            self.dd.lbl_set_reqamount = QLabel()
            if str(s.REQ_AMOUNT) != 'nan':
                self.dd.lbl_set_reqamount.setText('{:,}'.format(int(s.REQ_AMOUNT)).replace(',', ' '))
            
            self.dd.lbl_reqterm = QLabel('Запрошенный срок:')
            self.dd.lbl_set_reqterm = QLabel()
            if str(s.REQ_TERM) != 'nan':
                self.dd.lbl_set_reqterm.setText(str(int(s.REQ_TERM)))
            
            self.dd.lbl_appamount = QLabel('Выбранная сумма:')
            self.dd.lbl_set_appamount = QLabel()
            if str(s.APP_AMOUNT) != 'nan':
                self.dd.lbl_set_appamount.setText('{:,}'.format(int(s.APP_AMOUNT)).replace(',', ' '))
            
            self.dd.lbl_appterm = QLabel('Выбранный срок:')
            self.dd.lbl_set_appterm = QLabel()
            if str(s.APP_TERM) != 'nan':
                self.dd.lbl_set_appterm.setText(str(int(s.APP_TERM)))
            
            self.dd.lbl_zp = QLabel('Рассчитанная ЗП:')
            self.dd.lbl_set_zp = QLabel()
            if str(s.MONTHLY_INCOME_CLIENT) !='nan':
                self.dd.lbl_set_zp.setText('{:,}'.format(int(s.MONTHLY_INCOME_CLIENT)).replace(',', ' '))
            
            self.dd.lbl_riskcat = QLabel('Риск категория:')
            self.dd.lbl_set_riskcat = QLabel()
            if str(s.RISK_CATEGORY) != 'nan':
                self.dd.lbl_set_riskcat.setText(str(int(s.RISK_CATEGORY)))
            
            self.dd.lbl_clientcat = QLabel('Категория клиента:')
            self.dd.lbl_set_clientcat = QLabel()
            if str(s.CLIENT_CATEGORY) != 'nan':
                self.dd.lbl_set_clientcat.setText(str(s.CLIENT_CATEGORY))

            self.dd.lbl_score = QLabel('Скорбалл:')
            self.dd.lbl_set_score = QLabel()
            if str(s.SCOREPOINT_VALUE) != 'nan':
                self.dd.lbl_set_score.setText(str(s.SCOREPOINT_VALUE))
            
            self.dd.lbl_code = QLabel('Отказ:')
            self.dd.lbl_set_code = QLabel()
            if str(s.DESCRIPTION_FOR_MANAGER) != 'nan':
                self.dd.lbl_set_code.setText(str(s.DESCRIPTION_FOR_MANAGER))
            
            self.dd.layout = QGridLayout()
            self.dd.layout.addWidget(self.dd.lbl_fio, 0, 0)
            self.dd.layout.addWidget(self.dd.lbl_set_fio, 0, 1)
            
            self.dd.layout.addWidget(self.dd.lbl_birthday, 1, 0)
            self.dd.layout.addWidget(self.dd.lbl_set_birthday, 1, 1)
            
            self.dd.layout.addWidget(self.dd.lbl_status, 2, 0)
            self.dd.layout.addWidget(self.dd.lbl_set_status, 2, 1)
            
            self.dd.layout.addWidget(self.dd.lbl_birthplace, 3, 0)
            self.dd.layout.addWidget(self.dd.lbl_set_birthplace, 3, 1)
            
            self.dd.layout.addWidget(self.dd.lbl_edu, 4, 0)
            self.dd.layout.addWidget(self.dd.lbl_set_edu, 4, 1)
            
            self.dd.layout.addWidget(self.dd.lbl_totalwork, 5, 0)
            self.dd.layout.addWidget(self.dd.lbl_set_totalwork, 5, 1)
            
            self.dd.layout.addWidget(self.dd.lbl_lastwork, 6, 0)
            self.dd.layout.addWidget(self.dd.lbl_set_lastwork, 6, 1)
            
            self.dd.layout.addWidget(self.dd.lbl_sphere, 7, 0)
            self.dd.layout.addWidget(self.dd.lbl_set_sphere, 7, 1)
            
            self.dd.layout.addWidget(self.dd.lbl_orgname, 8, 0)
            self.dd.layout.addWidget(self.dd.lbl_set_orgname, 8, 1)
            
            self.dd.layout.addWidget(self.dd.lbl_spec, 9, 0)
            self.dd.layout.addWidget(self.dd.lbl_set_spec, 9, 1)
            
            self.dd.layout.addWidget(self.dd.lbl_manager, 10, 0)
            self.dd.layout.addWidget(self.dd.lbl_set_manager, 10, 1)
            
            self.dd.layout.addWidget(self.dd.lbl_reqamount, 11, 0)
            self.dd.layout.addWidget(self.dd.lbl_set_reqamount, 11, 1)
            
            self.dd.layout.addWidget(self.dd.lbl_reqterm, 12, 0)
            self.dd.layout.addWidget(self.dd.lbl_set_reqterm, 12, 1)
            
            self.dd.layout.addWidget(self.dd.lbl_appamount, 13, 0)
            self.dd.layout.addWidget(self.dd.lbl_set_appamount, 13, 1)
            
            self.dd.layout.addWidget(self.dd.lbl_appterm, 14, 0)
            self.dd.layout.addWidget(self.dd.lbl_set_appterm, 14, 1)
            
            self.dd.layout.addWidget(self.dd.lbl_zp, 15, 0)
            self.dd.layout.addWidget(self.dd.lbl_set_zp, 15, 1)
            
            self.dd.layout.addWidget(self.dd.lbl_riskcat, 16, 0)
            self.dd.layout.addWidget(self.dd.lbl_set_riskcat, 16, 1)
            
            self.dd.layout.addWidget(self.dd.lbl_clientcat, 17, 0)
            self.dd.layout.addWidget(self.dd.lbl_set_clientcat, 17, 1)
            
            self.dd.layout.addWidget(self.dd.lbl_score, 18, 0)
            self.dd.layout.addWidget(self.dd.lbl_set_score, 18, 1)
            
            self.dd.layout.addWidget(self.dd.lbl_code, 19, 0)
            self.dd.layout.addWidget(self.dd.lbl_set_code, 19, 1)
            
            self.dd.setLayout(self.dd.layout)
            self.dd.show()
            
        except Exception:
            self.dd = DwhDialog(self)
            self.dd.setGeometry(300, 300, 300, 300)
            
            self.dd.lbl_error = QLabel('Данные по данному контракту не найдены')
            self.dd.layout = QVBoxLayout()
            self.dd.layout.addWidget(self.dd.lbl_error)
            
            self.dd.setLayout(self.dd.layout)
            self.dd.show()
            
    def basket_report(self):
        excel_file = QFileDialog.getSaveFileName(self, 'Сохранить файл', 'Анализ по корзинам', '*.xlsx')[0]
        if str(excel_file):
            all_rp = self.df
            basket_an.createReport(all_rp, excel_file)
            
    def dpk_report_(self):
        excel_file = QFileDialog.getSaveFileName(self, 'Сохранить файл', 'Список для ДПК', '*.xlsx')[0]
        if str(excel_file):
            all_rp = self.df
            dpk_report.createReport(all_rp, excel_file)
            
    def createGraph(self):
        base = self.df
        if self.com_prod.currentText() != 'Все продукты':
            base = base[base.PRODUCT == self.com_prod.currentText()]

        if self.com_fil.currentText() != 'Все филиалы':
            base = base[base.FILIAL == self.com_fil.currentText()]
            
        if self.com_type.currentText() == 'Vintage MIS':
            base2 = vintage_mis.createTable(base)
            fig = make_subplots(specs=[[{"secondary_y": True}]])
            fig.add_trace(go.Bar(x=base2['Month of Issue'], y=base2['Amount Loan'], name="Amount Loan"), secondary_y=False,)
            fig.add_trace(go.Scatter(x=base2['Month of Issue'], y=base2['1+@1MOB'], name="1+@1MOB"), secondary_y=True,)
            fig.add_trace(go.Scatter(x=base2['Month of Issue'], y=base2['30+@3MOB'], name="30+@3MOB"), secondary_y=True,)
            fig.add_trace(go.Scatter(x=base2['Month of Issue'], y=base2['90+@6MOB'], name="90+@6MOB"), secondary_y=True,)
            fig.update_xaxes(title_text="Month of Issue")
            fig.update_yaxes(title_text="Amount Loan", secondary_y=False)
            fig.update_yaxes(title_text="Share, %", secondary_y=True)
            
            plo.plot(fig)
            
        if self.com_type.currentText() == 'Vintage Amount 1+':
            base2 = vintage.vintage_amount(1, 0, base)
            base2 = base2.set_index('Month Issue')
            xcols = base2.columns[2:]
            data = []
            a = []
            for i in base2.index:
                for xcol in xcols:
                    a.append(base2.loc[i][xcol])
                data.append(go.Scatter(x=xcols, y=a, name=i))   
                a = []
            fig = go.Figure(data = data)
            
            plo.plot(fig)
    
        if self.com_type.currentText() == 'Vintage Amount 30+':
            base2 = vintage.vintage_amount(2, 30, base)
            base2 = base2.set_index('Month Issue')
            xcols = base2.columns[2:]
            data = []
            a = []
            for i in base2.index:
                for xcol in xcols:
                    a.append(base2.loc[i][xcol])
                data.append(go.Scatter(x=xcols, y=a, name=i))   
                a = []
            fig = go.Figure(data = data)
            
            plo.plot(fig)
            
        if self.com_type.currentText() == 'Vintage Amount 90+':
            base2 = vintage.vintage_amount(4, 90, base)
            base2 = base2.set_index('Month Issue')
            xcols = base2.columns[2:]
            data = []
            a = []
            for i in base2.index:
                for xcol in xcols:
                    a.append(base2.loc[i][xcol])
                data.append(go.Scatter(x=xcols, y=a, name=i))   
                a = []
            fig = go.Figure(data = data)
            
            plo.plot(fig)
            
        if self.com_type.currentText() == 'Vintage Amount 180+':
            base2 = vintage.vintage_amount(7, 180, base)
            base2 = base2.set_index('Month Issue')
            xcols = base2.columns[2:]
            data = []
            a = []
            for i in base2.index:
                for xcol in xcols:
                    a.append(base2.loc[i][xcol])
                data.append(go.Scatter(x=xcols, y=a, name=i))   
                a = []
            fig = go.Figure(data = data)
            
            plo.plot(fig)
            
        if self.com_type.currentText() == 'Vintage Count 1+':
            base2 = vintage.vintage_count(1, 0, base)
            base2 = base2.set_index('Month Issue')
            xcols = base2.columns[2:]
            data = []
            a = []
            for i in base2.index:
                for xcol in xcols:
                    a.append(base2.loc[i][xcol])
                data.append(go.Scatter(x=xcols, y=a, name=i))   
                a = []
            fig = go.Figure(data = data)
            
            plo.plot(fig)
    
        if self.com_type.currentText() == 'Vintage Count 30+':
            base2 = vintage.vintage_count(2, 30, base)
            base2 = base2.set_index('Month Issue')
            xcols = base2.columns[2:]
            data = []
            a = []
            for i in base2.index:
                for xcol in xcols:
                    a.append(base2.loc[i][xcol])
                data.append(go.Scatter(x=xcols, y=a, name=i))   
                a = []
            fig = go.Figure(data = data)
            
            plo.plot(fig)
            
        if self.com_type.currentText() == 'Vintage Count 90+':
            base2 = vintage.vintage_count(4, 90, base)
            base2 = base2.set_index('Month Issue')
            xcols = base2.columns[2:]
            data = []
            a = []
            for i in base2.index:
                for xcol in xcols:
                    a.append(base2.loc[i][xcol])
                data.append(go.Scatter(x=xcols, y=a, name=i))   
                a = []
            fig = go.Figure(data = data)
            
            plo.plot(fig)
            
        if self.com_type.currentText() == 'Vintage Count 180+':
            base2 = vintage.vintage_count(7, 180, base)
            base2 = base2.set_index('Month Issue')
            xcols = base2.columns[2:]
            data = []
            a = []
            for i in base2.index:
                for xcol in xcols:
                    a.append(base2.loc[i][xcol])
                data.append(go.Scatter(x=xcols, y=a, name=i))   
                a = []
            fig = go.Figure(data = data)
            
            plo.plot(fig)
            
        if self.com_type.currentText() == 'FPD SPD TPD':
            base2 = fpd.fpd_data(base)
            fig = make_subplots(specs=[[{"secondary_y": True}]])
            fig.add_trace(go.Bar(x=base2['Месяц выдачи'], y=base2['Сумма выдачи, млн, KZT'], name="Сумма выдачи, млн, KZT"), secondary_y=False,)
            fig.add_trace(go.Scatter(x=base2['Месяц выдачи'], y=base2['Доля FPD, %'], name="Доля FPD, %"), secondary_y=True,)
            fig.add_trace(go.Scatter(x=base2['Месяц выдачи'], y=base2['Доля SPD, %'], name="Доля SPD, %"), secondary_y=True,)
            fig.add_trace(go.Scatter(x=base2['Месяц выдачи'], y=base2['Доля TPD, %'], name="Доля TPD, %"), secondary_y=True,)
            fig.add_trace(go.Scatter(x=base2['Месяц выдачи'], y=base2['Доля FDef, %'], name="Доля FDef, %"), secondary_y=True,)
            fig.add_trace(go.Scatter(x=base2['Месяц выдачи'], y=base2['Доля Loss, %'], name="Доля Loss, %"), secondary_y=True,)
            fig.update_xaxes(title_text="Месяц выдачи")
            fig.update_yaxes(title_text="Сумма выдачи, млн, KZT", secondary_y=False)
            fig.update_yaxes(title_text="Доля, %", secondary_y=True)
            
            plo.plot(fig)
    
    def createTabDialog(self, base):
        self.td = TabDialog(self)
        self.td.setGeometry(300, 300, 1100, 300)
        self.td.tableWidget = QTableWidget()
        self.td.btn_1 = QPushButton('Сохранить')
        self.td.btn_1.setFixedSize(120, 20)
        self.td.btn_2 = QPushButton('Добавить в корзину')
        self.td.btn_2.setFixedSize(120, 20)
        self.td.btn_2.clicked.connect(self.addToBucket)
        self.td.tableWidget.setRowCount(len(base.index))
        self.td.tableWidget.setColumnCount(len(base.columns))
        for i in range(len(base.index)):
                for j in range(len(base.columns)):
                    self.td.tableWidget.setItem(i, j, QTableWidgetItem(str(base.iat[i, j])))
        self.td.tableWidget.setHorizontalHeaderLabels(['Отчетная дата', 'Филиал', 'Номер договора', 'ФИО', 'ИИН', 'Дата выдачи', 'Дата закрытия', 'Сумма выдачи', 'Остаток долга', 'Продукт', 'Просрочка'])
        self.td.tableWidget.resizeColumnsToContents()
        self.td.layout = QVBoxLayout()
        self.td.layout.addWidget(self.td.tableWidget) 
        self.td.layout.addWidget(self.td.btn_1)
        self.td.layout.addWidget(self.td.btn_2)
        self.td.setLayout(self.td.layout)
        self.td.tableWidget.setEditTriggers(QTableWidget.NoEditTriggers)
        self.td.tableWidget.cellDoubleClicked.connect(self.show_contracts_vintage_mis)
        self.td.show()
        
    def createNewDialog(self, base):
        self.nd.tableWidget = QTableWidget()
        self.nd.tableWidget.setRowCount(len(base.index))
        self.nd.tableWidget.setColumnCount(len(base.columns))
        for i in range(len(base.index)):
            for j in range(len(base.columns)):
                self.nd.tableWidget.setItem(i, j, QTableWidgetItem(str(base.iat[i, j])))
        self.nd.tableWidget.move(10,70)
        self.nd.tableWidget.setHorizontalHeaderLabels(base.columns)
        self.nd.tableWidget.resizeColumnsToContents()
        self.nd.tableWidget.setEditTriggers(QTableWidget.NoEditTriggers)
        self.nd.layout = QVBoxLayout()
        self.nd.layout.addWidget(self.nd.tableWidget) 
        self.nd.setLayout(self.nd.layout)
        self.nd.show()
        
    def createBucketDialog(self):
        df = pd.concat(self.dfs, axis=0)
        self.excel_df = df
        self.bd = BucketDialog(self)
        self.bd.setGeometry(300, 300, 1100, 300)
        self.bd.tableWidget = QTableWidget()
        self.bd.btn_excel = QPushButton('В Excel')
        self.bd.btn_excel.setFixedSize(120, 20)
        self.bd.btn_excel.clicked.connect(self.toExcel)
        self.bd.tableWidget.setRowCount(len(df.index))
        self.bd.tableWidget.setColumnCount(len(df.columns))
        for i in range(len(df.index)):
            for j in range(len(df.columns)):
                self.bd.tableWidget.setItem(i, j, QTableWidgetItem(str(df.iat[i, j])))
        self.bd.tableWidget.setHorizontalHeaderLabels(df.columns)
        self.bd.tableWidget.resizeColumnsToContents()
        self.bd.tableWidget.setEditTriggers(QTableWidget.NoEditTriggers)
        self.bd.layout = QVBoxLayout()
        self.bd.layout.addWidget(self.bd.tableWidget) 
        self.bd.layout.addWidget(self.bd.btn_excel)
        self.bd.setLayout(self.bd.layout)
        self.bd.show()
        
    def calcScore(self):
        self.sd = ScoreDialog(self)
        self.sd.lbl_branch = QLabel('Филиал:')
        self.sd.com_branch = QComboBox()
        self.sd.com_branch.addItems(['Алматы', 'Шымкент', 'Нур-Султан', 'Павлодар', 'Караганда', 'Актау', 'Кокшетау', 'Каскелен', 'Тараз'])
        self.sd.com_branch.setFixedSize(120, 20)
        self.sd.lbl_age = QLabel('Возраст:')
        self.sd.line_age = QLineEdit()
        self.sd.line_age.setFixedSize(120, 20)
        self.sd.lbl_term = QLabel('Срок кредита:')
        self.sd.line_term = QLineEdit()
        self.sd.line_term.setFixedSize(120, 20)
        self.sd.lbl_kdd = QLabel('КДД:')
        self.sd.line_kdd = QLineEdit()
        self.sd.line_kdd.setFixedSize(120, 20)
        self.sd.lbl_stazh = QLabel('Стаж на посл. месте (в мес.):')
        self.sd.line_stazh = QLineEdit()
        self.sd.line_stazh.setFixedSize(120, 20)
        self.sd.lbl_numcredit = QLabel('Закрыто кредитов:')
        self.sd.line_numcredit = QLineEdit()
        self.sd.line_numcredit.setFixedSize(120, 20)
        self.sd.lbl_edu = QLabel('Образование:')
        self.sd.com_edu = QComboBox()
        self.sd.com_edu.addItems(['Среднее/не оконченное обр.', 'Средне специальное', 'Высшее/Ученая степень/MBA/Второе высшее'])
        self.sd.com_edu.setFixedSize(120, 20)
        self.sd.lbl_zp = QLabel('Получение ЗП:')
        self.sd.com_zp = QComboBox()
        self.sd.com_zp.addItems(['Через кассу предприятия', 'с других БВУ', 'Участник/сотрудник Tengri Bank'])
        self.sd.com_zp.setFixedSize(120, 20)
        self.sd.lbl_status = QLabel('Сем. положение:')
        self.sd.com_status = QComboBox()
        self.sd.com_status.setFixedSize(120, 20)
        self.sd.com_status.addItems(['Холост/незамужем', 'Женат/замужем', 'Гражданский брак', 'Разведен/разведена', 'Вдовец/вдова'])
        self.sd.lbl_gender = QLabel('Пол:')
        self.sd.com_gender = QComboBox()
        self.sd.com_gender.setFixedSize(120, 20)
        self.sd.com_gender.addItems(['Мужской', 'Женский'])
        self.sd.lbl_outstand = QLabel('Тек. просрочка:')
        self.sd.line_outstand = QLineEdit()
        self.sd.line_outstand.setFixedSize(120, 20)
        self.sd.btn_calc = QPushButton('Расчет')
        self.sd.btn_calc.setFixedSize(120, 20)
        self.sd.btn_calc.clicked.connect(self.btnCalcScore)
        self.sd.lbl_decision = QLabel('Решение:')
        self.sd.lbl_set_decision = QLabel()
        self.sd.lbl_scorepoint = QLabel('Скорбалл:')
        self.sd.lbl_set_scorepoint = QLabel()
        
        self.sd.grid = QGridLayout()
        self.sd.grid.addWidget(self.sd.lbl_branch, 0, 0)
        self.sd.grid.addWidget(self.sd.com_branch, 0, 1)
        self.sd.grid.addWidget(self.sd.lbl_age, 1, 0)
        self.sd.grid.addWidget(self.sd.line_age, 1, 1)
        self.sd.grid.addWidget(self.sd.lbl_term, 2, 0)
        self.sd.grid.addWidget(self.sd.line_term, 2, 1)
        self.sd.grid.addWidget(self.sd.lbl_kdd, 3, 0)
        self.sd.grid.addWidget(self.sd.line_kdd, 3, 1)
        self.sd.grid.addWidget(self.sd.lbl_stazh, 4, 0)
        self.sd.grid.addWidget(self.sd.line_stazh, 4, 1)
        self.sd.grid.addWidget(self.sd.lbl_numcredit, 5, 0)
        self.sd.grid.addWidget(self.sd.line_numcredit, 5, 1)
        self.sd.grid.addWidget(self.sd.lbl_edu, 6, 0)
        self.sd.grid.addWidget(self.sd.com_edu, 6, 1)
        self.sd.grid.addWidget(self.sd.lbl_zp, 7, 0)
        self.sd.grid.addWidget(self.sd.com_zp, 7, 1)
        self.sd.grid.addWidget(self.sd.lbl_status, 8, 0)
        self.sd.grid.addWidget(self.sd.com_status, 8, 1)
        self.sd.grid.addWidget(self.sd.lbl_gender, 9, 0)
        self.sd.grid.addWidget(self.sd.com_gender, 9, 1)
        self.sd.grid.addWidget(self.sd.lbl_outstand, 10, 0)
        self.sd.grid.addWidget(self.sd.line_outstand, 10, 1)
        self.sd.grid.addWidget(self.sd.btn_calc, 11, 0)
        self.sd.grid.addWidget(self.sd.lbl_decision, 12, 0)
        self.sd.grid.addWidget(self.sd.lbl_set_decision, 12, 1)
        self.sd.grid.addWidget(self.sd.lbl_scorepoint, 13, 0)
        self.sd.grid.addWidget(self.sd.lbl_set_scorepoint, 13, 1)
        self.sd.setLayout(self.sd.grid)
        self.sd.show()
        
    def predictModel(self):
        self.kd = KrishaDialog(self)
        self.kd.lbl_year = QLabel('Год постройки:')
        self.kd.line_year = QLineEdit()
        self.kd.line_year.setFixedSize(120, 20)
        self.kd.lbl_rooms = QLabel('Кол-во комнат:')
        self.kd.com_rooms = QComboBox()
        self.kd.com_rooms.setFixedSize(120, 20)
        self.kd.com_rooms.addItems(['1', '2', '3', '4', '5', '6', '7', '8', '9'])
        self.kd.lbl_floor = QLabel('Этаж:')
        self.kd.line_floor = QLineEdit()
        self.kd.line_floor.setFixedSize(120, 20)
        self.kd.lbl_maxfloor = QLabel('Кол-во этажей:')
        self.kd.line_maxfloor = QLineEdit()
        self.kd.line_maxfloor.setFixedSize(120, 20)
        self.kd.lbl_area = QLabel('Площадь:')
        self.kd.line_area = QLineEdit()
        self.kd.line_area.setFixedSize(120, 20)
        self.kd.lbl_type = QLabel('Тип строения:')
        self.kd.com_type = QComboBox()
        self.kd.com_type.setFixedSize(120, 20)
        self.kd.com_type.addItems(['кирпичный', 'монолитный', 'панельный'])
        self.kd.lbl_district = QLabel('Район:')
        self.kd.com_district = QComboBox()
        self.kd.com_district.setFixedSize(120, 20)
        self.kd.com_district.addItems(['Алатауский', 'Алмалинский', 'Ауэзовский', 'Бостандыкский', 'Жетысуский', 'Медеуский', 'Наурызбайский', 'Турксибский'])
        self.kd.btn_predict = QPushButton('Узнать цену')
        self.kd.btn_predict.setFixedSize(120, 20)
        self.kd.btn_predict.clicked.connect(self.btnPredictModel)
        self.kd.lbl_price = QLabel('Цена:')
        self.kd.lbl_set_price = QLabel()
        
        self.kd.grid = QGridLayout()
        self.kd.grid.addWidget(self.kd.lbl_year, 0, 0)
        self.kd.grid.addWidget(self.kd.line_year, 0, 1)
        self.kd.grid.addWidget(self.kd.lbl_rooms, 1, 0)
        self.kd.grid.addWidget(self.kd.com_rooms, 1, 1)
        self.kd.grid.addWidget(self.kd.lbl_floor, 2, 0)
        self.kd.grid.addWidget(self.kd.line_floor, 2, 1)
        self.kd.grid.addWidget(self.kd.lbl_maxfloor, 3, 0)
        self.kd.grid.addWidget(self.kd.line_maxfloor, 3, 1)
        self.kd.grid.addWidget(self.kd.lbl_area, 4, 0)
        self.kd.grid.addWidget(self.kd.line_area, 4, 1)
        self.kd.grid.addWidget(self.kd.lbl_type, 5, 0)
        self.kd.grid.addWidget(self.kd.com_type, 5, 1)
        self.kd.grid.addWidget(self.kd.lbl_district, 6, 0)
        self.kd.grid.addWidget(self.kd.com_district, 6, 1)
        self.kd.grid.addWidget(self.kd.btn_predict, 7, 0)
        self.kd.grid.addWidget(self.kd.lbl_price, 8, 0)
        self.kd.grid.addWidget(self.kd.lbl_set_price, 8, 1)
        self.kd.setLayout(self.kd.grid)
        self.kd.show()
        
    def btnCalcScore(self):
        if (self.sd.line_age.text()) and (self.sd.line_term.text()) and (self.sd.line_kdd.text()) and (self.sd.line_stazh.text()) and (self.sd.line_numcredit.text()) and (self.sd.line_outstand.text()):
            (total_score, decision) = calcscore.calcScorePoint(self.sd.com_branch.currentText(), int(self.sd.line_age.text()), int(self.sd.line_term.text()), float(self.sd.line_kdd.text()), 
                                                               int(self.sd.line_stazh.text()), int(self.sd.line_numcredit.text()), self.sd.com_edu.currentText(), self.sd.com_zp.currentText(), 
                                                               self.sd.com_status.currentText(), self.sd.com_gender.currentText(), int(self.sd.line_outstand.text()))
            self.sd.lbl_set_decision.setText(str(decision))
            self.sd.lbl_set_scorepoint.setText(str(total_score))
        else:
            QMessageBox.information(self, 'Ошибка', 'Необходимо заполнить все поля')
            
    def btnPredictModel(self):
        if (self.kd.line_year.text()) and (self.kd.line_floor.text()) and (self.kd.line_maxfloor.text()) and (self.kd.line_area.text()):
            price = krisha_model.predictPriceModel(self.model, int(self.kd.line_year.text()), int(self.kd.com_rooms.currentText()), int(self.kd.line_floor.text()), int(self.kd.line_maxfloor.text()), 
                                                    int(self.kd.line_area.text()), self.kd.com_type.currentText(), self.kd.com_district.currentText())
            self.kd.lbl_set_price.setText('{:,}'.format(int(price/1000)*1000).replace(',', ' '))           
        else:
            QMessageBox.information(self, 'Ошибка', 'Необходимо заполнить все поля')

    def aboutDialog(self):
        QMessageBox.aboutQt(self)
        
    def addToBucket(self):
        self.dfs.append(self.base_use)
        
    def toExcel(self):
        excel_path = QFileDialog.getSaveFileName(self, 'Сохранить файл', 'Список', '*.xlsx')[0]
        if excel_path:
            self.excel_df.to_excel(excel_path, index=False)
            
class NewDialog(QWidget):
    
    def __init__(self, parent):
        super(NewDialog, self).__init__(parent)
        
class TabDialog(QDialog):
    
    def __init__(self, parent):
        super(TabDialog, self).__init__(parent)
        
class DwhDialog(QDialog):
    
    def __init__(self, parent):
        super(DwhDialog, self).__init__(parent)

class KrishaDialog(QDialog):
    
    def __init__(self, parent):
        super(KrishaDialog, self).__init__(parent)

class ScoreDialog(QDialog):
    
    def __init__(self, parent):
        super(ScoreDialog, self).__init__(parent)
        
class BucketDialog(QDialog):
    
    def __init__(self, parent):
        super(BucketDialog, self).__init__(parent)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = App()
    sys.exit(app.exec_())