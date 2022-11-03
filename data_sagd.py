import traceback

from PyQt5.QtWidgets import QApplication, QFileDialog, QCheckBox, QTableWidgetItem
import datetime
import os, sys
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns

from SAGD_dialog import *

app = QtWidgets.QApplication(sys.argv)
MainWindow = QtWidgets.QMainWindow()
ui = Ui_MainWindow()
ui.setupUi(MainWindow)
MainWindow.show()

# todo необходим рефакторинг


def choose_file():
    global months
    ui.lineEdit_file.clear()
    ui.comboBox_param.clear()
    ui.comboBox_year.clear()
    file_name = QFileDialog.getOpenFileName(filter='(*.xls *.xlsx)')
    ui.lineEdit_file.setText(file_name[0])
    db = pd.read_excel(file_name[0], sheet_name=None, header=None, index_col=None)
    zak_sum = list()
    is_year = 1985
    is_month = 13
    months = list()

    for skv in db:
        if len(skv) > 5:
            tab = db.get(skv)
            r = ''
            for row in range(len(tab.index)):
                for column in range(len(tab.columns)):
                    if tab.iat[row, column] == 'Дата':
                        r = row
                        break
                if r != '':
                    break
            tab = tab.drop(index=range(0, r))
            tab = tab.dropna(how='all').reset_index()
            del tab['index']
            list_date = list()
            for n, i in enumerate(tab.iloc[0].tolist()):
                if i == 'Дата':
                    list_date.append(n)
            ui.label_between.setText(str(list_date[1] - list_date[0]))

            for n, i in enumerate(tab.iloc[0]):
                if isinstance(i, str):
                    if i.startswith('Зак'):
                        zak_sum.append(n)
                    if i not in ['Дата', 'Время работы']:
                        ui.comboBox_param.addItem(str(n) + '. ' + i)
                    if n == list_date[1]:
                        ui.comboBox_param.addItem('{}/{}. Закачка суммарная'.format(str(zak_sum[0]), str(zak_sum[1])))
                        break
            for i in tab[list_date[0]]:
                if isinstance(i, datetime.datetime):
                    is_month = i.month
                    is_year = i.year
                    ui.comboBox_year.addItem(str(i.year))

                    break
            for i in tab[list_date[0]]:
                if isinstance(i, datetime.datetime):
                    if i.month != is_month:
                        months.append('{}/{}'.format(str(is_month), str(is_year)))
                        is_month = i.month
                    if i.year != is_year:
                        ui.comboBox_year.addItem(str(i.year))
                        is_year = i.year
            months.append('{}/{}'.format(str(is_month), str(is_year)))
            break


def calculate_parameter():
    data_year = pd.read_excel(ui.lineEdit_file.text(), sheet_name=None, header=None, index_col=None)
    file_name = ui.lineEdit_file.text().split('/')[-1]
    param_name = ui.comboBox_param.currentText().split('. ')[1]
    if param_name == 'Закачка суммарная':
        param = ui.comboBox_param.currentText().split('. ')[0].split('/')
    else:
        param = int(ui.comboBox_param.currentText().split('. ')[0])
    date_year = int(ui.comboBox_year.currentText())
    between_column_skv = int(ui.label_between.text())
    if ui.checkBox_all_years.isChecked():
        list_months = months
    else:
        list_months = ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь', 'июль', 'август', 'сентябрь', 'октябрь', 'ноябрь',
               'декабрь']
    result_tab = pd.DataFrame(index=list_months)

    for skv in data_year:
        db_skv = data_year.get(skv)
        i = 0
        while i < 500:
            if not isinstance(db_skv[1][i], datetime.datetime):
                i += 1
            else:
                break
        db_skv = db_skv.drop(index=range(0, i))
        data_year_by_month_skv1 = []
        data_year_by_month_skv2 = []
        month_param_skv1 = []
        month_param_skv2 = []
        month_data = 1

        if ui.checkBox_all_years.isChecked():
            month_data = int(list_months[0].split('/')[0])
        else:
            for data_day in db_skv.iterrows():
                if isinstance(data_day[1][1], datetime.datetime):
                    if (data_day[1][1]).year == date_year:
                        month_data = (data_day[1][1]).month
                        break
            if month_data > 1:
                data_year_by_month_skv1.extend([None] * (month_data - 1))
                if len(skv) > 5:
                    data_year_by_month_skv2.extend([None] * (month_data - 1))

        for data_day in db_skv.iterrows():
            if isinstance(data_day[1][1], datetime.datetime):
                if ui.checkBox_all_years.isChecked():
                    if (data_day[1][1]).month == month_data:
                        if param_name == 'Закачка суммарная':
                            month_param_skv1.append(data_day[1][int(param[0])] + data_day[1][int(param[1])])
                            if len(skv) > 5:
                                month_param_skv2.append(data_day[1][int(param[0]) + between_column_skv] +
                                                        data_day[1][int(param[1]) + between_column_skv])
                        else:
                            if data_day[1][param] == 0:  # для корректного расчета средних значений температуры,
                                data_day[1][param] = None  # если часть месяца 0
                            month_param_skv1.append(data_day[1][param])
                            if len(skv) > 5:
                                if data_day[1][param + between_column_skv] == 0:
                                    data_day[1][param + between_column_skv] = None
                                month_param_skv2.append(data_day[1][param + between_column_skv])
                    else:
                        if param_name.startswith(('Доб', 'Зак')):
                            data_year_by_month_skv1.append(np.sum(list(filter(None, month_param_skv1))))
                            # удаление None из списка перед расчетом суммы и среднего
                            if len(skv) > 5:
                                data_year_by_month_skv2.append(np.sum(list(filter(None, month_param_skv2))))
                        else:
                            data_year_by_month_skv1.append(np.mean(list(filter(None, month_param_skv1))))
                            if len(skv) > 5:
                                data_year_by_month_skv2.append(np.mean(list(filter(None, month_param_skv2))))
                        month_data += 1
                        if month_data == 13:
                            month_data = 1
                        if param_name == 'Закачка суммарная':
                            month_param_skv1 = [data_day[1][int(param[0])] + data_day[1][int(param[1])]]
                            if len(skv) > 5:
                                month_param_skv2 = [data_day[1][int(param[0]) + between_column_skv] +
                                                    data_day[1][int(param[1]) + between_column_skv]]
                        else:
                            if data_day[1][param] == 0:
                                data_day[1][param] = None
                            month_param_skv1 = [data_day[1][param]]
                            if len(skv) > 5:
                                if data_day[1][param + between_column_skv] == 0:
                                    data_day[1][param + between_column_skv] = None
                                month_param_skv2 = [data_day[1][param + between_column_skv]]
                else:
                    if (data_day[1][1]).year == date_year:
                        if (data_day[1][1]).month == month_data:
                            if param_name == 'Закачка суммарная':
                                month_param_skv1.append(data_day[1][int(param[0])] + data_day[1][int(param[1])])
                                if len(skv) > 5:
                                    month_param_skv2.append(data_day[1][int(param[0]) + between_column_skv] +
                                                            data_day[1][int(param[1]) + between_column_skv])
                            else:
                                if data_day[1][param] == 0:  # для корректного расчета средних значений температуры,
                                    data_day[1][param] = None  # если часть месяца 0
                                month_param_skv1.append(data_day[1][param])
                                if len(skv) > 5:
                                    if data_day[1][param + between_column_skv] == 0:
                                        data_day[1][param + between_column_skv] = None
                                    month_param_skv2.append(data_day[1][param + between_column_skv])
                        else:
                            if param_name.startswith(('Доб', 'Зак')):
                                data_year_by_month_skv1.append(np.sum(list(filter(None, month_param_skv1))))
                                # удаление None из списка перед расчетом суммы и среднего
                                if len(skv) > 5:
                                    data_year_by_month_skv2.append(np.sum(list(filter(None, month_param_skv2))))
                            else:
                                data_year_by_month_skv1.append(np.mean(list(filter(None, month_param_skv1))))
                                if len(skv) > 5:
                                    data_year_by_month_skv2.append(np.mean(list(filter(None, month_param_skv2))))
                            month_data += 1
                            if param_name == 'Закачка суммарная':
                                month_param_skv1 = [data_day[1][int(param[0])] + data_day[1][int(param[1])]]
                                if len(skv) > 5:
                                    month_param_skv2 = [data_day[1][int(param[0]) + between_column_skv] +
                                                        data_day[1][int(param[1]) + between_column_skv]]
                            else:
                                if data_day[1][param] == 0:
                                    data_day[1][param] = None
                                month_param_skv1 = [data_day[1][param]]
                                if len(skv) > 5:
                                    if data_day[1][param + between_column_skv] == 0:
                                        data_day[1][param + between_column_skv] = None
                                    month_param_skv2 = [data_day[1][param + between_column_skv]]
            else:
                break
        if param_name.startswith(('Доб', 'Зак')):
            data_year_by_month_skv1.append(np.sum(list(filter(None, month_param_skv1))))
            if len(skv) > 5:
                data_year_by_month_skv2.append(np.sum(list(filter(None, month_param_skv2))))
        else:
            data_year_by_month_skv1.append(np.mean(list(filter(None, month_param_skv1))))
            if len(skv) > 5:
                data_year_by_month_skv2.append(np.mean(list(filter(None, month_param_skv2))))
        if not ui.checkBox_all_years.isChecked():
            if month_data < 12:
                data_year_by_month_skv1.extend([None] * (12 - month_data))
                if len(skv) > 5:
                    data_year_by_month_skv2.extend([None] * (12 - month_data))
        result_tab[skv[0:5]] = data_year_by_month_skv1
        if len(skv) > 5:
            result_tab[skv[-5:]] = data_year_by_month_skv2
    print(result_tab)
    if ui.checkBox_all_years.isChecked():
        year_for_title = ''
    else:
        year_for_title = '_' + str(date_year)
    if ui.checkBox_save.isChecked():
        result_tab.to_excel(ui.lineEdit_file.text()[:-4] + '_' + param_name + year_for_title + '.xlsx')
    plt.figure(figsize=(len(result_tab.index)*0.3, len(result_tab.columns)*0.17), dpi=80)
    sns.heatmap(result_tab.transpose(), linewidths=.5, linecolor='black', cmap='jet', yticklabels=1)
    plt.title((param_name + '_' + file_name)[:-4] + year_for_title)
    plt.tight_layout()
    plt.show()


def log_uncaught_exceptions(ex_cls, ex, tb):
    # вывод ошибок
    text = '{}: {}:\n'.format(ex_cls.__name__, ex)
    text += ''.join(traceback.format_tb(tb))
    print(text)
    QtWidgets.QMessageBox.critical(None, 'Error', text)
    sys.exit()


sys.excepthook = log_uncaught_exceptions

ui.pushButton_file.clicked.connect(choose_file)
ui.pushButton_calc.clicked.connect(calculate_parameter)

sys.exit(app.exec_())