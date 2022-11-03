import datetime
from tkinter import *
from tkinter import filedialog as fd
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns

root = Tk()
root.title("Анализ параметров разработки залежей СВН")
root.geometry("550x205")
root.resizable(False, False)


def new_file():
    name_file_excel.delete(0, END)
    name_file = fd.askopenfilename(filetypes=(("xls files", "*.xls"), ("all files", "*.*")))
    name_file_excel.insert(0, name_file)
    name_file_excel.focus()


def calculate_param():
    data_year = pd.read_excel(name_file_excel.get(), sheet_name=None, header=None, index_col=None)
    file_name = name_file_excel.get().split('/')[-1]

    if param_value.get() == 't_ust':
        param, param_name = 2, 'темп_на_устье'
    elif param_value.get() == 't_nas':
        param, param_name = 3, 'темп_на_насосе'
    elif param_value.get() == 't_mas':
        param, param_name = 15, 'темп_на_массомере'
    elif param_value.get() == 'p_mas':
        param, param_name = 16, 'плотн_на_массомере'
    elif param_value.get() == 'dob_j':
        param, param_name = 4, 'добыча_жидкости'
    elif param_value.get() == 'dob_n':
        param, param_name = 7, 'добыча_нефти'
    elif param_value.get() == 'obv':
        param, param_name = 13, 'обводненность'
    elif param_value.get() == 'obv_p':
        param, param_name = 17, 'обводн_по_плотн'
    elif param_value.get() == 'zak_n':
        param, param_name = 5, 'закачка_носок'
    elif param_value.get() == 'zak_p':
        param, param_name = 6, 'закачка_пятка'
    elif param_value.get() == 'zak_sum':
        param_name = 'закачка_суммарная'

    result_tab = pd.DataFrame(
        index=['январь', 'февраль', 'март', 'апрель', 'май', 'июнь', 'июль', 'август', 'сентябрь', 'октябрь', 'ноябрь',
               'декабрь'])

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
        for data_day in db_skv.iterrows():
            if isinstance(data_day[1][1], datetime.datetime):
                if (data_day[1][1]).year == int(date_year.get()):
                    if (data_day[1][1]).month == month_data:
                        if param_value.get() == 'zak_sum':
                            month_param_skv1.append(data_day[1][5] + data_day[1][6])
                            if len(skv) > 5:
                                month_param_skv2.append(data_day[1][5 + int(columns_between_skv.get())] +
                                                        data_day[1][6 + int(columns_between_skv.get())])
                        else:
                            if data_day[1][param] == 0:  # для корректного расчета средних значений температуры,
                                data_day[1][param] = None  # если часть месяца 0
                            month_param_skv1.append(data_day[1][param])
                            if len(skv) > 5:
                                if data_day[1][param + int(columns_between_skv.get())] == 0:
                                    data_day[1][param + int(columns_between_skv.get())] = None
                                month_param_skv2.append(data_day[1][param + int(columns_between_skv.get())])
                    else:

                        if param_value.get() == 'dob_j' or param_value.get() == 'dob_n' or param_value.get() == 'zak_n' or \
                                param_value.get() == 'zak_p' or param_value.get() == 'zak_sum':
                            data_year_by_month_skv1.append(np.sum(list(filter(None, month_param_skv1))))  # удаление None
                            # из списка перед расчетом суммы и среднего
                            if len(skv) > 5:
                                data_year_by_month_skv2.append(np.sum(list(filter(None, month_param_skv2))))
                        else:
                            data_year_by_month_skv1.append(np.mean(list(filter(None, month_param_skv1))))
                            if len(skv) > 5:
                                data_year_by_month_skv2.append(np.mean(list(filter(None, month_param_skv2))))
                        month_data += 1
                        if param_value.get() == 'zak_sum':
                            month_param_skv1 = [data_day[1][5] + data_day[1][6]]
                            if len(skv) > 5:
                                month_param_skv2 = [data_day[1][5 + int(columns_between_skv.get())] +
                                                    data_day[1][6 + int(columns_between_skv.get())]]
                        else:
                            if data_day[1][param] == 0:
                                data_day[1][param] = None
                            month_param_skv1 = [data_day[1][param]]
                            if len(skv) > 5:
                                if data_day[1][param + int(columns_between_skv.get())] == 0:
                                    data_day[1][param + int(columns_between_skv.get())] = None
                                month_param_skv2 = [data_day[1][param + int(columns_between_skv.get())]]
            else:
                break
        if param_value.get() == 'dob_j' or param_value.get() == 'dob_n' or param_value.get() == 'zak_n' or \
                param_value.get() == 'zak_p' or param_value.get() == 'zak_sum':
            data_year_by_month_skv1.append(np.sum(list(filter(None, month_param_skv1))))
            if len(skv) > 5:
                data_year_by_month_skv2.append(np.sum(list(filter(None, month_param_skv2))))
        else:
            data_year_by_month_skv1.append(np.mean(list(filter(None, month_param_skv1))))
            if len(skv) > 5:
                data_year_by_month_skv2.append(np.mean(list(filter(None, month_param_skv2))))
        if month_data != 12:
            data_year_by_month_skv1.extend([None] * (12 - month_data))
            if len(skv) > 5:
                data_year_by_month_skv2.extend([None] * (12 - month_data))
        result_tab[skv[0:5]] = data_year_by_month_skv1
        if len(skv) > 5:
            result_tab[skv[-5:]] = data_year_by_month_skv2
    print(result_tab)
    result_tab.to_excel(name_file_excel.get()[:-4] + '_' + param_name + date_year.get() + '.xlsx')
    plt.figure()
    sns.heatmap(result_tab.transpose(), linewidths=.5, linecolor='black', cmap='jet', yticklabels=1)
    plt.title((param_name + '_' + file_name)[:-4])
    plt.tight_layout()
    plt.show()


find_excel = Label(text='Выберите Excel-файл параметров разработки')
find_excel.grid(row=0, column=1, padx=10, pady=5)

file_excel = Button(text='Файл Excel', command=new_file, width=10)
file_excel.grid(row=1, column=0, padx=10, pady=5)
name_file_excel = Entry(width=70)
name_file_excel.grid(row=1, column=1, padx=10)

date_year = Entry(width=5)
date_year.insert(0, datetime.date.today().year)
date_year.grid(row=2, column=0, pady=5)
select_param = Label(text='Введите год разработки')
select_param.grid(row=2, column=1)

param_value = StringVar(value='t_ust')
param_column = OptionMenu(root, param_value, 't_ust', 't_nas', 't_mas', 'p_mas', 'dob_j', 'dob_n', 'obv', 'obv_p',
                          'zak_n', 'zak_p', 'zak_sum')
param_column.grid(row=3, column=0, pady=5)
select_param = Label(text='Выберите параметр разработки')
select_param.grid(row=3, column=1)

columns_between_skv = Entry(width=5)
columns_between_skv.insert(0, '19')
columns_between_skv.grid(row=4, column=0, pady=5)
select_column_between_skv = Label(text='Количество колонок между парами скважин')
select_column_between_skv.grid(row=4, column=1)

calculate = Button(text='Рассчитать', command=calculate_param, width=10)
calculate.grid(row=5, column=0)

root.mainloop()
