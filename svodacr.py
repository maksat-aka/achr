import numpy as np
import tkinter as tk
from tkinter import filedialog
import pandas as pd

root = tk.Tk()

canvas1 = tk.Canvas(root, width=300, height=300, bg='lightsteelblue')
canvas1.pack()


# АктМЭС АлмМЭС АкмМЭС ЦентрМЭС ВостМЭС СарбМЭС ФСМЭС Уральск ЮжМЭС ЗапМЭС
# akty.xls', 'alma.xls', 'asta.xls', 'cent.xls', 'east.xls', 'kust.xls', 'nort.xls', 'oral.xls', 'sout.xls', 'west.xls'

def getExcel():
    global df_fs
    global df_data
    global filenames
    global import_folder_path
    # import_file_path = filedialog.askopenfilename()
    import_folder_path = filedialog.askdirectory()

    filenames = ['akty', 'alma', 'asta', 'cent', 'east', 'kust', 'nort', 'oral', 'sout', 'west']
    canvas1.create_text(150, 80, text='Выбран каталог: ' + import_folder_path)

    df_fs = {}
    df_data = {}
    i = 12
    for name in filenames:
        import_file_path = import_folder_path + '/' + name + '.xls'

        try:
            df_fs[name] = pd.read_excel(import_file_path, nrows=1)
            canvas1.create_text(150, 80 + i, text='загружен файл: ' + import_file_path)

            df_fs[name].columns = ['MESName', 'Fakt', 'PotrFakt', 'PotrPlan', 'PlanACR', 'PotrDate', '', '', '', '', '']

            df_data[name] = pd.read_excel(import_file_path, skiprows=2)
            df_data[name].columns = ['N_ACR1', 'N_ACR2', 'U_F_ACR', 'U_T_ACR', 'V_ACR', 'U_F_ACR2', 'U_T_ACR2',
                                     'V_ACR2',
                                     'U_F_CAPV', 'U_T_CAPV', 'V_CAPV']
        except:
            canvas1.create_text(150, 80 + i, text='Ошибка загрузки файла: ' + import_file_path)
            print('Ошибка загрузки файла: ' + import_file_path)
        i += 12


def genSvod2():
    df_nort = {}
    df_nort2 = {}
    fn_nort = ['east', 'cent', 'kust', 'nort', 'asta', 'akty']
    df_west = {}
    df_west2 = {}
    fn_west = ['west', 'oral']
    df_sout = {}
    df_sout2 = {}
    fn_sout = ['alma', 'sout']

    for name in fn_nort:
        count = 0
        for i, row in df_data[name].iterrows():
            if row['N_ACR2'] > 0:
                break
            else:
                count += 1
        count2 = df_data[name]['U_F_ACR'].count() - count
        df_nort[name] = df_data[name].head(count).copy()
        df_nort2[name] = df_data[name].tail(count2).copy()

    for name in fn_west:
        count = 0
        for i, row in df_data[name].iterrows():
            if row['N_ACR2'] > 0:
                break
            else:
                count += 1
        count2 = df_data[name]['U_F_ACR'].count() - count
        df_west[name] = df_data[name].head(count).copy()
        df_west2[name] = df_data[name].tail(count2).copy()

    for name in fn_sout:
        count = 0
        for i, row in df_data[name].iterrows():
            if row['N_ACR2'] > 0:
                break
            else:
                count += 1
        count2 = df_data[name]['U_F_ACR'].count() - count
        df_sout[name] = df_data[name].head(count).copy()
        df_sout2[name] = df_data[name].tail(count2).copy()

    # находим суммы объемов соответсвенно уставкам

    frames = [df_nort['east'], df_nort['cent'], df_nort['kust'], df_nort['nort'], df_nort['asta'], df_nort['akty']]
    df_nort_all = pd.concat(frames)
    df_nort['nort_all'] = pd.concat(frames)
    fn_nort.append('nort_all')

    df_nort_U_F_ACR_list = df_nort_all.groupby('U_F_ACR').nunique().index.tolist()
    df_nort_U_T_ACR_list = df_nort_all.groupby('U_T_ACR').nunique().index.tolist()
    # for name in fn_nort:
    #     df_nort_U_F_ACR_list = df_nort[name].groupby('U_F_ACR').nunique().index.tolist()
    #     df_nort_U_T_ACR_list = df_nort[name].groupby('U_T_ACR').nunique().index.tolist()
    sv_nort = pd.DataFrame(columns=['count', 'ufacr', 'T', 'east', 'cent', 'kust', 'nort', 'asta', 'akty', 'nort_all'])
    tmp = {}
    count = 0
    for ufacr in df_nort_U_F_ACR_list:
        for utacr in df_nort_U_T_ACR_list:
            sv_nort.loc[count, 'ufacr'] = ufacr
            sv_nort.loc[count, 'T'] = utacr
            sv_nort.loc[count, 'count'] = count
            for name in fn_nort:
                tmp[name] = df_nort[name].loc[(df_nort[name].U_F_ACR == ufacr) & (df_nort[name].U_T_ACR == utacr)]
                sv_nort[count, name] = tmp[name].V_ACR.sum()
            count += 1

    tmp_series1 = pd.Series()
    tmp_series2 = pd.Series()
    tmp_series3 = pd.Series()
    tmp_series4 = pd.Series()
    count = 0
    for name in fn_nort:
        for ufacr in df_nort_U_F_ACR_list:
            for utacr in df_nort_U_T_ACR_list:
                tmp[name] = df_nort[name].loc[(df_nort[name].U_F_ACR == ufacr) & (df_nort[name].U_T_ACR == utacr)]
                # if(tmp[name].empty == False):
                #     str1 = name + ',' + str(ufacr) + ',' + str(utacr) + ',' +  str(tmp[name].V_ACR.sum())
                #     print (str1)
                #     tmp_series1.loc[count] = name
                #     tmp_series2.loc[count] = ufacr
                #     tmp_series3.loc[count] = utacr
                #     tmp_series4.loc[count] = tmp[name].V_ACR.sum()
                #     count += 1
                #     if (name == 'nort_all'):
                #         qwe = tmp[name].iloc[:, 8:]
                #         print(qwe.groupby('U_F_CAPV').sum())
                #         print('V_CAPV = ', tmp[name].V_CAPV.sum())
                ##################################################
                str1 = name + ',' + str(ufacr) + ',' + str(utacr) + ',' + str(tmp[name].V_ACR.sum())
                print(str1)
                tmp_series1.loc[count] = name
                tmp_series2.loc[count] = ufacr
                tmp_series3.loc[count] = utacr
                tmp_series4.loc[count] = tmp[name].V_ACR.sum()
                count += 1
                if (name == 'nort_all'):
                    qwe = tmp[name].iloc[:, 8:]
                    print(qwe.groupby('U_F_CAPV').sum())
                    print('V_CAPV = ', tmp[name].V_CAPV.sum())

    ##################################################

    print(sv_nort)
    sv_nort2 = pd.DataFrame({'name': tmp_series1,
                             'ufacr': tmp_series2,
                             'utacr': tmp_series3,
                             'vacr': tmp_series4})
    print(sv_nort2)
    sv_nort2.to_excel(import_folder_path+'/sv_nort2.xls')


# >>> ds = pd.Series([1,2,3,4,5])
# >>> ds.append(pd.Series([6]))
# 0 1 1 2 2 3 3 4 4 5 0 6 dtype: int64
# df = pd.DataFrame({'population': population,
#                    'area': area})


def printExcel():
    df = pd.DataFrame(index=['SumACR', 'SPotr', 'SumACR_P', 'ACRkPort_P', 'SACR1', 'SACR',
                             'SACR_P', 'SACR2N', 'SACR2N_P', 'SACR2S', 'SACR2S_P', 'SCAPV', 'SCAPV_P'],
                      columns=['PotrDate',
                               'east', 'cent', 'kust', 'nort', 'asta', 'akty', 'summ_nort',
                               'west', 'oral', 'summ_west',
                               'alma', 'sout', 'summ_sout',
                               'summ_all'
                               ])

    PotrDate = df_fs['east'].PotrDate.values[0]
    for name in filenames:
        if (PotrDate != df_fs[name].PotrDate.values[0]):
            PotrDate = 0
            # print('Даты не совпадают')
            print(name, df_fs[name].PotrDate.values[0])
    df['PotrDate'] = pd.Series(PotrDate,
                               ['SumACR', 'SPotr', 'SumACR_P', 'ACRkPort_P', 'SACR1', 'SACR', 'SACR_P', 'SACR2N',
                                'SACR2N_P', 'SACR2S', 'SACR2S_P', 'SCAPV', 'SCAPV_P']).copy()

    for name in filenames:
        tmp = df_data[name].loc[df_data[name].N_ACR2 == 0]
        tmp2 = df_data[name].loc[df_data[name].N_ACR2 > 0]
        tmp3 = df_data[name].loc[df_data[name].U_F_ACR == 49.2]

        df[name].SPotr = df_fs[name].PotrFakt.values[0]  # Потребление
        df[name].SCAPV = df_data[name].V_CAPV.sum()  # Всего ЧАПВ
        df[name].SACR1 = tmp.V_ACR.sum()  # Всего АЧР1
        df[name].SACR2S = tmp.V_ACR2.sum()  # Всего АЧР2 совм
        df[name].SACR2N = tmp2.V_ACR.sum()  # Всего АЧР2 несовм
        df[name].SACR = tmp3.V_ACR.sum()  # САЧР
        df[name].SumACR = df[name].SACR1 + df[name].SACR2N
        df[name].SumACR_P = df[name].SumACR / df[name].SPotr * 100
        df[name].SACR_P = df[name].SACR / df[name].SPotr * 100
        df[name].SACR2N_P = df[name].SACR2N / df[name].SPotr * 100
        df[name].SACR2S_P = df[name].SACR2S / df[name].SACR1 * 100
        df[name].SCAPV_P = df[name].SCAPV / df[name].SumACR * 100

    df['summ_nort'] = df.east + df.cent + df.kust + df.nort + df.asta + df.akty
    df['summ_west'] = df.west + df.oral
    df['summ_sout'] = df.alma + df.sout
    df['summ_all'] = df.summ_nort + df.summ_west + df.summ_sout

    filenames2 = ['summ_nort', 'summ_west', 'summ_sout', 'summ_all']
    for name in filenames2:
        df[name].SumACR_P = df[name].SumACR / df[name].SPotr * 100
        df[name].SACR_P = df[name].SACR / df[name].SPotr * 100
        df[name].SACR2N_P = df[name].SACR2N / df[name].SPotr * 100
        df[name].SACR2S_P = df[name].SACR2S / df[name].SACR1 * 100
        df[name].SCAPV_P = df[name].SCAPV / df[name].SumACR * 100

    df.columns = ['Час замера', 'Восточный регион', 'Центральный регион', 'Костанайский регион',
                  'Северный регион', 'Акмолинский регион', 'Актюб. регион', 'Северная часть',
                  'Западный регион', 'Ур.э/у', 'Западная часть',
                  'Алматинский регион', 'Южный регион', 'Южная часть', 'ЕЭС Казахстана']
    df.rename(index={
        'SumACR': 'Всего АЧР',
        'SPotr': 'Потребление',
        'SumACR_P': 'АЧР к потр., %',
        'ACRkPort_P': 'АЧР к потр., % (зад)',
        'SACR1': 'Всего АЧР-I ',
        'SACR': 'в т.ч. САЧР',
        'SACR_P': 'САЧР к потр., %',
        'SACR2N': 'Всего АЧР-II н/с',
        'SACR2N_P': 'АЧР-II н/с к  потр.,%',
        'SACR2S': 'Всего АЧР-II совм.',
        'SACR2S_P': 'АЧР-II совм. от АЧР-I,%',
        'SCAPV': 'Всего ЧАПВ',
        'SCAPV_P': 'ЧАПВ от АЧР, %'},
        inplace=True
    )

    writer = pd.ExcelWriter(import_folder_path + '/Сводная АЧР ' + str(pd.to_datetime(PotrDate).date()) + ' ' +
                            str(pd.to_datetime(PotrDate).hour) + '.xlsx',
                            engine='xlsxwriter',
                            datetime_format='hh:mm',
                            date_format='mmmm dd yyyy', )
    df.to_excel(writer, sheet_name='Sheet1')
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']

    format1 = workbook.add_format({'num_format': '#,##0.00'})
    format2 = workbook.add_format({'num_format': '0%'})

    worksheet.set_column('C:P', 12, format1)

    writer.save()
    print(df)


browseButton_Excel = tk.Button(text='Выбрать папку с файлами', command=getExcel, bg='green', fg='white',
                               font=('helvetica', 12, 'bold'))
browseButton_ExcelPr = tk.Button(text='Сводная 1', command=printExcel, bg='blue', fg='white',
                                 font=('helvetica', 12, 'bold'))
browseButton_Svod2 = tk.Button(text='Сводная 2', command=genSvod2, bg='blue', fg='white',
                               font=('helvetica', 12, 'bold'))
canvas1.create_window(150, 50, window=browseButton_Excel)
canvas1.create_window(70, 250, window=browseButton_ExcelPr)
canvas1.create_window(200, 250, window=browseButton_Svod2)

root.mainloop()
