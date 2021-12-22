# Author Loik Andrey 7034@balancedv.ru

import pandas as pd
import os


FOLDER = 'Исходные данные'
BALANCE_NAME = 'Остатки'
MIN_STOCK_NAME = 'МО'
NEW_FILE_NAME= 'Анализ потребностей.xlsx'
NEW_FILE_NAME1= 'Потребность для заказа.xlsx'


def Run():

    balanceFilelist = search_file(BALANCE_NAME)  # запускаем функцию по поиску файлов и получаем список файлов
    minStockFilelist = search_file(MIN_STOCK_NAME)  # запускаем функцию по поиску файлов и получаем список файлов
    df_balance = create_df (balanceFilelist, BALANCE_NAME) # создаём DF по остаткам
    df_minStock = create_df (minStockFilelist, MIN_STOCK_NAME) # создаём DF по мин остаткам
    df_balance= df_sum(df_balance)
    df_general = concat_df (df_balance, df_minStock)
    quant_order = payment(df_general)
    df_analysis = concat_df (df_general, quant_order)
    df_write_xlsx(df_analysis, NEW_FILE_NAME)
    df_write_xlsx(quant_order.dropna(), NEW_FILE_NAME1)
    return

def search_file(name):
    """
    :param name: Поиск всех файлов в папке FOLDER, в наименовании которых, содержится name
    :return: filelist список с наименованиями фалов
    """
    filelist = []
    for item in os.listdir(FOLDER):
        if name in item and item.endswith('.xlsx'): # если файл содержит name и с расширенитем .xlsx, то выполняем
            filelist.append(FOLDER + "/" + item) # добаляем в список папку и имя файла для последующего обращения из списка
        else:
            pass
    return filelist

def create_df (file_list, add_name):
    """
    :param file_list: Загружаем в DataFrame файлы из file_list
    :param add_name: Добавляем add_name в наименование колонок DataFrame
    :return: df_result Дата фрэйм с данными из файлов
    """
    df_result = pd.DataFrame()

    for filename in file_list: # проходим по каждому элементу списка файлов
        print (filename) # для тестов выводим в консоль наименование файла с которым проходит работа
        df = read_my_excel(filename)
        df_search_header = df.iloc[:15, :10] # для ускорения работы выбираем из DataFrame первую колонку и 15 строк
        # создаём маску и отмечаем True строку где есть слово "Номенклатура", остальные False
        mask = (df_search_header == 'Номенклатура')
        # Преобразуем Dataframe согласно маски. После обработки все значения будут NaN кроме нужного нам.
        # В этой же строке кода удаляем все строки со значением NaN и далее получаем индекс оставшейся строки
        f = df_search_header[mask].dropna(axis=0, how='all').index.values # Удаление пустых колонок, если axis=0, то строк
        col = df[mask].dropna(axis=1, how='all').columns.values
        df = df.iloc[int(f):, :] # Убираем все строки с верха DF до заголовков
        df = df.dropna(axis=1, how='all')  # Убираем пустые колонки

        if add_name == MIN_STOCK_NAME:
            df.iloc[int(f), 1] = 'Артикул'
            df.iloc[int(f), col-1] = 'Номенклатура'
            df.columns = df.iloc[int(f)] # Значения из найденной строки переносим в заголовки DataFrame для простоты дальнейшего обращения
            df = df[['Артикул', 'Номенклатура', 'МО внешний']]
            df = df.iloc[3:len(df)-1, :]  # Убираем три строки с верха DF и одну снизу
            df['МО внешний'] = df['МО внешний'].replace(',', '.', regex=True).astype('float64')
        else:
            df.iloc[0, 0] = 'Артикул'
            df.iloc[0, 1] = 'Номенклатура'
            df.columns = df.iloc[0] # Значения из найденной строки переносим в заголовки DataFrame для простоты дальнейшего обращения
            df = df.iloc[2:, :] # Убираем две строки с верха DF

        df.columns.name = None
        df['Номенклатура'] = df['Номенклатура'].str.strip() # Удалить пробелы с обоих концов строки в ячейке
        df.set_index(['Артикул', 'Номенклатура'], inplace=True) # переносим колонки в индекс, для упрощения дальнейшей работы
        df_result = concat_df(df_result, df)
    return df_result

def read_my_excel (file_name):
    """
    Пытаемся прочитать файл xlxs, если не получается, то исправляем ошибку и опять читаем файл
    :param file_name: Имя файла для чтения
    :return: DataFrame
    """
    print ('Попытка загрузки файла:'+file_name)
    try:
        df = pd.read_excel(file_name, sheet_name='TDSheet', header=None, skipfooter=0, engine='openpyxl')
        return (df)
    except KeyError as Error:
        print (Error)
        df = None
        if str(Error) == "\"There is no item named 'xl/sharedStrings.xml' in the archive\"":
            bug_fix (file_name)
            print('Исправлена ошибка: ', Error, f'в файле: \"{file_name}\"\n')
            df = pd.read_excel(file_name, sheet_name='TDSheet', header=None, skipfooter=0, engine='openpyxl')
            return df
        else:
            print('Ошибка: >>' + str(Error) + '<<')

def bug_fix (file_name):
    """
    Переименовываем не корректное имя файла в архиве excel
    :param file_name: Имя excel файла
    """
    import shutil
    from zipfile import ZipFile
    from rarfile import RarFile

    # Создаем временную папку
    tmp_folder = '/temp/'
    os.makedirs(tmp_folder, exist_ok=True)

    # Распаковываем excel как zip в нашу временную папку и удаляем excel
    try:
        with ZipFile(file_name) as excel_container:
            excel_container.extractall(tmp_folder)
    except:
        with RarFile(file_name) as excel_container:
            excel_container.extractall(tmp_folder)
    os.remove(file_name)

    # Переименовываем файл с неверным названием
    wrong_file_path = os.path.join(tmp_folder, 'xl', 'SharedStrings.xml')
    correct_file_path = os.path.join(tmp_folder, 'xl', 'sharedStrings.xml')
    os.rename(wrong_file_path, correct_file_path)

    # Запаковываем excel обратно в zip и переименовываем в исходный файл
    try:
        shutil.make_archive(f'{FOLDER}/correct_file', 'zip', tmp_folder)
    except:
        shutil.make_archive(f'{FOLDER}/correct_file', 'rar', tmp_folder)
    os.rename(f'{FOLDER}/correct_file.zip', file_name)

def concat_df (df1, df2):
    df = pd.concat([df1, df2], axis=1, ignore_index=False)
    return df

def df_sum(df):
    dfsum = pd.DataFrame()
    dfsum['Остатки по компании'] = df.sum(axis=1)
    return dfsum

def payment(df):
    df = df.fillna(0)
    df['Компания MaCar'] = df['МО внешний']
    df = df[['Остатки по компании', 'Компания MaCar']]
    mask1 = df['Компания MaCar'] >= 1
    mask = df['Компания MaCar'][mask1] > df['Остатки по компании'][mask1]
    df['Потребность'] = df['Компания MaCar'][mask1][mask] - df['Остатки по компании'][mask1][mask]
    df = df['Потребность']
    return df

def df_write_xlsx(df, name_file):
    # Сохраняем в переменные значения конечных строк и столбцов
    try:
        row_end, col_end = len(df), len(df.columns)
    except:
        row_end, col_end = len(df), 1
    row_end_str, col_end_str = str(row_end), str(col_end)

    # Сбрасываем встроенный формат заголовков pandas
    pd.io.formats.excel.ExcelFormatter.header_style = None

    # Создаём эксель и сохраняем данные
    sheet_name = 'Данные'  # Наименование вкладки для сводной таблицы
    writer = pd.ExcelWriter(name_file, engine='xlsxwriter')
    workbook = writer.book
    df.to_excel(writer, sheet_name=sheet_name)
    wks1 = writer.sheets[sheet_name]  # Сохраняем в переменную вкладку для форматирования

    # Получаем словари форматов для эксель
    header_format, con_format, border_storage_format_left, border_storage_format_right, \
    name_format, MO_format, data_format = format_custom(workbook)

    # Форматируем таблицу
    wks1.set_default_row(12)
    wks1.set_row(0, 20, header_format)
    wks1.set_column('A:A', 12, name_format)
    wks1.set_column('B:B', 32, name_format)
    wks1.set_column('C:E', 10, data_format)

    # Делаем жирным рамку между складами и форматируем колонку с МО по всем складам
    """    wks1.set_column(2, 2, None, border_storage_format_left)
    wks1.set_column(5, 5, None, border_storage_format_right)
    wks1.set_column(6, 6, None, border_storage_format_left)
    wks1.set_column(7, 7, None, border_storage_format_right)
    wks1.set_column(7, 7, None, MO_format)"""

    # Добавляем фильтр в первую колонку
    wks1.autofilter(0, 0, row_end+1, col_end+1)

    # Сохраняем файл
    writer.save()
    return

def format_custom(workbook):
    header_format = workbook.add_format({
        'font_name': 'Arial',
        'font_size': '7',
        'align': 'center',
        'valign': 'top',
        'text_wrap': True,
        'bold': True,
        'bg_color': '#F4ECC5',
        'border': True,
        'border_color': '#CCC085'
    })

    border_storage_format_left = workbook.add_format({
        'num_format': '# ### ##0.00',
        'font_name': 'Arial',
        'font_size': '8',
        'left': 2,
        'left_color': '#000000',
        'bottom': True,
        'bottom_color': '#CCC085',
        'top': True,
        'top_color': '#CCC085',
        'right': True,
        'right_color': '#CCC085',
    })
    border_storage_format_right = workbook.add_format({
        'num_format': '# ### ##0.00',
        'font_name': 'Arial',
        'font_size': '8',
        'right': 2,
        'right_color': '#000000',
        'bottom': True,
        'bottom_color': '#CCC085',
        'top': True,
        'top_color': '#CCC085',
        'left': True,
        'left_color': '#CCC085',
    })

    name_format = workbook.add_format({
        'font_name': 'Arial',
        'font_size': '8',
        'align': 'left',
        'valign': 'top',
        'text_wrap': True,
        'bold': False,
        'border': True,
        'border_color': '#CCC085'
    })

    MO_format = workbook.add_format({
        'num_format': '# ### ##0.00;;',
        'bold': True,
        'font_name': 'Arial',
        'font_size': '8',
        'font_color': '#FF0000',
        'right': 2,
        'right_color': '#000000',
        'bottom': True,
        'bottom_color': '#CCC085',
        'top': True,
        'top_color': '#CCC085',
        'left': True,
        'left_color': '#CCC085',
    })
    data_format = workbook.add_format({
        'num_format': '# ### ##0.00',
        'font_name': 'Arial',
        'font_size': '8',
        'text_wrap': True,
        'border': True,
        'border_color': '#CCC085'
    })
    con_format = workbook.add_format({
        'bg_color': '#FED69C',
    })

    return header_format, con_format, border_storage_format_left, border_storage_format_right, \
           name_format, MO_format, data_format

if __name__ == '__main__':
    Run()
