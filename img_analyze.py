from PIL import Image
import img2pdf
import sys
from tqdm import tqdm
from glob import glob
import pandas as pd

# from tkinter import Tk
# from tkinter import Button
# from tkinter import Label
from tkinter import filedialog
# from tkinter import ttk
# from tkinter.ttk import Progressbar
from tkinter import messagebox
# from tkinter import Checkbutton
# from tkinter import BooleanVar
# from tkinter import Entry
# from tkinter import END

from customtkinter import CTk
from customtkinter import CTkButton
from customtkinter import CTkLabel
from customtkinter import CTkCheckBox
from customtkinter import BooleanVar
from customtkinter import CTkProgressBar
from customtkinter import END
# from customtkinter import CTkInputDialog
# from customtkinter import IntVar
# from customtkinter import StringVar


from threading import Thread

import os
from random import randrange

import logging

from functools import partial

Image.MAX_IMAGE_PIXELS = 933120000
set_version = '2.6'
output = os.getcwd() + '\\report.xlsx'
# default dpi for calculating
set_dpi = 300
extensions = ['jpg', 'jpeg', 'tif', 'tiff', 'png']

# logging config
logging.basicConfig(level=logging.INFO,
                    format="%(asctime)s %(levelname)s %(message)s",
                    handlers=[
                        logging.FileHandler(os.getcwd() + '\\log.log', 'a'),
                        logging.StreamHandler()
                    ])


def get_format(size):
    df = formats.loc[(formats.variable == 'abt') & (formats.unit == 'pixel')].copy()
    dff = formats.loc[(formats.variable == 'standart') & (formats.unit == 'pixel')].copy()
    try:
        # Определение точных форматов в заданных границах
        page_format_height = df.loc[(df.h_min <= size[1]) & (df.h_max >= size[1]), 'format'].values[0]
        page_format_width = df.loc[(df.w_min <= size[0]) & (df.w_max >= size[0]), 'format'].values[0]
        if page_format_height == page_format_width:
            return page_format_height
        # Определение кратных форматов
        else:
            page_format = df.loc[(df.h_min <= size[0]) & (df.h_max >= size[0]), 'format'].values[0]
            multiplicity = round(size[1] / dff.loc[dff.format == page_format, 'w'].values[0])
            return page_format + 'X' + str(multiplicity)
    except Exception as e:
        messagebox.showerror('Error', str(e))
        exit(-1)


def get_format3(size):
    if size[0] >= 1189 and size[1] >= 2523:
        return 'A0X3'
    if size[0] >= 1189 and size[1] >= 1682:
        return 'A0X2'
    if size[0] >= 841 and size[1] >= 2378:
        return 'A1X4'
    if size[0] >= 841 and size[1] >= 1783:
        return 'A1X3'
    if size[0] >= 841 and size[1] >= 1189:
        return 'A0X1'
    if size[0] >= 594 and size[1] >= 2102:
        return 'A2X5'
    if size[0] >= 594 and size[1] >= 1682:
        return 'A2X4'
    if size[0] >= 594 and size[1] >= 1261:
        return 'A2X3'
    if size[0] >= 594 and size[1] >= 841:
        return 'A1X1'
    if size[0] >= 420 and size[1] >= 2080:
        return 'A3X7'
    if size[0] >= 420 and size[1] >= 1783:
        return 'A3X6'
    if size[0] >= 420 and size[1] >= 1486:
        return 'A3X5'
    if size[0] >= 420 and size[1] >= 1189:
        return 'A3X4'
    if size[0] >= 420 and size[1] >= 891:
        return 'A3X3'
    if size[0] >= 420 and size[1] >= 594:
        return 'A2X1'
    if size[0] >= 297 and size[1] >= 1892:
        return 'A4X9'
    if size[0] >= 297 and size[1] >= 1682:
        return 'A4X8'
    if size[0] >= 297 and size[1] >= 1471:
        return 'A4X7'
    if size[0] >= 297 and size[1] >= 1261:
        return 'A4X6'
    if size[0] >= 297 and size[1] >= 1051:
        return 'A4X5'
    if size[0] >= 297 and size[1] >= 841:
        return 'A4X4'
    if size[0] >= 297 and size[1] >= 630:
        return 'A4X3'
    if size[0] >= 297 and size[1] >= 420:
        return 'A3X1'
    if size[0] >= 210 and size[1] >= 297:
        return 'A4X1'
    if size[0] <= 166 and size[1] <= 295:
        return 'a4_incorrected'
    else:
        return 'no_format'
    # if size[0] >= 186 and size[1] >= 276:
    #     return 'a4_incorrected'


def page_info(path=None):
    try:
        path = filedialog.askdirectory()
        if path == '':
            logging.info('Задан пустой путь или нажата кнопка отмены')
            messagebox.showwarning('Некорректный путь', 'Укажи ка корректный путь')
            return
        logging.info('Запущен анализ файлов')
        # files = [file for file in glob.glob(path + '\\**\\*.jpg', recursive=recursive.get())]
        files = [file for file in glob(path + '\\**\\*.*', recursive=recursive.get()) if file.split('.')[-1] in extensions]
        df = pd.DataFrame({'path': files})
        df['file'] = df.path.map(lambda x: x.split('\\')[-1])
        df['dir'] = df.path.map(lambda x: x.split('\\')[-2])
        # df[['size', 'color', 'dpi', 'format', 'comment']] = ''
        df[['size', 'size1', 'size2', 'color', 'dpi', 'format', 'comment', 'resize']] = ''
        counter = 0
        cnt = len(df)
        for i, row in tqdm(df.iterrows(), total=len(df)):
            try:
                img = Image.open(row.path)
                dpi = img.info.get('dpi')
                size = img.size
                if dpi != (set_dpi, set_dpi):
                    df.loc[i, 'comment'] = f'incorrect dpi: {dpi}'
                    # continue
                # reformat
                if size[0] > size[1]:
                    size = (size[1], size[0])
                    df.loc[i, 'resize'] = 1
                # pixels to mm
                px1 = round(size[0] * 25.4 / set_dpi)
                px2 = round(size[1] * 25.4 / set_dpi)
                pxs = (px1, px2)
                if dmitry_spec.get() is False:
                    page_format = get_format3(pxs)
                else:
                    page_format = get_format(size)
                # df.loc[i, ['size', 'color', 'dpi', 'format']] = str(img.size), img.mode, str(dpi), get_format2(img.size)
                df.loc[i, ['size', 'size1', 'size2', 'color', 'dpi', 'format']] = str(size), pxs[0], pxs[
                    1], img.mode, str(dpi), page_format
            except Exception as e:
                logging.error(f'\nНе удается прочитать файл {row.path}')
                df.loc[i, 'comment'] = 'Ошибка чтения файла: ' + str(e)
            counter += 1
            # bar['value'] = counter / cnt * 100
            bar.set(counter / cnt * 100)
            # bar_value.set(counter / cnt * 100)
            # bar.update_idletasks()
        # # delete string in 40 state
        # if questionaries.get() != '':
        #     df['Удаленный'] = ''
        #     drop_files = pd.read_csv(questionaries.get(), sep='\t', encoding='cp1251')
        #     drop_files = drop_files.loc[drop_files.QuestionaryStateID == 40, 'ID'].to_list()
        #     drop_files = [str(x) + '.jpg' for x in drop_files]
        #     df.loc[df.file.isin(drop_files), 'Удаленный'] = 'Да'
        #     # if len(df.file.isin(drop_files)) > 0:
        #     #     messagebox.showwarning('Info', 'Обнаружены строки помеченных файлов, которые удалены. Будет удалено ' + str(len(df.file.isin(drop_files))) + ' строк')
        #     #     df.drop(df.loc[df.file.isin(drop_files)].index, inplace=True)
        #     del drop_files
        output_temp = output
        try:
            df.to_excel(output_temp)
        except PermissionError as e:
            logging.error('Не удалось сохранить отчет: '+ str(e))
            messagebox.showerror('Не удалось сохранить отчет', 'Нажмите ОК чтобы сохранить с другим именем')
            logging.warning('Возобновлена попытка сохранить с другим именем...')
            output_temp = output_temp[:-5] + str(randrange(1, 999999)) + '.xlsx'
            df.to_excel(output_temp)
        logging.info(f'Отчет сохранен в {output_temp}')
        messagebox.showinfo('Success', 'Отчет сохранен в ' + output_temp)
        # pack_to_pdf(df)
    except Exception as e:
        logging.critical(str(e))
        messagebox.showerror('Error', 'Не удалось завершить анализ форматов, информация передана в лог')


def pack_to_pdf(df):
    for dir in tqdm(df.dir.unique()):
        df2 = df.loc[df.dir == dir].sort_values(by='path').copy()
        files = df2['path'].to_list()
        with open('./output.pdf', 'wb') as write_pdf:
            write_pdf.write(img2pdf.convert(files))


def start():
    task = Thread(target=page_info)
    task.run()


def debug():
    size = (3678, 13000)
    if size[0] > size[1]:
        size = (size[1], size[0])
    print(size)
    print(get_format(size))


def get_path(line, f_types):
    if f_types is None:
        path = filedialog.askdirectory()
    else:
        path = filedialog.askopenfilename(filetypes=[f_types])
    line.delete(0, END)
    line.insert(0, path)
    return path


def rebuild_format(df, dpi):
    logging.info(f'Эталон по размерам в пикселях приведен к dpi {dpi}')
    # dff = df.loc[df['variable'] == 'abt']
    for row in df.iterrows():
        for size in ('h', 'w', 'h_min', 'h_max', 'w_min', 'w_max'):
            df.loc[row[0], size] = round(row[1][size] / set_dpi * dpi) if row[1][size] not in (0, 99999999) else row[1][size]
    df.to_excel(os.getcwd() + '\\formats_temp.xlsx')
    logging.info('Приведенный эталон записан в корневую папку с приложением (formats_temp.xlsx)')
    return df


def read_format_file():
    try:
        logging.info('Найден файл с параметрами для вычисления, чтение файла...')
        formats = pd.read_excel(os.getcwd() + '/format.xlsx', sheet_name='formats')
        formats.fillna(0, inplace=True)
        settings = pd.read_excel(os.getcwd() + '/format.xlsx', sheet_name='settings')
        settings = {row[1]['parameter']: row[1]['value'] for row in settings.iterrows()}
        if settings['dpi'] != set_dpi:
            formats = rebuild_format(formats, settings['dpi']).copy()
    except Exception as e:
        logging.critical(str(e))
        messagebox.showerror('Критическая ошибка', 'Файл "format" старой версии, замените файл!\n'
                                                   'При нажатии ОК будет осуществлен выход из программы\n')
        exit(-1)
    logging.info('Параметры для вычисления успешно переданы')
    return formats, settings


logging.info(f'Запуск программы. Версия: {set_version}')
window = CTk()
window.title(f'Анализ форматов файлов. Версия: {set_version}')
window.geometry('300x215')

logging.info(f'Доступные форматы изображений (в будущем): Fully (BLP, BMP, DDS, DIB, EPS), Load (GIF, ICNS, ICO, IM, '
             f'JPEG, JPEG 2000, MSP, PCX, PNG, APNG, PPM, SGI, SPIDER, TGA, TIFF, WebP, XBM), Read (CUR, DCX, FITS, '
             f'FLI, FLC, FPX, FTEX, GBR, GD), Open (IMT, IPTC/NAA, MCIDAS, MIC, MPO, PCD, PIXAR, PSD, QOI, SUN, WAL, '
             f'WMF, EMF, XPM)')
logging.info(f'Поддерживаемые типы изображений: {str(extensions)}')

CTkLabel(window, text='Необходимо выбрать коневую папку с пачками').grid(column=0, row=0)
btn_start = CTkButton(window, text='Выбрать папку с пачками', command=start)
btn_start.grid(column=0, row=1)
# # debugging button
# btn_deb = Button(window, text='debug', command=debug)
# btn_deb.grid(column=1, row=1)
# Label(window, text='dpi для анализа прописано по умолчанию 300!').grid(column=0, row=2)
# CTkLabel(window, text='dpi для анализа прописано по умолчанию 300!').grid(column=0, row=2, sticky="nsew")
CTkCheckBox(window, text='Формировать pdf по пачкам', state='disabled').grid(column=0, row=3, sticky="nsew", padx=5)

# add checkbox for search recursive_folder
recursive = BooleanVar(value=True)
CTkCheckBox(window, text='просмотр файлов в подкаталогах', variable=recursive).grid(column=0, row=4, sticky="nsew", padx=5, pady=5)
# add checkbox for alternative format
if os.path.exists(os.getcwd() + '/format.xlsx'):
    formats, settings = read_format_file()
    set_dpi = int(settings['dpi']) if 'dpi' in settings else 300
    dmitry_spec = BooleanVar(value=True)
    dmitry_state = 'normal'
else:
    logging.warning('Файл с параметрами не найден! Используется стандартный функционал с альтернативным методом '
                    'вычисления')
    dmitry_spec = BooleanVar(value=False)
    dmitry_state = 'disabled'
CTkCheckBox(window, text='Расчет по алгоритму Дмитрия', variable=dmitry_spec, state=dmitry_state).grid(column=0, row=5, sticky="nsew", padx=5)

# CTkLabel(window, text='dpi для анализа прописано по умолчанию 300!').grid(column=0, row=2, sticky="nsew")
CTkLabel(window, text=f'dpi для анализа: ' + str(set_dpi), compound='left').grid(column=0, row=2)

# # add support for delete 40 state
# Label(window, text='Путь к файлу Questionaries').grid(column=0, row=6)
# questionaries = Entry(window, width=25)
# btn_q = Button(window, text='...', command=partial(get_path, questionaries, ('CSV files', '*.csv')))
# btn_q.grid(column=1, row=6)
# questionaries.grid(column=2, row=6)


CTkLabel(window, text='Статус:').grid(column=0, row=7)
# style = ttk.Style()
# style.theme_use('default')
# style.configure("black.Horizontal.TProgressbar", background='black')
# bar_value = IntVar(value=0)
bar = CTkProgressBar(window)
# bar = Progressbar(window, length=200, style='black.Horizontal.TProgressbar')
# bar['value'] = 0
bar.set(0, 100)
bar.grid(column=0, row=8)

logging.info(f'Установлен DPI {set_dpi}')
logging.info('Программа успешно запущена')

window.mainloop()
