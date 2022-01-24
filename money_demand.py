#!/usr/bin/env python3


### >>> Подключение модулей
# <
import os
import shutil
import xlrd, xlwt
from re import compile, sub, findall
from glob import glob1
from decimal import Decimal
# />


### >>> Подключение библиотек справочников
# <
from lib.liblore import opsDict, allNodesDict, monthDict2
# />


### >>> Каталоги скрипта
# <
dir = {
    'in'  : os.path.join(os.getcwd(), '_in'),
    'out' : os.path.join(os.getcwd(), '_out'),
    'arch': os.path.join(os.getcwd(), 'arch'),
    'lib' : os.path.join(os.getcwd(), 'lib')
}
# />


### >>> Восстановление структуры каталогов, если она нарушена
# <
for name in dir.keys():
    path = dir.get(name)
    if not os.path.exists(path):
        try:
            os.mkdir(path)
        except:
            pass
# />


### >>> Входящие файлы
# <
IN_FILES_NAME_TEMPLATE       = r'*.xls'         # шаблон имён
IN_FILES_ENCODING            = 'cp1251'         # кодировка входящих файлов
IN_FILE_NAME_PERIOD_TEMPLATE = r'^\d{1}'        # шаблон зашифрованного в имени периода выплаты
IN_FILE_NAME_DATE_TEMPLATE   = r'\d{2}_\d{4}'   # шаблон зашифрованного в имени месяца и года
IN_FILE_NAME_NODE_TEMPLATE   = r'[a-z]{3}'      # шаблон зашифрованного в имени узла
# />


### >>> Список имён столбцов для исключения во входящем файле
# <
IN_FILE_COLUMN_NAME_2_EXCLUDE = ['city']
# />


### >>> Исходящие файлы
# <
OUT_FILE_NAME_TEMPLATE = 'Пенсия. Денежная потребность на {0} период {1} месяца {2} г..xls'     # шаблон имени
#               {0} -- период выплаты
#               {1} -- название месяца в родительном падеже
#               {2} -- год
# />


# OUT_FILE_NAME = 'money_demand_of_primary_kassa.xls'


#dateTmpl   = r'[a-z]{1}'


required_nodes = ('ber', 'brs')

######################


### >>> Список узлов, обслуживаемых Берёзовской Главной Кассой
# <
MAIN_NODE = '3'
SERVED_NODES = list(MAIN_NODE)
SERVED_NODES.extend(allNodesDict.get(MAIN_NODE).get('upsList'))
SERVED_NODES = tuple(SERVED_NODES)
#                   +++ Узлы из справочника +++
#                    -------------------------
#                    3  -- Берёзовский РУПС
#                    6  -- Дрогичинский УПС
#                    11 -- Кобринский УПС
#                    16 -- Пружанский УПС
# />


# Получение списка входящих файлов, удовлетворяющих шаблону IN_FILES_NAME_TEMPLATE
inFileList = glob1(dir.get('in'), IN_FILES_NAME_TEMPLATE)


### >>> Предварительный анализ входящих файлов и заполнение словаря filesInfoDict
# <
if inFileList:

    filesInfoDict = {}
    for inFile in inFileList:
        baseNameInFile = os.path.splitext(inFile)[0]

        # открываем файл
        rb = xlrd.open_workbook(
                                    filename=os.path.join(dir.get('in'), inFile),
                                    encoding_override=IN_FILES_ENCODING,
                                    on_demand=True,
                                    formatting_info=True
                                )

        # выбираем активный лист
        sheet = rb.sheet_by_index(0)

        # получаем значение первой строки (заголовок)
        rowHead = sheet.row_values(0)

        # формируем список индексов столбцов для исключения из списка имён столбцов
        idx2excludeList = [idx for idx, column in enumerate(rowHead) if column in IN_FILE_COLUMN_NAME_2_EXCLUDE]

        rows = []
        for rownum in range(1, sheet.nrows):
            index = str(int(sheet.row_values(rownum)[0]))
            if index in opsDict.keys():
                node = opsDict.get(index).get('node')
                if node in SERVED_NODES:
                    rows.append([value for idx, value in enumerate(sheet.row_values(rownum)) if idx not in idx2excludeList])

        # освобождаем ресурсы
        rb.release_resources()

        filesInfoDict[inFile] = {
                                  'period' : findall(IN_FILE_NAME_PERIOD_TEMPLATE, baseNameInFile)[0],
                                  'node'   : findall(IN_FILE_NAME_NODE_TEMPLATE,   baseNameInFile)[0].lower(),
                                  'fdate'  : findall(IN_FILE_NAME_DATE_TEMPLATE,   baseNameInFile)[0], # baseNameInFile)[0].split('_')
                                  'head'   : [column for idx, column in enumerate(rowHead) if idx not in idx2excludeList],
                                  'rows'   : rows
                                }
# />


if filesInfoDict:
    period, node, fdate = ([] for _ in range(3))
    #period = []
    #node   = []
    #fdate  = []
    for file in filesInfoDict:
        period.append(filesInfoDict[file].get('period'))
        node.append(filesInfoDict[file].get('node'))
        fdate.append(filesInfoDict[file].get('fdate'))
    
    period = list(set(period))
    node   = list(set(node))
    fdate  = list(set(fdate))


if len(period) == 1 and len(fdate) == 1:
    # шаблон пути архивирования
    archPathTmpl = '{0}\\{1}\\{2}\\{3}\\{4}_period\\'.format(
                                                                dir.get('arch'),
                                                                '{0}',
                                                                fdate[0].split('_')[1],
                                                                fdate[0].split('_')[0],
                                                                period[0]
                                                            )
    
    # создание структуры каталогов в 'arch'
    for path in ['in', 'out']:
        archPath = archPathTmpl.format(path)
        if not os.path.exists(archPath):
            os.makedirs(archPath)
    
    # архивирование входящих файлов
    for file in inFileList:
        shutil.move(
                      os.path.join(dir.get('in'), file),
                      os.path.join(archPathTmpl.format('in'), file)
                   )


    # Получить список заголовков из словаря filesInfoDict
    heads = [filesInfoDict.get(i).get('head') for i in filesInfoDict.keys()]

    # Сравнение заголовков из списка heads
    # Если заголовки отличаются, то выполнение скрипта прерывается
    head = heads.pop(0)
    for item in heads:
        if not head == item:
            head = None
            break


    # преобразуем заголовок в удобочитаемый вид
    if head:

        for item in range(len(head)):
            if head[item]=='otdelen':
                head[item] = 'Индекс ОПС'
            elif head[item]=='nazv':
                head[item] = 'Название ОПС'
            elif head[item]=='city':
                head[item] = 'Тип ОПС'  # 0 -- село, 1 -- город
            elif head[item]=='d':
                head[item] = 'Всего'
            elif findall(r'[a-z]{1}', head[item])[0] == 'd' and len(head[item]) > 1:
                head[item] = head[item][1:]

        # Инициализация словаря rowsNodesDict пустыми значениями
        rowsNodesDict = {node:{'rows':[],'sum':[Decimal("0.00")]*len(head[2:])} for node in SERVED_NODES}

        for inFile in filesInfoDict.keys():
            for row in filesInfoDict.get(inFile).get('rows'):

                index = str(int(row[0]))
                node = opsDict.get(index).get('node')

                # Запись строки только если последнее значение ненулевое
                if not int(row[-1]) == 0:
                    rowsNodesDict[node].get('rows').append(row)

                for i in range(len(rowsNodesDict.get(node).get('sum'))):
                    rowsNodesDict.get(node).get('sum')[i] += Decimal(row[2:][i])

        # Создние нового Excel-файла с новой рабочей книгой:
        wb = xlwt.Workbook()

        ### === > Создание листа Excel "Общая потребность"
        # Создаем шрифт
        font = xlwt.easyxf('font: height 240,\
                            name Arial,\
                            colour_index black,\
                            bold on,\
                            italic off;\
                            align: wrap on,\
                            vert top,\
                            horiz right;\
                            pattern: pattern solid,\
                            fore_colour yellow;')

        # Создание excel (xls) файла в python
        # https://py-my.ru/post/4e15588e1d41c81105000003/index.html

        # Высота строки
        # sheet.row(1).height = 2500

        # Ширина колонки
        # sheet.col(0).width = 20000

        # Лист в положении "альбом"
        # sheet.portrait = False

        # Масштабирование при печати
        # sheet.set_print_scaling(85)

        # Working with Excel Files in Python
        # http://www.python-excel.org/

        # Как из шаблона exl прочитать данные и сравнить их с ячейкой из шаблона
        # http://python.su/forum/topic/24115/?page=1#post-127313

        # Python для начинающих
        # http://www.cyberforum.ru/python-beginners/thread2498064.html

        # 2750 # ширина столбца 10
        # 3300 # ширина столбца 12
        # 4400 # ширина столбца 16
        # 6050 # ширина столбца 22

        ws = wb.add_sheet('Общая потребность')

        head1 = head.copy() # аналог head[:]
        head1.pop(0)
        head1[0] = 'Узел'

        # Запись заголовка
        for j, item in enumerate(head1, start=0):
            ws.write(0, j, item)
            ws.col(j).width = 3300        # ширина всех столбцов 12
        ws.col(0).width = 2750            # ширина первого столбца 10
        ws.col(len(head1)-1).width = 4400 # ширина последнего столбца 16

        nodes_sum = [Decimal("0.00")] * len(head[2:])
        for i, node in enumerate(rowsNodesDict.keys(), start=1):
            # Запись региона
            ws.write(i, 0, allNodesDict.get(node).get('city'))
            
            # Запись строки сумм
            for k, item in enumerate(rowsNodesDict.get(node).get('sum'), start=1): # start=1 - запись начиная со второй ячейки таблицы
                ws.write(i, k, item.quantize(Decimal(".00")))

            for x in range(len(nodes_sum)):
                nodes_sum[x] += rowsNodesDict.get(node).get('sum')[x]

            ws.row(i).height = 4125 # высота строки 15
        
        # Запись строки сумм
        pos = len(rowsNodesDict.keys()) # номер последней записанной строки
        for k, item in enumerate(nodes_sum, start=1): # start=1 - запись начиная с третьей ячейки таблицы
            ws.write(pos + 2, k, item.quantize(Decimal(".00")), font)
        ### ===

        # Создание листов Excel с наименованием РУПС (УПС)
        for node in rowsNodesDict.keys():
            ws = wb.add_sheet(allNodesDict.get(node).get('label'))

            # Запись заголовка
            for j, item in enumerate(head, start=0): # start=0 - запись в первой строке таблицы
                ws.write(0, j, item)
                ws.col(j).width = 3300        # ширина всех столбцов 12
            ws.col(1).width = 6050            # ширина второго столбца 22

            # Запись значений, отфильтрованных по индексу
            for i, row in enumerate(rowsNodesDict.get(node).get('rows'), start=1): # start=1 - запись начиная со второй строки таблицы
                index = str(int(row[0]))
                nameOPS = opsDict.get(index).get('nameOPS')
                for j, cell in enumerate(row, start=0):
                    if j==1:
                        ws.write(i, 1, nameOPS)
                    else:
                        ws.write(i, j, cell)

            # Запись строки сумм
            pos = len(rowsNodesDict.get(node).get('rows')) # номер последней записанной строки
            for k, item in enumerate(rowsNodesDict.get(node).get('sum'), start=2): # start=2 - запись начиная с третьей ячейки таблицы
                ws.write(pos + 2, k, item.quantize(Decimal(".00")))


        outFileName = OUT_FILE_NAME_TEMPLATE.format(
                                                     period[0],
                                                     monthDict2.get(fdate[0].split('_')[0])[1],
                                                     fdate[0].split('_')[1]
                                                   )
        # Сохранить книгу Excel
        wb.save(os.path.join(dir.get('out'), outFileName))
        # архивирование исходящего файла
        shutil.copy2(
                      os.path.join(dir.get('out'), outFileName),
                      os.path.join(archPathTmpl.format('out'), outFileName)
                    )
