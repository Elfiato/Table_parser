import string

'''
result_dict - словарь, где ключ - название элемента(может быть без ТУ если ТУ написанов следующей строке),
значение - количество элементов.
elem_dict - словарь, где ключ, позиционное обозначение элемента на схеме,
значение - массив, где:
0 эл. - название элемента
1 эл. - количесвто
2 эл.(если есть)  - примечание(тоже только из той строки, где есть название).
'''


def if_count_zero(values):
    """
    Функция вычисляет количество элементов через их позиционное обозначение
    """
    if '-' in values[0]:
        amount_of_el = [el.strip() for el in values[0].split('-')]
    if ',' in values[0]:
        amount_of_el = [el.strip() for el in values[0].split(',')]
    else:
        amount_of_el = [values[0]]
    for char in amount_of_el[0]:
        if char in string.digits:
            position = amount_of_el[0].index(char)
    if len(amount_of_el) == 2 and amount_of_el[1]:
        values[2] = int(amount_of_el[1][position:]) - int(amount_of_el[0][position:]) + 1
    else:
        values[2] = 1
    return values


def add_name(values):
    """
    Функция добавляет тип элеменета к названию
    """
    element_name = {
        'C': 'Конденсатор',
        'L': 'Индуктивность',
        'R': 'Резистор',
        'DA': 'Микросхема аналоговая',
        'DD': 'Микросхема цифровая',
        'VD': 'Диод',
        'VT': 'Транзистор',
        'WZ': 'Фильтр',
        'X': 'Разъем',
        'XW': 'Разъем',
        'С': 'Конденсатор'
    }
    for elem in element_name:
        if values[0].startswith(elem):
            if values[1].lower().startswith(element_name[elem].lower()):
                return values
            else:
                values[1] = element_name[elem] + ' ' + values[1]
                return values


# название исходного и выходного файла писать сюда
name_base_file = 'test_input.csv'
name_write_file = 'test_output.csv'

where_tabs = []
count_letter = -1
result_dict = {}
elem_dict = {}
values = []

with open(name_base_file, 'r', encoding='ANSI') as file:
    for i in file:
        # ищем где ';' добавляем в массив
        if i[0] in string.ascii_letters or i[0] in 'СЕТАХНОВМК':
            for j in i:
                count_letter += 1
                if j == ';':
                    where_tabs.append(count_letter)
        # создаем массив значений с названием элементов и их количеством
        for k in range(len(where_tabs)):
            if k == 0:
                values.append(i[:where_tabs[k]].strip())
            elif k == 3:
                values.append(i[where_tabs[k]:].strip())
            else:
                values.append(i[where_tabs[k - 1] + 1:where_tabs[k]].strip())
        # если массив не пустой, смотрим чтобы количество элементов было проставлено, иначе вычисляем его
        # через позиционные обозначения элементов
        if values:
            add_name(values)
            if result_dict.get(values[1], 0) == 0:
                if values[2] == '':
                    if_count_zero(values)
                    result_dict[values[1]] = int(values[2])
                else:
                    result_dict[values[1]] = int(values[2])
            else:
                if values[2] == '':
                    if_count_zero(values)
                    result_dict[values[1]] += int(values[2])
                else:
                    result_dict[values[1]] += int(values[2])
            elem_dict[values[0]] = values[1:]
        where_tabs = []
        count_letter = -1
        values = []

sort_key = sorted(list(result_dict))

with open(name_write_file, 'w', encoding='ANSI') as w_file:
    for key in sort_key:
        print(f'{key};{result_dict[key]}', file=w_file, end='\n')
