from docx import Document
import xlsxwriter


def get_name_prefix(values: str) -> str:
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
        'С': 'Конденсатор',
    }
    for name in element_name:
        if values.startswith(name):
            return element_name[name]
    return ''


def get_dop_info(data: list, row_num: int, result_dict=None, not_visited_flag=True, name='') -> str:
    """
    Функция сохраняет примечания элементов
    """
    dop_info = ''
    while True:
        if data[row_num][3]:
            dop_info += f'{data[row_num][3]} '
            row_num += 1
        else:
            break
    if not_visited_flag:
        return dop_info.strip()
    else:
        return (result_dict[name][1] + '\n' + dop_info).strip()


def get_full_el_name(data: list, row_num: int) -> tuple[str, int]:
    """
    Функция формирует полное имя элемента, даже если оно находится на нескольких строках
    """
    name = ''
    row_num_start = row_num
    amount_of_str_in_name = 0
    prefix = get_name_prefix(data[row_num][0])
    while True:
        if data[row_num][1]:
            name += f'{data[row_num][1]} '
            row_num += 1
        else:
            break
    if row_num != row_num_start + 1:
        amount_of_str_in_name = row_num - row_num_start - 1
    if not name.startswith(prefix):
        name = prefix + ' ' + name
    return name.strip(), amount_of_str_in_name


def get_amount_of_el(data: list, row_num: int, amount_of_str_in_name: int) -> int:
    """
    Функция  считывает количество элементов
    """
    try:
        amount = int(data[row_num + amount_of_str_in_name][2])
    except ValueError:
        amount = data[row_num + amount_of_str_in_name][2]
    return amount


def make_el_table(data_table: list) -> dict:
    """
    Функция формирует словарь result, где ключ название элемента с учетом его типа, а значение список с количеством
    элементов данного типа и их примечанием.
    {'Название элемента: [Количество, примечания]'}
    """
    result = {}
    row_num = 0
    amount_of_str_in_name = 0
    while row_num < len(data_table):
        if len(data_table[row_num]) == 4:
            if data_table[row_num][1] and data_table[row_num][1] != 'Наименование':
                name, amount_of_str_in_name = get_full_el_name(data_table, row_num)
                amount = get_amount_of_el(data_table, row_num, amount_of_str_in_name)
                if name in result:
                    result[name] = [result[name][0] + amount, get_dop_info(data_table, row_num, result, False, name)]
                else:
                    result[name] = [amount, get_dop_info(data_table, row_num)]
        row_num += 1 + amount_of_str_in_name
        amount_of_str_in_name = 0
    return result


def main():
    # Названия входного и выходного файлов
    input_file = 'PE3.docx'
    output_file = 'test_output.xlsx'

    # считывание данных из исходного перечня
    f = open(input_file, 'rb')
    document = Document(f)
    f.close()

    data = [[]]
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                data[-1].append(cell.text)
            data.append([])

    result = make_el_table(data)

    # запись в выходной файл
    workbook = xlsxwriter.Workbook(output_file)
    worksheet = workbook.add_worksheet()

    r = 0
    for name in result:
        worksheet.write(r, 0, name)
        for i in range(1, 3):
            worksheet.write(r, i, result[name][i - 1])
        r += 1

    workbook.close()
    print('Файл сформирован.')


if __name__ == "__main__":
    main()
