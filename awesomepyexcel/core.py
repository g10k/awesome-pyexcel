# coding: utf-8
import os
import random
import re

import xlrd
import xlrd.sheet
import xlwt
import xlwt.Workbook
import xlwt.Cell

HEADER_HORIZONTAL = 'horizontal'
HEADER_VERTICAL = 'vertical'
HEADER_ORIENTATION = (HEADER_HORIZONTAL, HEADER_VERTICAL)


class Field(object):
    WIDTH_AUTO = 'auto'

    class GetValueException(Exception): pass

    def __init__(self, verbose_name=None, key=None, keys=None, is_counter=False, width=WIDTH_AUTO, style=None,
                 need_sum=False, need_count=False, need_average=False, empty_filler='', header_orientation=None, formula=False):
        """
        Описание колонки таблицы, того как она оформляется и заполняется.
        :param verbose_name: Название поля/атрибута/данных, которые подставляются в колонку.
                             verbose_name задает заголовок таблицы
        :param key: название ключа (для словаря) или трибута (для объекта), по которому
                    можно доступаться до данных для ячейки таблицы.
                    Поддерживаются лукапы: 'client__fio'.
        :param keys: набор ключей, для доступа к данным. На случай, если Field описывает
                    работу с несколькими источниками данных
                    (например, список заявок (orders) и пересдач (re_orders))
                    доступ к данным задачется через словарь вида:
                    {'order': 'client__fio', 're_order': 'order__client__fio'}
        :param width: ширина колонки. Допустимо либо число
                      (пикселей или чего-то такого что использует библиотека xlwt),
                      либо ключевое слово 'auto' (Field.WIDTH_AUTO)
        :param is_counter: Является ли колонка счетчиком строк
        :param need_sum: флаг - нужно ли втсавлять формулу суммы под колонкой таблицы
        :param need_count: строковое значение - если указано, то будет использоваться формула COUNTIF с ним.
        :param need_average: флаг - нужно ли втсавлять формулу среднего значения под колонкой таблицы
        :param empty_filler: заполнитель для пустой ячейки
        :param header_orientation: ориентация для заголовка - горизонтальный (HEADER_HORIZONTAL),
                             или вертикальный (HEADER_VERTICAL). Может быть не задан.
        :param formula: вставка формулы, в доступны значения ячеек той же строки


        """
        # assert key or keys Не работает в случае с is_counter Field-ом
        assert header_orientation is None or header_orientation in HEADER_ORIENTATION
        if need_sum and need_count:
            assert False

        self.verbose_name = verbose_name
        self.key = key
        self.keys = keys
        self.width = width
        self.style = style
        self.is_counter = is_counter
        self.need_sum = need_sum
        self.need_count = need_count
        self.need_average = need_average
        self.empty_filler = empty_filler
        self.header_orientation = header_orientation
        self.formula = formula

    def get_value(self, obj, key_name=None, key=None):
        """
        :param obj: объект из котороо вятягиваем данные: или словарь, или объект, или функция
        :param key_name: названия ключа (для случая поддержки несколих источников данных).
                         Например, если в __init__ в параметре keys передаваля словарь с ключами
                         'order' или 're_order', то в качестве key_name, можно указать 'order' или 're_order'
        :param key: ключ для доступа к данным obj. Если не задан, будет определен автоматически из настроек Field
        :return: данные объекта obj по ключу key
        """
        if not key:
            if key_name:
                key = self.keys[key_name]
            elif self.key:
                key = self.key
            if not key:
                raise self.GetValueException(u'Не удалось получить ключ (key) для доступа к объекту "%s"' % obj)

        if isinstance(obj, dict) and key in obj:
            return obj[key]

        if isinstance(key, basestring):
            # ищем по лукапу
            if '__' in key:
                key_parts = key.split('__')

                attr_name = key_parts[0]
                obj_attr = self.get_value(obj, key=attr_name)

                if obj_attr:
                    look_up_key = '__'.join(key_parts[1:])
                    value = self.get_value(obj_attr, key=look_up_key)
                else:
                    value = None

            # данные из словаря
            elif isinstance(obj, dict):
                value = obj.get(key)

            # данные из атрибута объекта
            elif hasattr(obj, key):
                value = getattr(obj, key)

            else:
                value = self.empty_filler

        # ключ генерит функция
        elif hasattr(key, '__call__'):
            value = key(obj)

        else:
            raise self.GetValueException(
                u'Не смог получить значения для объекта "%s" по ключу "%s"' % (obj, key)
            )

        # если в качестве значени получили функцию, вызовем ее для получения значения ячейки
        if hasattr(value, '__call__'):
            value = value()
        if value is None:
            return self.empty_filler
        return value


class Sheet(xlwt.Worksheet):
    class WriteSheetException(Exception): pass

    need_top_part = True

    logo_rows = None    # кол-во строк для логотипа
    logo_width = None    # ширина логотипа в единицах, которая применяется в методе

    DEFAULT_HEADER_HEIGHT_VERTICAL = 2000  # Высота заголовка в вертикальной ориентации
    DEFAULT_HEADER_HEIGHT_HORIZONTAL = 255  # Высота заголовка в горизонтальной ориентации
    MIN_COL_WIDTH = 800  # Минимальная ширина столбца. Используется для № (counter) или для колонки с вертикальной ориентацией
    MAX_COUNT_CHAR_IN_VERTICAL_LINE = 22  # Максимальное кол-во символов в одной строке вертикального заголовка
    ONE_LETTER_HEIGHT = 90 # Высота для одной буквы в вертикальной ориентации
    ONE_LETTER_WIDTH = 350  # Ширина одной буквы в горизонтальной ориентации
    AUTO_COLUMN_WIDTH = 2938  # примерное значение ширины столбца, которое создает xlwt по умолчанию
    MIN_TITLE_WIDTH = 21000  # Минимальная ширина для title
    max_lines_in_row = 10
    COEFF = 36.72  # Коэффициент соответствия width и 36.72 px

    title_style = xlwt.easyxf(
        'font: height 300, name Times New Roman, bold True;'
        'align: vertical center, horizontal center;'
        'alignment: wrap on;'
        'borders: left thin, right thin, top thin, bottom thin;')
    under_picture_style = xlwt.easyxf(
        'font: height 220, name Times New Roman, bold True;'
        'align: vertical center, horizontal center;'
        'borders: left thin, right thin, top thin, bottom thin;')
    vertical_header_style = xlwt.easyxf(
        'alignment: wrap on, rota -90;'
        'font: name Times New Roman, bold True;'
        'borders: left thin, right thin, top thin, bottom thin;'
    )
    horizontal_header_style = xlwt.easyxf(
        'font: name Times New Roman, bold True;'
        'alignment: wrap on;'
        'borders: left thin, right thin, top thin, bottom thin;'
    )
    signature_style = xlwt.easyxf(
        'align:  horiz center;'
        'alignment: wrap on;'
        'font: italic on;'
    )
    table_style = xlwt.easyxf(
        'font: name Times New Roman;'
        'alignment: wrap on;'
        'borders: left thin, right thin, top thin, bottom thin;'
    )

    def __init__(self, *args, **kwargs):
        super(Sheet, self).__init__(*args, **kwargs)
        self.current_row_i = 0  # Индекс строки, в которую ведется запись
        self._data_start_row_i = 0  #  Индекс строки, с которой наинается содержимое таблицы
        self.calculate_max_count_char_in_vertical_line()

    def calculate_max_count_char_in_vertical_line(self):
        self.MAX_COUNT_CHAR_IN_VERTICAL_LINE = self.parent.header_height / self.ONE_LETTER_HEIGHT

    def get_logo_path(self):
        raise NotImplementedError()

    def write_header(self):
        self._set_cols_widths()
        if self.parent.need_top_part:
            self._write_top_part()
        self._write_table_header()

    def _parse_formula(self, formula):
        """Формат формулы 2col/5col + 5 заменяем на
           B(current_row)/E(current_row) + 5
        """
        for column_text, col_number in re.findall(r'((\d+)col)', formula):
            letter = xlrd.colname(int(col_number))
            formula = formula.replace(column_text, '%s%s' % (letter, self.current_row_i+1))
        return formula

    def write_table_body(self, objects, key_name=None):
        """
        запишем данные
        :param objects:
        :param key_name:
        """
        for counter, obj in enumerate(objects, start=1):
            max_height = 0
            for col_i, field in enumerate(self.parent.headers):
                style = field.style if field.style else self.table_style
                if field.is_counter:
                    self.write(self.current_row_i, col_i, counter, style)
                else:
                    value = field.get_value(obj, key_name=key_name)
                    self.write(self.current_row_i, col_i, value, style)
                    if not self.parent.constant_row_height and self.parent.calculate_row_heights:
                        cell_height = self._get_row_height(unicode(value), width=self.col(col_i).width)
                        if cell_height > max_height:
                            max_height = cell_height

            if self.parent.constant_row_height:
                self.row(self.current_row_i).height = self.parent.constant_row_height
            elif self.parent.calculate_row_heights:
                self.row(self.current_row_i).height = max_height

            self.current_row_i += 1

    def write_footer(self):
        self._write_sum()
        if self.parent.need_signature:
            self.write_signature()

    def write_signature(self):
        """
        подпись под таблицей
        """
        pass

    def get_title(self):
        if hasattr(self, 'title'):
            return self.title
        return self.parent.title

    def _write_top_part(self):
        self._insert_logo()
        self.current_row_i = self.logo_rows
        logo_cols = self._get_logo_cols()

        # Вставка текста справа от картинки
        title_col_count = self._get_cols_for_title()
        title_col_i = logo_cols + 1

        self.write_merge(0, self.current_row_i,
                         title_col_i,  title_col_i + title_col_count,
                         self.get_title(),
                         self.title_style)

        # Вставка текста под картинкой
        self.write_merge(
            self.current_row_i, self.current_row_i,
            0, logo_cols,
            self.parent.subscription,
            self.under_picture_style)

        self.current_row_i += 1

    def _write_table_header(self):
        for col_number, field in enumerate(self.parent.headers):
            if field.header_orientation == HEADER_VERTICAL:
                header_style = self.vertical_header_style
            elif field.header_orientation == HEADER_HORIZONTAL:
                header_style = self.horizontal_header_style
            else:
                header_style = self._get_header_style()

            self.write(self.current_row_i, col_number, field.verbose_name, header_style)

        self.row(self.current_row_i).height = self._get_header_height()

        self.current_row_i += 1
        self._data_start_row_i = self.current_row_i

    def _write_sum(self):
        for col, field in enumerate(self.parent.headers):
            if field.need_sum or field.need_count or field.need_average or field.formula:
                letter = xlrd.colname(col)
                if field.formula:
                    formula = xlwt.Formula(self._parse_formula(field.formula))
                else:
                    if field.need_sum:
                        function_ = u'SUM(%s%s:%s%s)'
                    elif field.need_average:
                        function_ = u'ROUND(AVERAGE(%s%s:%s%s);0)'
                    else:
                        function_ = u'COUNTIF(%s%s:%s%s,"{condition}")'.format(condition=field.need_count)
                    formula = xlwt.Formula(
                        function_ % (
                            letter, self._data_start_row_i+1, letter, self.current_row_i
                        ))
                self.write(self.current_row_i, col, formula)
        self.current_row_i += 1

    def _get_header_style(self):
        if self.parent.header_orientation == HEADER_HORIZONTAL:
            return self.horizontal_header_style
        elif self.parent.header_orientation == HEADER_VERTICAL:
            return self.vertical_header_style
        else:
            return self.horizontal_header_style

    def _get_header_height(self):
        if self.parent.is_set_header_row_height:
            header = min(self.parent.headers, key=lambda x: x.width/len(x.verbose_name))
            return self._get_row_height(unicode(header.verbose_name))
        elif self.parent.header_height:
            return self.parent.header_height
        elif self.parent.header_orientation == HEADER_VERTICAL:
            return self.DEFAULT_HEADER_HEIGHT_VERTICAL
        else:
            return self.DEFAULT_HEADER_HEIGHT_HORIZONTAL

    def _get_logo_x_coordinate(self):
        cols_width = 0
        for col_index in range(self._get_logo_cols() + 1):
            cols_width += self.col(col_index).width

        if cols_width > self.logo_width:
            start_x_coordinate = (cols_width - self.logo_width) / 2
            start_x_coordinate /= self.COEFF
        else:
            start_x_coordinate = 0
        return start_x_coordinate

    def _insert_logo(self):
        """
            Функция вставки логотипа в отчет
        """
        logo_path = self.get_logo_path()
        if os.path.isfile(logo_path):
            start_x_coordinate = self._get_logo_x_coordinate()
            self.merge(0, self.logo_rows-1, 0, self._get_logo_cols())
            self.insert_bitmap(logo_path, 0, 0, start_x_coordinate, 0, 1, 1)
        else:
            raise self.WriteSheetException(u'Не правильный путь к логотипу "%s"' % logo_path)

    def _get_row_height(self, value, width=None):
        """
        Посчитаем какую ширину нужно поставить строке Excel, в зависимости от количества строк в ячейке
        """
        if width:
            line_counts = int(((len(value) * self.ONE_LETTER_WIDTH) / width)) + 1
        else:
            line_counts = value.count('\n') + 1
            line_counts = min(line_counts, self.max_lines_in_row)
        return line_counts * self.DEFAULT_HEADER_HEIGHT_HORIZONTAL

    def _get_enough_columns_count(self, need_width, start_column=0, max_columns_count=50):
        """Получить количество столбцов, ширина которых будет больше чем need_width"""
        sum_cols_widths, index = 0, start_column
        while 1:
            sum_cols_widths += self.col(index).width
            if sum_cols_widths >= need_width or index > max_columns_count:
                break
            index += 1
        return index - start_column

    def _get_logo_cols(self):
        """
        Количество колонок, необходимых для логотипа.
        """
        width = max(self.logo_width, len(self.parent.subscription) * self.ONE_LETTER_WIDTH)
        return self._get_enough_columns_count(width)

    def _get_cols_for_title(self):
        """
        количество колонок для title
        """
        logo_cols = self._get_logo_cols()
        return self._get_enough_columns_count(self.MIN_TITLE_WIDTH, start_column=logo_cols + 1)

    def _set_cols_widths(self):
        """
        устанавливает ширину столбцов
        :return:
        """
        cols_widths = self._calc_cols_widths()
        for num, col_width in enumerate(cols_widths):
            self.col(num).width = col_width

    def _calc_cols_widths(self):
        """
        расчитывает ширину столбцов
        :return:
        """
        widths = []
        for field in self.parent.headers:
            if field.is_counter:
                if isinstance(field.width, int):
                    widths.append(field.width)
                else:
                    widths.append(self.MIN_COL_WIDTH)

            elif field.width == Field.WIDTH_AUTO:
                header_orientation = field.header_orientation or self.parent.header_orientation
                verbose_name_len = len(field.verbose_name)

                # вертикальный заголовок
                if header_orientation == HEADER_VERTICAL:
                    # если заголов поместился в одну строку
                    if verbose_name_len < self.MAX_COUNT_CHAR_IN_VERTICAL_LINE:
                        widths.append(self.MIN_COL_WIDTH)

                    # многострочный заголовок
                    else:
                        line_count = int(verbose_name_len // self.MAX_COUNT_CHAR_IN_VERTICAL_LINE) + 1
                        many_lines_coefficient = 0.8    # а то сильно широкая колонка получается
                        length = line_count * self.MIN_COL_WIDTH * many_lines_coefficient
                        widths.append(length)

                # горизонтальный
                else:
                    if field.width is int:
                        widths.append(field.width)
                    else:
                        widths.append(verbose_name_len * self.ONE_LETTER_WIDTH)

            elif field.width:
                widths.append(field.width)
            else:
                raise Exception('o_O')

        return widths


class SaveBookMixin(object):

    def save(self, filename=None, return_file_path=False):
        """
        Сохраняет книгу в файл.
        :param filename: имя файла. Если не указана, будет сформировано случайное число.
        :param return_file_path: флаг определяющий нужно ли возвращать путь к файлу
        :return: название файли или название файла и полный путь к нему (если задан флаг return_file_path)
        """
        file_name = '%s.xls' % random.randint(1, 999999)
        super(SaveBookMixin, self).save(file_name)
        return file_name

class Book(SaveBookMixin, xlwt.Workbook):
    title = u'Имя не указано'
    need_top_part = False
    need_signature = False
    is_set_header_row_height = False
    headers = []
    header_orientation = HEADER_HORIZONTAL
    header_height = 255
    sheet_class = Sheet
    # Считать height для ячеек строки и ставить строке максимальный из них. Только для тела таблицы
    calculate_row_heights = False
    constant_row_height = None

    def __init__(self, data=None, create=True, metadata=None, *args, **kwargs):
        """

        :param data: итерируемый объект с данными отчета, которые будут почещены в тело таблицы.
        :param create: создавать excel-книгу сразу, или просто инициализировать объект.
                       Имеет смысл, если перед созданием книги необходимо дополнительно ее сконфигурирвать
                       (например задать заголовки таблицы и т.п.)
        :param metadata: дополнительные данные об отчете. Например, период за который формируется отчет.
                         В качестве значения можно передавать cleaned_data формы.
        :param args:
        :param kwargs:
        """
        super(Book, self).__init__(*args, **kwargs)
        self.headers = list(self.headers)  # копируем
        self.metadata = metadata or {}
        if create:
            self.create(data)

    def create(self, data):
        """
        заполняет книгу данными
        :param data: итерируемый объект с данными отчета
        """
        sheet = self.add_sheet(self.title)
        sheet.write_header()
        self.process_data(data, sheet)
        sheet.write_footer()

    def process_data(self, data, sheet):
        """
        Заполняет данные одной старицы.
        Правильное место для переопределения хитрого вывода данных в таблицу.
        Например, если таблица выводить два списка объектов: сначала заявки, а потом пересдачи.
        :param data:
        :param sheet:
        :return:
        """
        sheet.write_table_body(data)

    def add_sheet(self, sheetname, cell_overwrite_ok=False):
        """ Взято из Workbook. все переменные __worksheets заменены на _Workbook__worksheets"""
        if not isinstance(sheetname, unicode):
            sheetname = sheetname.decode(self.encoding)
        lower_name = sheetname.lower()
        if lower_name in self._Workbook__worksheet_idx_from_name:
            raise Exception("duplicate worksheet name %r" % sheetname)
        self._Workbook__worksheet_idx_from_name[lower_name] = len(self._Workbook__worksheets)

        sheet = self.sheet_class(sheetname, self, cell_overwrite_ok)
        self._Workbook__worksheets.append(sheet)
        return sheet