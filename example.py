#! /usr/bin/env python3
# -*- coding: utf-8 -*-

import libreoffice_wrapper as lw
import importlib
importlib.reload(lw)

pid = lw.start_soffice()
soffice = lw.soffice()

# %%
writer = soffice.Writer()
writer.save()
writer.close()

impress = soffice.Impress()
impress.save()
impress.close()

draw = soffice.Draw()
draw.save()
draw.close()

math = soffice.Math()
math.save()
math.close()

# %%
# %%
calc = soffice.Calc()
sheet = calc.get_sheet_by_position(0)

# %%
sheet.set_row_height([0, 1, 2, 3, 4, 5, 6, 7, 8 , 9], [10, 20, 30, 40, 500, 60, 70, 80, 90, 20])
sheet.cell_properties(1, 1)

sheet.get_cell_property(2, 2, 'CellBackColor')
sheet.set_cell_property(5, 5, 'CellBackColor', 16776960)
sheet.get_cell_property(2, 2, 'CellBackColor')

sheet.get_cell_property(2, 2, 'TopBorder')
sheet.set_cell_property(2, 2, 'TopBorder.LineWidth', 10)
sheet.get_cell_property(2, 2, 'TopBorder')

d = sheet.get_cell_property(5, 5, 'TopBorder')
d['LineWidth'] = 7
sheet.set_cell_property(5, 5, 'TopBorder', d)
sheet.get_cell_property(5, 5, 'TopBorder')

# %%
c.close()
kill(p)
s.kill()
kill_tmux()
# %%









# %%

class sheet_dump():

    def set_cell(self, *args, **kwargs):
        """
        kwargs trumps args

        args order:
        column
        row
        value
        format

        or

        column
        row
        value

        or

        cell
        value
        format
        """
        format = 'string'

        if len(args) == 2:
            column, row = cell2num(args[0])
            value = args[1]

        elif len(args) == 3:
            try:
                column, row = cell2num(args[0])
                value = args[1]
                format = args[2]
            except AttributeError:
                column = args[0]
                row = args[1]
                value = args[2]

        elif len(args) == 4:
            column = args[0]
            row = args[1]
            value = args[2]
            format = args[3]

        if 'column' in kwargs:
            column = kwargs['column']
        if 'row' in kwargs:
            row = kwargs['row']
        if 'cell' in kwargs:
            column, row = cell2num(kwargs['cell'])
        if 'value' in kwargs:
            value = kwargs['value']
        if 'format' in kwargs:
            format = kwargs['format']

        row = _check_row_value(row)
        column = _check_column_value(column)

        if format == 'formula':
            self.write(f"sheet_{self.tag}.getCellByPosition({column}, {row}).setFormula('{value}')")
        elif format == 'string':
            self.write(f"sheet_{self.tag}.getCellByPosition({column}, {row}).setString('{value}')")
        elif format == 'number':
            self.write(f"sheet_{self.tag}.getCellByPosition({column}, {row}).setValue({value})")
        else:
            raise ValueError(f"{format} is not a valid format (valid formats: 'formula', 'string', 'number').")

    def get_cell(self, *args, **kwargs):
        """
        kwargs trumps args

        args order:
        column
        row
        format

        or

        column
        row

        or

        cell
        format
        """
        format = 'string'

        if len(args) == 1:
            column, row = cell2num(args[0])
        elif len(args) == 2:
            try:
                column, row = cell2num(args[0])
                format = args[1]
            except AttributeError:
                column = args[0]
                row = args[1]
        elif len(args) == 3:
            column = args[0]
            row = args[1]
            format = args[2]

        if 'column' in kwargs:
            column = kwargs['column']
        if 'row' in kwargs:
            row = kwargs['row']
        if 'cell' in kwargs:
            column, row = cell2num(kwargs['cell'])
        if 'format' in kwargs:
            format = kwargs['format']

        row = _check_row_value(row)
        column = _check_column_value(column)

        if format == 'formula':
            value =  self.write(f"print(sheet_{self.tag}.getCellByPosition({column}, {row}).getFormula())")
            if len(value) == 1: return value[0]
            else: return ''
        elif format == 'string':
            value = self.write(f"print(sheet_{self.tag}.getCellByPosition({column}, {row}).getString())")
            if len(value) == 1: return value[0]
            else: return ''
        elif format == 'number':
            return float(self.write(f"print(sheet_{self.tag}.getCellByPosition({column}, {row}).getValue())")[0])
        else:
            raise ValueError(f"{format} is not a valid format (valid formats: 'formula', 'string', 'number').")

    def set_cells(self, *args, **kwargs):
        """
        formula set formulas, but also set numbers fine. Dates and time not so much because it changes the formating (if setting date and time with formula you might wanna format the
        cell like date or time using copy_cells to copy formatting).

        string (data) works fine with date, time and number, but formulas are set as string. Therefore, formulas do not work.

        value (data_number) works fine for numbers ONLY.



        value

        range
        value

        cell init
        value

        value
        format

        range
        value
        format

        cell init
        value
        format

        column_start
        row_start
        value

        column_start
        row_start
        value
        format
        """
        format = 'formula'
        column_start = 0
        row_start    = 0


        if len(args) == 1:
            value = args[0]

        elif len(args) == 2:
            try:
                column_start, row_start = cell2num(args[0])
                value = args[1]
            except ValueError:
                column_start, row_start, _, _ = range2num(args[0])
                value = args[1]
            except TypeError:
                value = args[0]
                format = args[1]

        elif len(args) == 3:
            try:
                column_start, row_start = cell2num(args[0])
                value = args[1]
                format = args[2]
            except ValueError:
                column_start, row_start, _, _ = range2num(args[0])
                value = args[1]
                format = args[2]
            except AttributeError:
                column_start = args[0]
                row_start    = args[1]
                value = args[2]
            # except TypeError:

        elif len(args) == 4:
            column_start = args[0]
            row_start    = args[1]
            value = args[2]
            format = args[3]


        if 'column_start' in kwargs:
            column_start = kwargs['column_start']
        if 'row_start' in kwargs:
            row_start = kwargs['row_start']
        if 'cell_start' in kwargs:
            column_start, row_start = cell2num(kwargs['cell_start'])
        if 'range' in kwargs:
            column_start, row_start, column_stop, row_stop = range2num(kwargs['range'])
        if 'value' in kwargs:
            value = kwargs['value']
        if 'format' in kwargs:
            format = kwargs['format']

        # check if value is valid
        if isinstance(value, np.ndarray):
            value = value.tolist()

        if isinstance(value, list):
            for v in value:
                if not isinstance(v, list):
                    raise TypeError('value must be a list of lists or a numpy array')
        else:
            raise TypeError('value must be a list of lists or a numpy array')

        for v in value:
            if len(value[0]) != len(v):
                raise ValueError('value must be a square matrix')

        column_start = _check_column_value(column_start)
        row_start    = _check_row_value(row_start)
        column_stop  = column_start + len(value[0]) - 1
        row_stop     = row_start + len(value) - 1

        if format == 'formula':
            self.write(f"sheet_{self.tag}.getCellRangeByPosition({column_start}, {row_start}, {column_stop}, {row_stop}).setFormulaArray({value})")
        elif format == 'string':
            self.write(f"sheet_{self.tag}.getCellRangeByPosition({column_start}, {row_start}, {column_stop}, {row_stop}).setDataArray({value})")
        elif format == 'number':
            self.write(f"sheet_{self.tag}.getCellRangeByPosition({column_start}, {row_start}, {column_stop}, {row_stop}).setDataArray({value})")
        else:
            raise ValueError(f"{format} is not a valid format (valid formats: 'formula', 'string', 'number').")

    def get_cells(self, *args, **kwargs):
        """
        nothing

        range

        cell init

        format

        range
        format

        cell init
        format

        column_start
        row_start

        cell init
        cell end
        format

        column_start
        row_start
        format

        column_start
        row_start
        column_stop
        row_stop

        column_start
        row_start
        column_stop
        row_stop
        format
        """
        format = 'string'
        column_start = 0
        row_start    = 0


        if len(args) == 1:
            try:
                column_start, row_start = cell2num(args[0])
            except ValueError:
                try:
                    column_start, row_start, column_stop, row_stop = range2num(args[0])
                except ValueError:
                    format = args[0]

        elif len(args) == 2:
            try:
                column_start, row_start = cell2num(args[0])
                format = args[1]
            except ValueError:
                try:
                    column_start, row_start, column_stop, row_stop = range2num(args[0])
                    format = args[1]
                except IndexError:
                    column_start = args[0]
                    row_start    = args[1]

        elif len(args) == 3:
            try:
                column_start, row_start = cell2num(args[0])
                column_stop, row_stop   = cell2num(args[1])
                format = args[2]
            except AttributeError:
                column_start = args[0]
                row_start    = args[1]
                format = args[2]
        elif len(args) == 4:
            column_start = args[0]
            row_start    = args[1]
            column_stop  = args[2]
            row_stop     = args[3]
        elif len(args) == 5:
            column_start = args[0]
            row_start    = args[1]
            column_stop  = args[2]
            row_stop     = args[3]
            format = args[4]


        if 'column_start' in kwargs:
            column_start = kwargs['column_start']
        if 'row_start' in kwargs:
            row_start = kwargs['row_start']
        if 'column_stop' in kwargs:
            column_stop = kwargs['column_stop']
        if 'row_stop' in kwargs:
            row_stop = kwargs['row_stop']
        if 'cell_start' in kwargs:
            column_start, row_start = cell2num(kwargs['cell_start'])
        if 'cell_stop' in kwargs:
            column_stop, row_stop = cell2num(kwargs['cell_stop'])
        if 'range' in kwargs:
            column_start, row_start, column_stop, row_stop = range2num(kwargs['range'])
        if 'format' in kwargs:
            format = kwargs['format']

        row_start = _check_row_value(row_start)
        column_start = _check_column_value(column_start)
        try:
            row_stop = _check_row_value(row_stop)
        except NameError:
            row_stop     = self.get_last_row()
        try:
            column_stop = _check_column_value(column_stop)
        except NameError:
            column_stop  = self.get_last_column()

        if column_stop < column_start:
            raise ValueError(f'column_start ({column_start}) cannot be bigger than column_stop ({column_stop})')
        if row_stop < row_start:
            raise ValueError(f'row_start ({row_start}) cannot be bigger than row_stop ({row_stop})')

        if format == 'formula':
            return list(eval(self.write(f"print(sheet_{self.tag}.getCellRangeByPosition({column_start}, {row_start}, {column_stop}, {row_stop}).getFormulaArray())")[0]))
        elif format == 'string':
            return list(eval(self.write(f"print(sheet_{self.tag}.getCellRangeByPosition({column_start}, {row_start}, {column_stop}, {row_stop}).getDataArray())")[0]))
        elif format == 'number':
            return list(eval(self.write(f"print(sheet_{self.tag}.getCellRangeByPosition({column_start}, {row_start}, {column_stop}, {row_stop}).getDataArray())")[0]))
        else:
            raise ValueError(f"{format} is not a valid format (valid formats: 'formula', 'string', 'number').")

    def _get_property_recursive_old(self, column, row, name, attrs=[]):

        i = eval(self.write(f"print([x.Name for x in sheet_{self.tag}.getCellByPosition({column}, {row}).getPropertySetInfo().Properties])")[0]).index(name)
        f=dict()
        if attrs == []:
            try:

                if "<class 'str'>" == self.write(f"print(type(sheet_{self.tag}.getCellByPosition({column}, {row}).getPropertyValue('{name}').value))")[0]:
                    return  self.write(f"print(sheet_{self.tag}.getCellByPosition({column}, {row}).getPropertyValue('{name}').value)")[0]
                else:
                    keys = self.write(f"for e in sheet_{self.tag}.getCellByPosition({column}, {row}).getPropertyValue('{name}').value.__dir__(): print(e)")
            except soffice_python_error:
                if any(x in self.write(f"print(sheet_{self.tag}.getCellByPosition({column}, {row}).getPropertySetInfo().Properties[{i}].Type)")[0] for x in ('float', 'unsigned hyper', 'long', 'short')):
                        value = self.write(f"print(sheet_{self.tag}.getCellByPosition({column}, {row}).getPropertyValue('{name}'))")
                        if len(value) == 0:
                            return None
                        elif len(value) == 1:
                            if value[0] == 'None':
                                return None
                            elif "<class 'int'>" == self.write(f"print(type(sheet_{self.tag}.getCellByPosition({column}, {row}).getPropertyValue('{name}')))")[0]:
                                return int(value[0])
                            else:
                                return float(value[0])
                elif 'boolean' in self.write(f"print(sheet_{self.tag}.getCellByPosition({column}, {row}).getPropertySetInfo().Properties[{i}].Type)")[0]:
                    return str2bool(self.write(f"print(sheet_{self.tag}.getCellByPosition({column}, {row}).getPropertyValue('{name}'))")[0])
                else:
                    value = self.write(f"print(sheet_{self.tag}.getCellByPosition({column}, {row}).getPropertyValue('{name}'))")
                    if len(value)==0:
                        return None
                    elif len(value)==1:
                        if value == 'None':
                            return None
                        else:
                            return value[0]
        else:
            try:
                t = ''
                for attr in attrs:
                    t += f"__getattr__('{attr}')."
                    if "<class 'str'>" == self.write(f"print(type(sheet_{self.tag}.getCellByPosition({column}, {row}).getPropertyValue('{name}').{t[:-1]}.value))")[0]:
                        return self.write(f"print(sheet_{self.tag}.getCellByPosition({column}, {row}).getPropertyValue('{name}').{t[:-1]}.value)")[0]
                    else:
                        keys = self.write(f"for e in sheet_{self.tag}.getCellByPosition({column}, {row}).getPropertyValue('{name}').{t}value.__dir__(): print(e)")
            except soffice_python_error:
                if "<class 'int'>" == self.write(f"print(type(sheet_{self.tag}.getCellByPosition({column}, {row}).getPropertyValue('{name}').{t[:-1]}))")[0]:
                    return int(self.write(f"print(sheet_{self.tag}.getCellByPosition({column}, {row}).getPropertyValue('{name}').{t[:-1]})")[0])
                elif "<class 'bool'>" == self.write(f"print(type(sheet_{self.tag}.getCellByPosition({column}, {row}).getPropertyValue('{name}').{t[:-1]}))")[0]:
                    return str2bool(self.write(f"print(sheet_{self.tag}.getCellByPosition({column}, {row}).getPropertyValue('{name}').{t[:-1]})")[0])
                else:
                    value = self.write(f"print(sheet_{self.tag}.getCellByPosition({column}, {row}).getPropertyValue('{name}').{t[:-1]})")
                    if len(value)==0:
                        return None
                    elif len(value)==1:
                        if value == 'None':
                            return None
                        else:
                            return value[0]

        for key in keys:
            f[key] = self._get_property_recursive_old(column, row, name, attrs+[key])
        return f

    def _set_cell_property(self, *args, **kwargs):
        if len(args) == 3:
            column, row = cell2num(args[0])
            name = args[1]
            value = args[2]
        elif len(args) == 4:
            column = args[0]
            row = args[1]
            name = args[2]
            value = args[3]

        if 'column' in kwargs:
            column = kwargs['column']
        if 'row' in kwargs:
            row = kwargs['row']
        if 'cell' in kwargs:
            column, row = cell2num(kwargs['cell'])
        if 'value' in kwargs:
            value = kwargs['value']
        if 'name' in kwargs:
            value = kwargs['name']

        column = _check_column_value(column)
        row = _check_row_value(row)

        if '.' not in name:

            if isinstance(value, list) or isinstance(value, np.ndarray):
                raise ValueError(f'Property ({name}) seems to be simple valued, but new value ({value}) seems to be a list/array/dict.')
            elif isinstance(value, dict):
                d = self.get_cell_property(column, row, name)
                for k, v in value.items():
                    if k in d:
                        d[k] = v
                    else:
                        raise ValueError(f'{k} is not a valid option of {name}.')

                func = self._get_property_function(column, row, name)
                f_string = f"{func}(" + ','.join([f"{k}={v}" for k, v in d.items()]) + ")"
                self.write(f"sheet_{self.tag}.getCellByPosition({column}, {row}).setPropertyValue('{name}', {f_string})")
            else:
                self.write(f"sheet_{self.tag}.getCellByPosition({column}, {row}).setPropertyValue('{name}', {value})")


        elif name.count('.') == 1:
            name0 = name.split('.')[0]
            name1 = name.split('.')[1]

            func = self._get_property_function(column, row, name0)

            if func is None:
                raise ValueError(f'Property name ({name}) seems to be a nested property but actual property is not nested.')
            else:
                d = self.get_cell_property(column, row, name0)
                if name1 in d:
                    d[name1] = value
                else:
                    raise ValueError(f'{name1} is not a valid option of {name0}.')

                f_string = f"{func}(" + ','.join([f"{k}={v}" for k, v in d.items()]) + ")"
                self.write(f"sheet_{self.tag}.getCellByPosition({column}, {row}).setPropertyValue('{name0}', {f_string})")
        else:
            raise ValueError(f'Only single nested properties can be edited. Note that complex properties can always be edited through other simple nested properties.')

    def _get_cell_property(self, *args, **kwargs):
        """
        cannot get values from here:

        UserDefinedAttributes
        NumberingRules
        Validation
        ValidationLocal
        ValidationXML
        ConditionalFormat
        ConditionalFormatLocal
        ConditionalFormatXML
        """
        if len(args) == 2:
            column, row = cell2num(args[0])
            name = args[1]
        elif len(args) == 3:
            column = args[0]
            row = args[1]
            name = args[2]

        if 'column' in kwargs:
            column = kwargs['column']
        if 'row' in kwargs:
            row = kwargs['row']
        if 'cell' in kwargs:
            column, row = cell2num(kwargs['cell'])
        if 'name' in kwargs:
            value = kwargs['name']

        column = _check_column_value(column)
        row = _check_row_value(row)

        t = self.write(f"print(type(get_cell_property_recursively({self.tag}, {column}, {row}, '{name}')))")[0]
        try:
            output = self.write(f"print(get_cell_property_recursively({self.tag}, {column}, {row}, '{name}'))")[0]
        except IndexError:
            return None

        if  t == "<class 'int'>" :
            return int(output)
        elif t == "<class 'bool'>" :
            return str2bool(output)
        elif t == "<class 'dict'>":
            return eval(output)
        else:
            return output

    def _set_cells_property(self, *args, **kwargs):
        column_start = 0
        row_start    = 0

        if len(args) == 3:
            try:
                column_start, row_start = cell2num(args[0])
                name = args[1]
                value = args[2]
            except ValueError:
                column_start, row_start, column_stop, row_stop = range2num(args[0])
                name = args[1]
                value = args[2]
        elif len(args) == 4:
            try:
                column_start, row_start = cell2num(args[0])
                column_stop, row_stop   = cell2num(args[1])
                name = args[2]
                value = args[3]
            except AttributeError:
                column_start = args[0]
                row_start    = args[1]
                name = args[2]
                value = args[3]
        # elif len(args) == 4:
        #     column_start = args[0]
        #     row_start    = args[1]
        #     column_stop  = args[2]
        #     row_stop     = args[3]
        elif len(args) == 6:
            column_start = args[0]
            row_start    = args[1]
            column_stop  = args[2]
            row_stop     = args[3]
            name = args[4]
            value = args[5]


        if 'column_start' in kwargs:
            column_start = kwargs['column_start']
        if 'row_start' in kwargs:
            row_start = kwargs['row_start']
        if 'column_stop' in kwargs:
            column_stop = kwargs['column_stop']
        if 'row_stop' in kwargs:
            row_stop = kwargs['row_stop']
        if 'cell_start' in kwargs:
            column_start, row_start = cell2num(kwargs['cell_start'])
        if 'cell_stop' in kwargs:
            column_stop, row_stop = cell2num(kwargs['cell_stop'])
        if 'range' in kwargs:
            column_start, row_start, column_stop, row_stop = range2num(kwargs['range'])
        if 'name' in kwargs:
            name = kwargs['name']
        if 'value' in kwargs:
            value = kwargs['value']

        row_start = _check_row_value(row_start)
        column_start = _check_column_value(column_start)
        try:
            row_stop = _check_row_value(row_stop)
        except NameError:
            row_stop     = self.get_last_row()
        try:
            column_stop = _check_column_value(column_stop)
        except NameError:
            column_stop  = self.get_last_column()

        if column_stop < column_start:
            raise ValueError(f'column_start ({column_start}) cannot be bigger than column_stop ({column_stop})')
        if row_stop < row_start:
            raise ValueError(f'row_start ({row_start}) cannot be bigger than row_stop ({row_stop})')

        if '.' not in name:

            if isinstance(value, list) or isinstance(value, np.ndarray):
                raise ValueError(f'Property ({name}) seems to be simple valued, but new value ({value}) seems to be a list/array/dict.')
            elif isinstance(value, dict):
                d = self.get_cell_property(column, row, name)
                for k, v in value.items():
                    if k in d:
                        d[k] = v
                    else:
                        raise ValueError(f'{k} is not a valid option of {name}.')

                func = self._get_property_function(column, row, name)
                f_string = f"{func}(" + ','.join([f"{k}={v}" for k, v in d.items()]) + ")"
                self.write(f"sheet_{self.tag}.getCellRangeByPosition({column_start}, {row_start}, {column_stop}, {row_stop}).setPropertyValue('{name}', {f_string})")
            else:
                self.write(f"sheet_{self.tag}.getCellRangeByPosition({column_start}, {row_start}, {column_stop}, {row_stop}).setPropertyValue('{name}', {value})")


        elif name.count('.') == 1:
            name0 = name.split('.')[0]
            name1 = name.split('.')[1]

            func = self._get_property_function(column, row, name0)

            if func is None:
                raise ValueError(f'Property name ({name}) seems to be a nested property but actual property is not nested.')
            else:
                d = self.get_cell_property(column, row, name0)
                if name1 in d:
                    d[name1] = value
                else:
                    raise ValueError(f'{name1} is not a valid option of {name0}.')

                f_string = f"{func}(" + ','.join([f"{k}={v}" for k, v in d.items()]) + ")"
                self.write(f"sheet_{self.tag}.getCellRangeByPosition({column_start}, {row_start}, {column_stop}, {row_stop}).setPropertyValue('{name0}', {f_string})")
        else:
            raise ValueError(f'Only single nested properties can be edited. Note that complex properties can always be edited through other simple nested properties.')

    def _get_cells_property(self, *args, **kwargs):
        """
        cannot get values from here:

        UserDefinedAttributes
        NumberingRules
        Validation
        ValidationLocal
        ValidationXML
        ConditionalFormat
        ConditionalFormatLocal
        ConditionalFormatXML
        """
        column_start = 0
        row_start    = 0

        if len(args) == 2:
            try:
                column_start, row_start = cell2num(args[0])
                name = args[1]
            except ValueError:
                column_start, row_start, column_stop, row_stop = range2num(args[0])
                name = args[1]

        elif len(args) == 3:
            try:
                column_start, row_start = cell2num(args[0])
                column_stop, row_stop   = cell2num(args[1])
                name = args[2]
            except AttributeError:
                column_start = args[0]
                row_start    = args[1]
                name = args[2]
        # elif len(args) == 4:
        #     column_start = args[0]
        #     row_start    = args[1]
        #     column_stop  = args[2]
        #     row_stop     = args[3]
        elif len(args) == 5:
            column_start = args[0]
            row_start    = args[1]
            column_stop  = args[2]
            row_stop     = args[3]
            name = args[4]


        if 'column_start' in kwargs:
            column_start = kwargs['column_start']
        if 'row_start' in kwargs:
            row_start = kwargs['row_start']
        if 'column_stop' in kwargs:
            column_stop = kwargs['column_stop']
        if 'row_stop' in kwargs:
            row_stop = kwargs['row_stop']
        if 'cell_start' in kwargs:
            column_start, row_start = cell2num(kwargs['cell_start'])
        if 'cell_stop' in kwargs:
            column_stop, row_stop = cell2num(kwargs['cell_stop'])
        if 'range' in kwargs:
            column_start, row_start, column_stop, row_stop = range2num(kwargs['range'])
        if 'name' in kwargs:
            name = kwargs['name']

        row_start = _check_row_value(row_start)
        column_start = _check_column_value(column_start)
        try:
            row_stop = _check_row_value(row_stop)
        except NameError:
            row_stop     = self.get_last_row()
        try:
            column_stop = _check_column_value(column_stop)
        except NameError:
            column_stop  = self.get_last_column()

        if column_stop < column_start:
            raise ValueError(f'column_start ({column_start}) cannot be bigger than column_stop ({column_stop})')
        if row_stop < row_start:
            raise ValueError(f'row_start ({row_start}) cannot be bigger than row_stop ({row_stop})')

        t = self.write(f"print(type(get_cells_property_recursively({self.tag}, {column_start}, {row_start}, {column_stop}, {row_stop}, '{name}')))")[0]
        try:
            output = self.write(f"print(get_cells_property_recursively({self.tag}, {column_start}, {row_start}, {column_stop}, {row_stop}, '{name}'))")[0]
        except IndexError:
            return None

        if  t == "<class 'int'>" :
            return int(output)
        elif t == "<class 'bool'>" :
            return str2bool(output)
        elif t == "<class 'dict'>":
            return eval(output)
        else:
            return output
