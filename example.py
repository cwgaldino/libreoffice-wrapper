#! /usr/bin/env python3
# -*- coding: utf-8 -*-

import libreoffice_wrapper as lw
import importlib
importlib.reload(lw)

# start LibreOffice and establish communication
pid = lw.start_soffice()
soffice = lw.soffice()

# Open Calc
calc = soffice.Calc()  # it will try connect with any open Calc instance. If nothing is open, it will start a new spreadsheet
# calc = soffice.Calc('<path-to-spreadsheet-file>')  # connects/opens specific file
# calc = soffice.Calc(force_new=True)  # open a new file

# Calc info
print(calc.get_title())
print(calc.get_filepath())
print(calc.get_sheets_count())
print(calc.get_sheets_name())

# save
calc.save()
# calc.save('<path-to-save>')

# close Calc
# calc.close()

# insert new sheet
calc.insert_sheet('my_new_sheet')
calc.insert_sheet('sheet_to_be_remove')
calc.insert_sheet('another_sheet_to_be_remove')

# remove sheet
calc.remove_sheets_by_position(3)
calc.remove_sheet('sheet_to_be_remove')

# move sheet
calc.move_sheet(name='my_new_sheet', position=0)

# copy_sheet
calc.copy_sheet(name='my_new_sheet', new_name='copied_sheet', position=2)

# sheet name and position
print(calc.get_sheet_position(name='my_new_sheet'))
print(calc.get_sheet_name_by_position(position=0))

# Styles
print(calc.get_styles())
properties = {'CellBackColor':16776960, 'CharWeight':150}
calc.new_style(name='my_new_style', properties=properties)
calc.remove_style(name='my_new_style')

# get sheet
sheet = calc.get_sheet_by_position(0)
sheet = calc.get_sheet('my_new_sheet')

# sheet name
print(sheet.get_name())
sheet.set_name('new_name')

# visibility
print(sheet.isVisible())

# move
sheet.move(position=2)  # in this case moving to 0 or 1 yields the same result

# remove (delete)
# sheet.remove()

# set/get data (data can be set in many ways)
sheet.set_value('A1', 'hello')
print(sheet.get_value('A1'))

sheet.set_value('B', '1', 'hello 2')
print(sheet.get_value('B', '1'))

sheet.set_value('C', 0, 'hello 3')
print(sheet.get_value('C', 0))

sheet.set_value(3, 0, 'hello 4')
print(sheet.get_value(3, 0))

sheet.set_value(4, '1', 'hello 5')
print(sheet.get_value(4, '1'))

sheet.set_value('A2:C3', [['a', 'b', 'c'], [1, 2, 3]])
print(sheet.get_value('A2:C3'))

sheet.set_value('A4', 'C5', [['a', 'b', 'c'], [1, 2, 3]])
print(sheet.get_value('A4', 'C5'))

sheet.set_value('A6', [['a', 'b', 'c'], [1, 2, 3]])
print(sheet.get_value('A6:C7'))

sheet.set_value('A', '8', [['a', 'b', 'c'], [1, 2, 3]])
print(sheet.get_value('A8:C9'))

sheet.set_value('A', 9, [['a', 'b', 'c'], [1, 2, 3]])
print(sheet.get_value(0, 9, 2, 11))

sheet.set_value('A', '12', 'C', '13', [['a', 'b', 'c'], [1, 2, 3]])
print(sheet.get_value('A', '12', 'C', '13'))

sheet.set_value(0, 13, 2, 14, [['a', 'b', 'c'], [1, 2, 3]])
print(sheet.get_value(0, 13, 2, 14))





# FROM HERE===============
importlib.reload(lw)
calc = soffice.Calc()
sheet = calc.get_sheet_by_position(0)

sheet.set_row(row=1, value=[1, 2, 3, 4, 5, 6])
print(sheet.get_row(row=1))

sheet.set_row(row=0, value=[1, 2, 3, 4, 5, 6])
print(sheet.get_row(row=0))


sheet.set_column(column, value=[1, 2, 3, 4, 5, 6])
sheet.get_column()

# last used row/column
print(sheet.get_last_row())
print(sheet.get_last_column())


# set/get row/column data



# length of row/column
sheet.get_row_length(row)
sheet.get_column_length(column)

sheet.set_column_width(column, width)
sheet.get_column_width(column)
sheet.set_row_height(row, height)
sheet.get_row_height(row)




sheet.merge()
sheet.unmerge()
sheet.remove_conditional_format()
sheet.get_conditional_formats()
sheet.new_conditional_format()




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
