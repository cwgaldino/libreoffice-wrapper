#! /usr/bin/env python3
# -*- coding: utf-8 -*-
"""Fit.

TODO:
-method to plot parameter for various temperatures
-add fit button to libreoffice
-fit various data (like first and second der)


New atributes added to sheet object:
parameters
var_string
model_string
p_min
p_max
p_guess
p_fitted
p_error
linked_parameters
id_list
submodel
residue
p_cov

new methods:
get_parameters
update_model
update_submodels
fit

functions:
fake_sigma
"""

# standard libraries
import copy
import numpy as np
from pathlib import Path
import inspect
import importlib

# matplotlib
import matplotlib.pyplot as plt
import matplotlib.cm as cm

# fit
from scipy.optimize import curve_fit
from scipy.integrate import trapz

# model functions
from model_functions import *

import sys
sys.path.append('../../')
import libreoffice_wrapper as lw

plt.ion()

def index(x, value):
    """Returns the index of the element in array which is closest to value.

    Args:
        x (list or array): 1D array.
        value (float or int): value.

    Returns:
        index (int)
    """
    return np.argmin(np.abs(np.array(x)-value))

# %%

import importlib
importlib.reload(lw)

sys.path.append('/home/galdino/github/py-backpack')
import backpack.figmanip as figm
import backpack.filemanip as fm
import backpack.arraymanip as am
import time

pid = lw.start_soffice()
time.sleep(5)
soffice = lw.soffice()
calc = soffice.Calc('fit.ods')
calc.save()



# %%
def load_data(self, x, y, sigma=None):
    if len(x) == len(y):
        self.set_column(column='R', value=x, row_start='3')
        self.set_column(column='S', value=y, row_start='3')
    else:
        raise ValueError('x and y must have the same length')
    if sigma is not None:
        if len(x) == len(sigma):
            self.set_column(column='T', value=sigma, row_start='3')
        else:
            raise ValueError('x and sigma must have the same length')

def refresh():

    importlib.reload(sys.modules[__name__])

# %%
importlib.reload(lw)
pid = lw.start_soffice()
time.sleep(5)
soffice = lw.soffice()

calc = soffice.Calc('fit.ods')
sheet = calc.get_sheet_by_position(1)
sheet.load_data = load_data
data = fm.load_data('example/data/00_Co3O2BO3_15K.dat')
# x, y = am.extract(data['rshift'], data['intensity'], [0, 50])
x, y = data['rshift'], data['intensity']
x[1]
sheet.load_data(sheet, x, y)

# x = data['rshift']
# x = np.array([1,2,3])
type(x)
sheet.set_column(column='R', value=x[0:373], row_start='3')
sheet.set_column(column='R', value=x, row_start='3')
sheet.set_row(row='13', value=x[0:500])



x2 = lw.transpose(x[0:374])
sheet.set_value(value=x2, column_start=17, column_stop=17, row_start=2, row_stop=2 + len(x2)- 1, format='formula')




def chunk(lst, n):
    for i in range(0, len(lst), n):
        yield lst[i:i + n]


value = [x[0:500]]

# fix rows
i = 1
string_length = len(str(value[0]))
while string_length > 1000:
    i += 1
    g = chunk(value, int(len(value)/i))
    value2 = next(g)
    string_length = len(str(value2))

sheet.write(f"""value = {value2}""")
while True:
    try:
        sheet.write(f"""value.extend({next(g)})""")
    except StopIteration:
        break
sheet.write("""sheet_5.getCellRangeByPosition(0, 13, 2+len(value)-1, 13).setFormulaArray(value)""")




value = lw.transpose(x[0:500])

# fix columns
i = 1
string_length = len(str(value))
while string_length > 1000:
    i += 1
    g = chunk(value, int(len(value)/i))
    value2 = next(g)
    string_length = len(str(value2))
    if i > len(value):
        raise ValueError('Length of rows are too big. Try setting less columns at a time.')

sheet.write(f"""value = {value2}""")
while True:
    try:
        sheet.write(f"""value.extend({next(g)})""")
    except StopIteration:
        break

sheet.write("""sheet_5.getCellRangeByPosition(17, 2, 17, 2+len(value)-1).setFormulaArray(value)""")





# %%

def chunk(lst, n):
    for i in range(0, len(lst), n):
        yield lst[i:i + n]

def partitionate(value, max_string=4000, i=1):
    # find i
    while len(str(value)) > max_string:
        i += 1

        if i > len(value):
            raise ValueError(f'It seemns like length of rows are too big ({i}). Try setting less columns at a time and/or shorter data (less characters)')

        g = chunk(value, int(len(value)/i))
        while True:
            try:
                if len(str(next(g))) > max_string:
                    return partitionate(value, max_string=max_string, i=i)
            except StopIteration:
                break

        return chunk(value, int(len(value)/i))

    return chunk(value, int(len(value)/i))
# %%

max_string = 25
value = lw.transpose(x[0:500])
value = lw.transpose(x[0:6])
print(value)
len(str(value))
a=chunk(value, 2)
next(a)
len(str(next(a)))
max_string = 2000
a=partitionate(value, max_string)
print(value)
next(a)

len(a)
len(str(a[0]))



    sheet.write(f"""value = {value2}""")
    while True:
        try:
            sheet.write(f"""value.extend({next(g)})""")
        except StopIteration:
            break





# fix columns
i = 1
flag = True
string_length = len(str(value))
# while string_length > 1000:


def partitionate(value, flag=True):
    final = []
    if len(str(value)) > 40:
        a = partitionate(value[:int(len(value)/2)])
        b = partitionate(value[int(len(value)/2):])
        return [x for x in a]
        # return [partitionate(value[:int(len(value)/2)]), partitionate(value[int(len(value)/2):])]
    else:
        return value


value = lw.transpose(x[0:12])
a=partitionate(value)
len(a)
len(str(a[0]))

len(str(value[:int(len(value)/2)]))

    if i > len(value):
        raise ValueError('Length of rows are too big. Try setting less columns at a time.')

while True:
    try:
    except StopIteration:
        break

sheet.write("""sheet_5.getCellRangeByPosition(17, 2, 17, 2+len(value)-1).setFormulaArray(value)""")











sheet.set_value(value=x2, column_start=17, column_stop=17, row_start=2, row_stop=2 + len(x2)- 1, format='formula')

# %%
# soffice.kill()

exec("""sheet_3.getCellRangeByPosition(17, 2, 17, 370).setFormulaArray([[13.5527], [13.8821], [14.2095]
, [14.5367], [14.864], [15.1913], [15.5186], [15.8458], [16.173], [16.5003], [16.8275], [17.1547], [17.4819], [17.809], [18.1362], [18.4634], [18.7905], [19.1176], [19.4448], [19.7719], [20.1011], [20.4282], [20.7553], [21.0823], [21.4094], [21.7364], [22.0635], [22.3905], [22.7175], [23.0445], [23.3715], [23.6985], [24.0254], [24.3524], [24.6793], [25.0062], [25.3331], [25.6601], [25.9869], [26.3138], [26.6407], [26.9676], [27.2944], [27.6212], [27.9481], [28.2727], [28.5995], [28.9263], [29.2531], [29.5799], [29.9066], [30.2334], [30.5601], [30.8868], [31.2135], [31.5402], [31.8669], [32.1936], [32.5203], [32.8469], [33.1736], [33.5002], [33.8268], [34.1513], [34.4779], [34.8045], [35.1311], [35.4576], [35.7842], [36.1107], [36.4372], [36.7638], [37.0903], [37.4168], [37.7432], [38.0676], [38.394], [38.7205], [39.0469], [39.3733], [39.6998], [40.0262], [40.3526], [40.6768], [41.0032], [41.3295], [41.6559], [41.9822], [42.3085], [42.6348], [42.9611], [43.2853], [43.6116], [43.9378], [44.2641], [44.5903], [44.9166], [45.2428], [45.5669], [45.8931], [46.2193], [46.5454], [46.8716], [47.1978], [47.5218], [47.8479], [48.174], [48.5001], [48.8262], [49.1502], [49.4762], [49.8023], [50.1284], [50.4544], [50.7783], [51.1043], [51.4303], [51.7563], [52.0823], [52.4061], [52.7321], [53.058], [53.384], [53.7099], [54.0337], [54.3596], [54.6855], [55.0114], [55.3351], [55.6609], [55.9868], [56.3126], [56.6363], [56.9621], [57.2879], [57.6137], [57.9374], [58.2631], [58.5889], [58.9146], [59.2382], [59.564], [59.8897], [60.2154], [60.5389], [60.8646], [61.1903], [61.5159], [61.8394], [62.1651], [62.4907], [62.8142], [63.1398], [63.4654], [63.7909], [64.1144], [64.4399], [64.7655], [65.0889], [65.4144], [65.7399], [66.0632], [66.3887], [66.7142], [67.0375], [67.363], [67.6884], [68.0117], [68.3371], [68.6625], [68.9858], [69.3112], [69.6366], [69.9598], [70.2851], [70.6105], [70.9337], [71.259], [71.5843], [71.9074], [72.2327], [72.558], [72.8811], [73.2064], [73.5295], [73.8547], [74.1799], [74.503], [74.8282], [75.1534], [75.4764], [75.8016], [76.1246], [76.4497], [76.7748], [77.0978], [77.4229], [77.7459], [78.071], [78.396], [78.7189], [79.044], [79.3669], [79.6919], [80.0169], [80.3398], [80.6648], [80.9876], [81.3126], [81.6354], [81.9604], [82.2853], [82.6081], [82.933], [83.2558], [83.5807], [83.9034], [84.2283], [84.551], [84.8758], [85.2007], [85.5234], [85.8482], [86.1708], [86.4956], [86.8183], [87.143], [87.4657], [87.7904], [88.113], [88.4377], [88.7603], [89.085], [89.4075], [89.7322], [90.0547], [90.3794], [90.7019], [91.0265], [91.349], [91.6736], [91.9961], [92.3207], [92.6431], [92.9677], [93.2901], [93.6147], [93.937], [94.2616], [94.5839], [94.9084], [95.2308], [95.5553], [95.8776], [96.202], [96.5244], [96.8466], [97.1711], [97.4933], [97.8177], [98.14], [98.4644], [98.7866], [99.1109], [99.4332], [99.7553], [100.08], [100.402], [100.726], [101.048], [101.373], [101.695], [102.017], [102.341], [102.663], [102.987], [103.309], [103.631], [103.956], [104.278], [104.602], [104.924], [105.246], [105.57], [105.892], [106.214], [106.538], [106.86], [107.184], [107.506], [107.828], [108.152], [108.474], [108.796], [109.12], [109.442], [109.766], [110.087], [110.409], [110.733], [111.055], [111.377], [111.701], [112.023], [112.344], [112.668], [112.99], [113.312], [113.635], [113.957], [114.279], [114.603], [114.924], [115.246], [115.568], [115.891], [116.213], [116.535], [116.858], [117.18], [117.502], [117.825], [118.147], [118.468], [118.79], [119.114], [119.435], [119.757], [120.08], [120.402], [120.723], [121.045], [121.368], [121.69], [122.011], [122.333], [122.656], [122.977], [123.299], [123.62], [123.944], [124.265], [124.586], [124.908], [125.231], [125.553], [125.874], [126.195], [126.518], [126.84], [127.161], [127.482], [127.803], [128.127], [128.448], [128.769], [129.09], [129.412], [129.735], [130.056], [130.377], [130.698], [131.019], [131.343], [131.664], [131.985], [132.306], [132.627], [132.95], [132.95], [132.95], [132.95], [132.95], [132.95], [132.95], [132.95], [132.95], [132.95], [132.95]])""")



exec("""a=[[13.5527], [13.8821], [14.2095]
, [14.5367], [14.864], [15.1913], [15.5186], [15.8458], [16.173], [16.5003], [16.8275], [17.1547], [17.4819], [17.809], [18.1362], [18.4634], [18.7905], [19.1176], [19.4448], [19.7719], [20.1011], [20.4282], [20.7553], [21.0823], [21.4094], [21.7364], [22.0635], [22.3905], [22.7175], [23.0445], [23.3715], [23.6985], [24.0254], [24.3524], [24.6793], [25.0062], [25.3331], [25.6601], [25.9869], [26.3138], [26.6407], [26.9676], [27.2944], [27.6212], [27.9481], [28.2727], [28.5995], [28.9263], [29.2531], [29.5799], [29.9066], [30.2334], [30.5601], [30.8868], [31.2135], [31.5402], [31.8669], [32.1936], [32.5203], [32.8469], [33.1736], [33.5002], [33.8268], [34.1513], [34.4779], [34.8045], [35.1311], [35.4576], [35.7842], [36.1107], [36.4372], [36.7638], [37.0903], [37.4168], [37.7432], [38.0676], [38.394], [38.7205], [39.0469], [39.3733], [39.6998], [40.0262], [40.3526], [40.6768], [41.0032], [41.3295], [41.6559], [41.9822], [42.3085], [42.6348], [42.9611], [43.2853], [43.6116], [43.9378], [44.2641], [44.5903], [44.9166], [45.2428], [45.5669], [45.8931], [46.2193], [46.5454], [46.8716], [47.1978], [47.5218], [47.8479], [48.174], [48.5001], [48.8262], [49.1502], [49.4762], [49.8023], [50.1284], [50.4544], [50.7783], [51.1043], [51.4303], [51.7563], [52.0823], [52.4061], [52.7321], [53.058], [53.384], [53.7099], [54.0337], [54.3596], [54.6855], [55.0114], [55.3351], [55.6609], [55.9868], [56.3126], [56.6363], [56.9621], [57.2879], [57.6137], [57.9374], [58.2631], [58.5889], [58.9146], [59.2382], [59.564], [59.8897], [60.2154], [60.5389], [60.8646], [61.1903], [61.5159], [61.8394], [62.1651], [62.4907], [62.8142], [63.1398], [63.4654], [63.7909], [64.1144], [64.4399], [64.7655], [65.0889], [65.4144], [65.7399], [66.0632], [66.3887], [66.7142], [67.0375], [67.363], [67.6884], [68.0117], [68.3371], [68.6625], [68.9858], [69.3112], [69.6366], [69.9598], [70.2851], [70.6105], [70.9337], [71.259], [71.5843], [71.9074], [72.2327], [72.558], [72.8811], [73.2064], [73.5295], [73.8547], [74.1799], [74.503], [74.8282], [75.1534], [75.4764], [75.8016], [76.1246], [76.4497], [76.7748], [77.0978], [77.4229], [77.7459], [78.071], [78.396], [78.7189], [79.044], [79.3669], [79.6919], [80.0169], [80.3398], [80.6648], [80.9876], [81.3126], [81.6354], [81.9604], [82.2853], [82.6081], [82.933], [83.2558], [83.5807], [83.9034], [84.2283], [84.551], [84.8758], [85.2007], [85.5234], [85.8482], [86.1708], [86.4956], [86.8183], [87.143], [87.4657], [87.7904], [88.113], [88.4377], [88.7603], [89.085], [89.4075], [89.7322], [90.0547], [90.3794], [90.7019], [91.0265], [91.349], [91.6736], [91.9961], [92.3207], [92.6431], [92.9677], [93.2901], [93.6147], [93.937], [94.2616], [94.5839], [94.9084], [95.2308], [95.5553], [95.8776], [96.202], [96.5244], [96.8466], [97.1711], [97.4933], [97.8177], [98.14], [98.4644], [98.7866], [99.1109], [99.4332], [99.7553], [100.08], [100.402], [100.726], [101.048], [101.373], [101.695], [102.017], [102.341], [102.663], [102.987], [103.309], [103.631], [103.956], [104.278], [104.602], [104.924], [105.246], [105.57], [105.892], [106.214], [106.538], [106.86], [107.184], [107.506], [107.828], [108.152], [108.474], [108.796], [109.12], [109.442], [109.766], [110.087], [110.409], [110.733], [111.055], [111.377], [111.701], [112.023], [112.344], [112.668], [112.99], [113.312], [113.635], [113.957], [114.279], [114.603], [114.924], [115.246], [115.568], [115.891], [116.213], [116.535], [116.858], [117.18], [117.502], [117.825], [118.147], [118.468], [118.79], [119.114], [119.435], [119.757], [120.08], [120.402], [120.723], [121.045], [121.368], [121.69], [122.011], [122.333], [122.656], [122.977], [123.299], [123.62], [123.944], [124.265], [124.586], [124.908], [125.231], [125.553], [125.874], [126.195], [126.518], [126.84], [127.161], [127.482], [127.803], [128.127], [128.448], [128.769], [129.09], [129.412], [129.735], [130.056], [130.377], [130.698], [131.019], [131.343], [131.664], [131.985], [132.306], [132.627], [132.95]]""")


exec("""sheet_3.getCellRangeByPosition(17, 2, 17, 370).setFormulaArray(a)""")



# %%
last_col = 'L'
header_row = 1
hashtag_col = 1

exp_def     = dict(linewidth=2, markersize=8, color='black')
guess_def   = dict(linewidth=2, linestyle='--', color='green')
fit_def     = dict(linewidth=2, color='red')
ties_def    = dict(markersize=8, color='orange')
der_def     = dict(linewidth=0, marker='o', color='black')
der_fit_def = dict(linewidth=2, color='red')
der_guess_def = dict(linewidth=2, linestyle='--', color='green')
sub_def     = dict(linewidth=1)


def get_parameters(self,):
    # self = sheet

    # set # col
    self.set_col_values(np.arange(0, self.get_last_row()-1), col=hashtag_col, row_start=header_row+1)

    # get header and submodels
    header = self.get_row_values(header_row, col_stop=last_col)
    submodel_col = header.index('submodel')+1
    arg_col = header.index('arg')+1
    submodel_list = self.get_col_values(submodel_col)[1:]

    self.parameters = dict()
    for row_number, submodel in enumerate(submodel_list):
        if submodel == '':
            pass
        else:
            row_values = self.get_row_values(row_number+2, col_stop=last_col)
            arg = row_values[arg_col-1]

            try:
                if arg in self.parameters[submodel]: # check for repeated parameter id's
                    for key, value in self.parameters[submodel][arg].items():
                        # if type(value) != list:
                        #     value = [value]
                        value.append(row_values[header.index(key)])
                        self.parameters[submodel][arg][key] = value
                else:
                    self.parameters[submodel][arg]={header[col_number]: [value] for col_number, value in enumerate(row_values)}
            except KeyError:
                self.parameters[submodel] = {arg:{header[col_number]: [value] for col_number, value in enumerate(row_values)}}


def update_model(self):
    refresh()

    var_string = ''
    model_string = ''
    self.p_min = []
    self.p_max = []
    self.p_guess = []
    self.p_fitted = []
    self.p_error = []

    # get parameters
    self.get_parameters()

    # get header and submodels
    header = self.get_row_values(header_row, col_stop=last_col)
    guess_col = header.index('guess')+1
    fitted_col = header.index('fitted')+1
    error_col = header.index('error')+1
    id_col = header.index('id')+1
    self.set_col_values(data=['' for i in range(self.get_last_row()-1)], row_start=header_row+1, col=id_col)

    p = 0
    x = 0
    self.linked_parameters = {}
    for submodel in self.parameters:

        # check if this submodel should be used
        if 'y' in [use for sublist in [self.parameters[submodel][arg]['use'] for arg in self.parameters[submodel]] for use in sublist]:

            # get tag
            try:
                submodel_tag = submodel.split('#')[-1]
            except IndexError:
                submodel_tag = None
            submodel_name = submodel.split('#')[0]

            # get arguments from function
            args_expected = list(inspect.signature(eval(submodel_name)).parameters)

            # initialize model
            model_string += f"{submodel_name}(x, "

            # build min, max, guess, model
            for arg in args_expected[1: ]:
                # print(submodel, arg)

                # check if submodel has active argument
                missing_arg = False
                try:
                    to_use = self.parameters[submodel][arg]['use'].index('y')
                except ValueError:
                    missing_arg = True
                if missing_arg: raise MissingArgument(submodel, arg)

                # check if parameter must vary =========================================================
                vary = list(self.parameters[submodel][arg]['vary'])[to_use]
                hashtag = list(self.parameters[submodel][arg]['#'])[to_use]
                # linked parameter ===================================
                if vary != 'y' and vary != 'n':
                    submodel2link = vary.split(',')[0]
                    arg2link = vary.split(',')[-1]
                    if submodel2link in self.parameters and arg2link in self.parameters[submodel]:
                        to_use_linked = self.parameters[submodel2link][arg2link]['use'].index('y')
                        vary2 = list(self.parameters[submodel2link][arg2link]['vary'])[to_use_linked]
                    else:
                        raise ValueError(f"Cannot find submodel '{submodel2link}' with arg '{arg2link}'.")
                    while vary2 != 'y' and vary2 != 'n':
                        # print(vary2)
                        submodel2link = vary2.split(',')[0]
                        arg2link = vary2.split(',')[-1]
                        # print(submodel2link, arg2link)
                        if submodel2link in self.parameters and arg2link in self.parameters[submodel]:
                            to_use_linked = self.parameters[submodel2link][arg2link]['use'].index('y')
                            vary2 = list(self.parameters[submodel2link][arg2link]['vary'])[to_use_linked]
                        else:
                            raise ValueError(f"Cannot find submodel '{submodel2link}' with arg '{arg2link}'.")

                    if vary2 == 'n': #
                        v = list(self.parameters[submodel2link][arg2link]['guess'])[to_use_linked]
                        model_string += f'{v}, '
                        self.set_cell_value(value='-', row=hashtag+header_row+1, col=id_col)
                        self.set_cell_value(value=v, row=hashtag+header_row+1, col=guess_col)
                        self.set_cell_value(value=v, row=hashtag+header_row+1, col=fitted_col)
                        self.set_cell_value(value=0, row=hashtag+header_row+1, col=error_col)
                    else:
                        if submodel2link+','+arg2link in self.linked_parameters:
                            x_temp = self.linked_parameters[submodel2link+','+arg2link]
                            self.set_cell_value(value='x' + str(x_temp), row=hashtag+header_row+1, col=id_col)
                            self.parameters[submodel][arg]['id'][to_use] = 'x' + str(x_temp)
                            # var_string += f'x{x_temp}, '
                            model_string += f'x{x_temp}, '
                        else:
                            self.linked_parameters[submodel2link+','+arg2link] = x
                            self.set_cell_value(value='x' + str(x), row=hashtag+header_row+1, col=id_col)
                            self.parameters[submodel][arg]['id'][to_use] = 'x' + str(x)
                            # var_string += f'x{x}, '
                            model_string += f'x{x}, '
                            x += 1


                # fixed parameter ================================
                elif vary == 'n':
                    v = list(self.parameters[submodel][arg]['guess'])[to_use]
                    model_string += f'{v}, '
                    self.set_cell_value(value='-', row=hashtag+header_row+1, col=id_col)
                    self.parameters[submodel][arg]['id'][to_use] = '-'
                    self.set_cell_value(value=v, row=hashtag+header_row+1, col=fitted_col)
                    self.parameters[submodel][arg]['fitted'][to_use] = v
                    self.set_cell_value(value=0, row=hashtag+header_row+1, col=error_col)
                    self.parameters[submodel][arg]['error'][to_use] = 0


                # variable parameter =============================
                else:
                    self.p_min.append(list(self.parameters[submodel][arg]['min'])[to_use])
                    self.p_max.append(list(self.parameters[submodel][arg]['max'])[to_use])
                    self.p_guess.append(list(self.parameters[submodel][arg]['guess'])[to_use])
                    self.p_fitted.append(list(self.parameters[submodel][arg]['fitted'])[to_use])
                    self.p_error.append(list(self.parameters[submodel][arg]['error'])[to_use])

                    try:
                        if submodel+','+arg in self.linked_parameters:
                            x_temp = self.linked_parameters[submodel+','+arg]
                            self.set_cell_value(value='x' + str(x_temp), row=hashtag+header_row+1, col=id_col)
                            self.parameters[submodel][arg]['id'][to_use] = 'x' + str(x_temp)
                            var_string += f'x{x_temp}, '
                            model_string += f'x{x_temp}, '
                        else:
                            self.set_cell_value(value='p' + str(p), row=hashtag+header_row+1, col=id_col)
                            self.parameters[submodel][arg]['id'][to_use] = 'p' + str(p)
                            var_string += f'p{p}, '
                            model_string += f'p{p}, '
                            p += 1
                    except UnboundLocalError:
                        var_string += f'p{p}, '
                        model_string += f'p{p}, '
                        p += 1

            model_string += ') + '

    # finish model
    self.id_list = [s.strip() for s in eval('["' + var_string[:-2].replace(',', '","') + '"]')]

    self.model_string = f'lambda x, {var_string[:-2]}: {model_string[:-3]}'
    self.model = eval(self.model_string)

    # check guess, min, max ============================
    if '' in self.p_guess:
        guess_missing = [self.id_list[i] for i, x in enumerate(self.p_guess) if x == '']
        raise ValueError(f'Parameters with id {guess_missing} do not have a guess value.')

    if '' in self.p_min:
        self.p_min = [-np.inf if x == '' else x for x in self.p_min]
    if '' in self.p_max:
        self.p_max = [np.inf if x == '' else x for x in self.p_max]
    if '' in self.p_fitted:
        self.p_fitted = [0 if x == '' else x for x in self.p_fitted]
    if '' in self.p_error:
        self.p_error = [0 if x == '' else x for x in self.p_error]

    # submodel
    self.update_submodels()


def update_submodels(self):

    self.submodel = {}

    for submodel in self.parameters:

        # check if submodel should be used
        if 'y' in [use for sublist in [self.parameters[submodel][arg]['use'] for arg in self.parameters[submodel]] for use in sublist]:
            self.submodel[submodel] = {'guess_string': '', 'fit_string':''}

            # get tag
            try:
                submodel_tag = submodel.split('#')[-1]
            except IndexError:
                submodel_tag = None
            submodel_name = submodel.split('#')[0]

            # get arguments from function
            import __main__
            try:
                args_expected = list(inspect.signature(eval(f'__main__.{submodel_name}')).parameters)
            except AttributeError:
                args_expected = list(inspect.signature(eval(submodel_name)).parameters)

            # initialize submodel
            self.submodel[submodel]['guess_string'] += f'{submodel_name}(x, '
            self.submodel[submodel]['fit_string'] += f'{submodel_name}(x, '


            for arg in args_expected[1: ]:

                # check if submodel has active argument
                missing_arg = False
                try:
                    to_use = self.parameters[submodel][arg]['use'].index('y')
                except ValueError:
                    missing_arg = True
                if missing_arg: raise MissingArgument(submodel, arg)

                # build min, max, guess, model
                if self.parameters[submodel][arg]['id'][to_use] != '-':
                    id = self.parameters[submodel][arg]['id'][to_use]
                    self.submodel[submodel]['guess_string'] += str(self.p_guess[self.id_list.index(id)]) + ', '
                    self.submodel[submodel]['fit_string']   += str(self.p_fitted[self.id_list.index(id)]) + ', '
                else:
                    self.submodel[submodel]['guess_string'] += str(self.parameters[submodel][arg]['guess'][to_use]) + ', '
                    self.submodel[submodel]['fit_string'] += str(self.parameters[submodel][arg]['fitted'][to_use]) + ', '

            self.submodel[submodel]['guess_string'] = self.submodel[submodel]['guess_string'][:-2] + ')'
            self.submodel[submodel]['fit_string'] = self.submodel[submodel]['fit_string'][:-2] + ')'

            self.submodel[submodel]['guess'] = eval(f'lambda x:' + self.submodel[submodel]['guess_string'])
            self.submodel[submodel]['fit'] = eval(f'lambda x:' + self.submodel[submodel]['fit_string'])


def fit(self, x, y, ties=None, global_sigma=1e-13, save=True):

    self.update_model()

    # fit
    if global_sigma is not None:
        sigma = fake_sigma(x, global_sigma=global_sigma, sigma_specific=ties)
    self.p_fitted, self.p_cov = curve_fit(self.model, x, y, self.p_guess, sigma=sigma, bounds=[self.p_min, self.p_max])
    self.p_error = np.sqrt(np.diag(self.p_cov))  # One standard deviation errors on the parameters

    # get residue
    self.residue = trapz(abs(y - self.model(x, *self.p_fitted)), x)

    # save to sheet and self.parameter =====================================================
    # get header and submodels
    header = self.get_row_values(header_row, col_stop=last_col)
    fitted_col = header.index('fitted')+1
    error_col = header.index('error')+1

    for submodel in self.parameters:
        # check if submodel should be used
        if 'y' in [use for sublist in [self.parameters[submodel][arg]['use'] for arg in self.parameters[submodel]] for use in sublist]:

            # get tag
            try:
                submodel_tag = submodel.split('#')[-1]
            except IndexError:
                submodel_tag = None
            submodel_name = submodel.split('#')[0]

            # get arguments from function
            args_expected = list(inspect.signature(eval(submodel_name)).parameters)

            # build min, max, guess, model
            for arg in args_expected[1: ]:

                # check if submodel has active argument
                missing_arg = False
                try:
                    to_use = self.parameters[submodel][arg]['use'].index('y')
                except ValueError:
                    missing_arg = True
                if missing_arg: raise MissingArgument(submodel, arg)

                hashtag = list(self.parameters[submodel][arg]['#'])[to_use]
                if self.parameters[submodel][arg]['id'][to_use] != '-':
                    id = self.parameters[submodel][arg]['id'][to_use]
                    v1 = self.p_fitted[self.id_list.index(id)]
                    v2 = self.p_error[self.id_list.index(id)]
                    self.set_cell_value(value=v1, row=hashtag+header_row+1, col=fitted_col)
                    self.set_cell_value(value=v2, row=hashtag+header_row+1, col=error_col)
                    self.parameters[submodel][arg]['fitted'][to_use] = v1
                    self.parameters[submodel][arg]['error'][to_use] = v2

    self.update_submodels()

    if save:
        self.calc.save()


def fake_sigma(x, global_sigma=10**-10, sigma_specific=None):
    """Build a fake sigma array which determines the uncertainty in ydata.

    Adaptaded from the `scipy.optimize.curve_fit() <https://docs.scipy.org/doc/scipy/reference/generated/scipy.optimize.curve_fit.html>`_ documentation:

        If we define residuals as ``r = ydata - model(xdata, *popt)``, then sigma
        for a 1-D data should contain values of standard deviations of errors in
        ydata. In this case, the optimized function is ``chisq = sum((r / sigma) ** 2)``.

    Args:
        x (list): x array.
        sigma (float, optional): sigma value to be used for all points in ``x``.
        sigma_specific (list, optional): list of triples specfing new sigma for specific ranges, e.g.,
            ``sigma_specific = [[x_init, x_final, sigma], [x_init2, x_final2, sigma2], ]``.

    Returns:
        array.
    """
    p_sigma = np.ones(len(x))*global_sigma

    if sigma_specific is not None:
        for sigma in sigma_specific:
            init = index(x, sigma[0])
            final = index(x, sigma[1])
            p_sigma[init:final] = global_sigma/sigma[2]

    return p_sigma





def plot_fit(self, x, y, ax=None, show_exp=True, show_derivative=False, show_submodels=False, smoothing=10, ties=None, submodels_bkg=None, derivative_order=1, derivative_offset=None, derivative_factor=None, derivative_window_size=1):

    if smoothing == 0:
        x_smooth = x
    elif smoothing < 1:
        smoothing = 1
        x_smooth = np.linspace(min(x), max(x), len(x)*smoothing)
    else:
        x_smooth = np.linspace(min(x), max(x), len(x)*smoothing)

    if ax is None:
        fig = figm.figure()
        ax = fig.add_subplot(111)

    # exp
    if show_exp:
        ax.plot(x, y, **exp_def, label='exp')

    # ties
    if ties is not None:
        for pair in ties:
            ax.plot(x[am.index(x, pair[0]):am.index(x, pair[1])], y[am.index(x, pair[0]):am.index(x, pair[1])], **ties_def)

    # derivative
    if show_derivative:
        x_der, y_der = am.derivative(x, y, order=derivative_order, window_size=derivative_window_size)

        if derivative_factor is None:
            derivative_factor = (max(y) - np.mean(y))/(max(y_der) - np.mean(y_der))
        if derivative_offset is None:
            derivative_offset = -(max(y_der*derivative_factor) - np.mean(y))

        ax.plot(x_der, y_der*derivative_factor+derivative_offset, **der_def)

    # fit
    self.update_model()
    y_fit = self.model(x_smooth, *self.p_fitted)
    ax.plot(x_smooth, y_fit, **fit_def)

    if show_derivative:
        x_fir_der, y_fit_der = am.derivative(x_smooth, y_fit, order=derivative_order)
        ax.plot(x_fir_der, y_fit_der*derivative_factor+derivative_offset, **der_fit_def)

    # submodels
    if show_submodels:
        sub_colors = cm.get_cmap('Set1').colors
        if submodels_bkg is not None:
            bkg = self.submodel[submodels_bkg]['fit'](x_smooth)
        else:
            bkg=0
        for submodel, color in zip(self.submodel, sub_colors):
            if submodel != submodels_bkg:
                ax.plot(x_smooth, self.submodel[submodel]['fit'](x_smooth) + bkg, **sub_def, color=color, label=submodel)

    plt.legend()
    return ax


def plot_guess(self, x, y, ax=None, show_exp=True, show_derivative=False, show_submodels=False, smoothing=10, ties=None, submodels_bkg=None, derivative_order=1, derivative_offset=None, derivative_factor=None, derivative_window_size=1):

    if smoothing == 0:
        x_smooth = x
    elif smoothing < 1:
        smoothing = 1
        x_smooth = np.linspace(min(x), max(x), len(x)*smoothing)
    else:
        x_smooth = np.linspace(min(x), max(x), len(x)*smoothing)

    if ax is None:
        fig = figm.figure()
        ax = fig.add_subplot(111)

    # exp
    if show_exp:
        ax.plot(x, y, **exp_def, label='exp')

    # ties
    if ties is not None:
        for pair in ties:
            ax.plot(x[am.index(x, pair[0]):am.index(x, pair[1])], y[am.index(x, pair[0]):am.index(x, pair[1])], **ties_def)

    # derivative
    if show_derivative:
        x_der, y_der = am.derivative(x, y, order=derivative_order, window_size=derivative_window_size)

        if derivative_factor is None:
            derivative_factor = (max(y) - np.mean(y))/(max(y_der) - np.mean(y_der))
        if derivative_offset is None:
            derivative_offset = -(max(y_der*derivative_factor) - np.mean(y))

        ax.plot(x_der, y_der*derivative_factor+derivative_offset, **der_def)

    # fit
    y_guess = self.model(x_smooth, *self.p_guess)
    ax.plot(x_smooth, y_guess, **guess_def)

    if show_derivative:
        x_fir_der, y_fit_der = am.derivative(x_smooth, y_guess, order=derivative_order)
        ax.plot(x_fir_der, y_fit_der*derivative_factor+derivative_offset, **der_guess_def)

    # submodels
    if show_submodels:
        sub_colors = cm.get_cmap('Set1').colors
        if submodels_bkg is not None:
            bkg = self.submodel[submodels_bkg]['guess'](x_smooth)
        else:
            bkg=0
        for submodel, color in zip(self.submodel, sub_colors):
            if submodel != submodels_bkg:
                ax.plot(x_smooth, self.submodel[submodel]['guess'](x_smooth) + bkg, **sub_def, color=color, label=submodel)
    plt.legend()
    return ax


class MissingArgument(Exception):

    # Constructor or Initializer
    def __init__(self, submodel, arg):
        self.submodel = submodel
        self.arg = arg

    # __str__ is to print() the value
    def __str__(self):
        msg = f"Submodel '{self.submodel}' is missing argument '{self.arg}'."
        return(msg)

sheet.get_parameters = get_parameters
sheet.update_model = update_model
sheet.update_submodels = update_submodels
sheet.fit = fit
sheet.plot_fit = plot_fit
sheet.plot_guess = plot_guess


try:
    from __main__ import *
except ImportError:
    pass
