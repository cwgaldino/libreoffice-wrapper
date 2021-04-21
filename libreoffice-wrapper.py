#! /usr/bin/env python3
# -*- coding: utf-8 -*-
"""
tmux
tmux ls
tmux attach -t 0
tmux attach-session -t 0
tmux rename-session -t 0 new_name
tmux new -s session_name
tmux kill-session -t name

panes
ctrl+b %                create panes vertical
ctrl+b <arrow-keys>     move to panes
ctrl+b "               create pane horizontal

windows:
ctrl+b c               new window
ctrl+b 0               move to window 0
ctrl+b 1               move to window 1
ctrl+b ,               rename window
exit                   close window

sessions:


ctrl+b d detach


windows:
...

"""
# %%
import libtmux
import numpy as np
from pathlib import Path
import time
import psutil
import os
import signal
import re
import subprocess
import warnings
import sys

def start_soffice(port=8100, folder='/opt/libreoffice7.0', nodefault=True, norestore=False, nologo=False, tmux_config=True, timeout=10):

    # set libreoffice folder
    folder = Path(folder)/'program'

    # tmux server
    if tmux_config:
        tmux_server = libtmux.Server(config_file='tmux.conf')
    else:
        tmux_server = libtmux.Server()

    # new session
    if tmux_server.has_session('libreoffice-wrapper'):
        tmux_session = tmux_server.find_where({'session_name':'libreoffice-wrapper'})
    else:
        tmux_session = tmux_server.new_session('libreoffice-wrapper')

    # new pane
    if tmux_session.find_where({'window_name':'soffice'}) is None:
        tmux_pane = tmux_session.new_window(attach=False, window_name='soffice').panes[0]
    else:
        tmux_pane = tmux_session.find_where({'window_name':'soffice'}).panes[0]
    time.sleep(1)
    tmux_pane.capture_pane = lambda: [x.rstrip() for x in tmux_pane.cmd('capture-pane', '-p', '-J').stdout]

    # initialize libreoffice
    print('Initialing soffice (it may take a few moments.)')
    tmux_pane.send_keys(f'cd {folder}', enter=True, suppress_history=True)
    time.sleep(0.1)

    options = ''
    if nodefault:
        options += ' --nodefault'
    if norestore:
        options += ' --norestore'
    if nologo:
        options += ' --nologo'

    tmux_pane.send_keys(f"./soffice{options} --accept='socket,host=localhost,port={port};urp;' &", enter=True, suppress_history=True)
    time.sleep(0.1)

    pid = 0
    start_time = time.time()
    while time.time() < start_time + timeout:
        try:
            pid = int(tmux_pane.capture_pane()[-2].split()[-1])
            if pid != 0:
                print(f'soffice started\nProcess pid {pid}')
                if not _has_pid(pid):
                    warnings.warn(f'\nsoffice seems to be already running. \nCommunication stabilished. \nKilling process {pid} will not close soffice', stacklevel=2)
                return pid
        except ValueError:
            pass
    raise TimeoutError(f'soffice taking more than {timeout_soffice} seconds to load. Maybe try again.')

def _has_pid(pid):
    for proc in psutil.process_iter():
        if proc.as_dict(attrs=['pid'])['pid'] == pid: return True
    return False

def _get_children(pid):
    output = subprocess.check_output(["bash", "-c", f"pstree -p -n  {pid}"])
    pattern = re.compile(r"\((\d+)\)")
    pid_children = pattern.findall(output.decode('utf-8'))
    return [int(pid) for pid in pid_children]

def kill(*args, recursive=True):
    """Kill processes."""
    try:
        if recursive:
            pid2 = []
            for pid in args:
                for p in _get_children(pid):
                    pid2.append(p)
        else:
            pid2 = args
    except TypeError:
        if recursive:
            pid2 = []
            for pid in args[0]:
                for p in _get_children(pid):
                    pid2.append(p)
        else:
            pid2 = args

    for pid in pid2:
        try:
            os.kill(pid, signal.SIGKILL)
        except ProcessLookupError:
            pass

def kill_tmux(session_name='libreoffice-wrapper'):
    tmux_server = libtmux.Server()
    if tmux_server.has_session('libreoffice-wrapper'):
        tmux_server.kill_server()

def str2bool(string):
    if string == 'False':
        return False
    else:
        return True

def _name2ImplementationName(name):
    if name == 'scalc' or name == 'calc':
        return 'ScModelObj'
    elif name == 'swriter' or name == 'writer':
        return 'SwXTextDocument'
    elif name == 'simpress' or name == 'impress' or name == 'draw':
        return 'SdXImpressDocument'
    elif name == 'smath' or name == 'math':
        return 'com.sun.star.comp.Math.FormulaDocument'
    elif name == 'base' or name == 'sbase' or name == 'obase':
        return 'com.sun.star.comp.dba.ODatabaseDocument'
    else:
        raise ValueError('type must be calc, writer, impress, draw, math, or base')

def query_yes_no(question, default="yes"):
    """Ask a yes/no question and return answer.

    Note:
        It accepts many variations of yes and no as answer, like, "y", "YES", "N", ...

    Args:
        question (str): string that is presented to the user.
        default ('yes', 'no' or None): default answer if the user just hits
            <Enter>. If None, an answer is required of the user.

    Returns:
        True for "yes" or False for "no".
    """
    valid = {"yes": True, "y": True, "ye": True, "Y": True, "YES": True, "YE": True,
             "no": False, "n": False, "No":True, "NO":True, "N":True}
    if default is None:
        prompt = " [y/n] "
    elif default == "yes":
        prompt = " [Y/n] "
    elif default == "no":
        prompt = " [y/N] "
    else:
        raise ValueError("invalid default answer: '%s'" % default)

    while True:
        sys.stdout.write(question + prompt + '\n')
        choice = input().lower()
        if default is not None and choice == '':
            return valid[default]
        elif choice in valid:
            return valid[choice]
        else:
            sys.stdout.write("Please respond with 'yes' or 'no' "
                             "('y' or 'n').\n")

def _letter2num(string):

    string = string.lower()
    alphabet = 'abcdefghijklmnopqrstuvwxyz'

    n = 0
    for idx, s in enumerate(string):
        n += alphabet.index(s)+(idx)*26

    return n

def cell2num(string):
    if ':' in string:
        raise ValueError('cell seems to be a range.')
    temp = re.compile("([a-zA-Z]+)([0-9]+)")
    res = temp.match(string).groups()
    return _letter2num(res[0]), int(res[1])-1

def range2num(string):
    if ':' not in string:
        raise ValueError('range seems to be a cell.')

    return flatten((cell2num(string.split(':')[0]), cell2num(string.split(':')[1])))

def _check_row_value(row):
    if type(row) == str:
        row = int(row)-1

    if row < 0:
        raise ValueError('row cannot be negative')

    return row

def _check_column_value(column):
    if type(column) == str:
        column = _letter2num(column)

    if column < 0:
        raise ValueError('column cannot be negative')

    return column

def flatten(x):
    """Returns the flattened list or tuple."""
    if len(x) == 0:
        return x
    if isinstance(x[0], list) or isinstance(x[0], tuple):
        return flatten(x[0]) + flatten(x[1:])
    return x[:1] + flatten(x[1:])

def transpose(l):
    """Transpose lists."""
    try:
        row_count, col_count = np.shape(l)
        return [list(x) for x in list(zip(*l))]
    except ValueError:
        return [[x] for x in l]

class soffice():

    def __init__(self, tmux_config=True, port=8100, folder='/opt/libreoffice7.0'):
        """
        server -> session (pycalc) -> window (soffice-python) -> pane (0)
        """
        timeout_python = 10
        timeout_soffice = 60

        # set libreoffice folder
        self.folder = Path(folder)/'program'

        # tmux server
        if tmux_config:
            self.tmux_server = libtmux.Server(config_file='tmux.conf')
        else:
            self.tmux_server = libtmux.Server()

        # new session
        if self.tmux_server.has_session('libreoffice-wrapper'):
            self.tmux_session = self.tmux_server.find_where({'session_name':'libreoffice-wrapper'})
        else:
            self.tmux_session = self.tmux_server.new_session('libreoffice-wrapper')

        # new pane
        if self.tmux_session.find_where({'window_name':'python'}) is None:
            self.tmux_pane = self.tmux_session.new_window(attach=False, window_name='python').panes[0]
        else:
            self.tmux_pane = self.tmux_session.find_where({'window_name':'python'}).panes[0]
        time.sleep(1)

        self.tmux_pane.capture_pane = lambda starting_line=0: [x.rstrip() for x in self.tmux_pane.cmd('capture-pane', '-p', '-J', f'-S {starting_line}').stdout]

        # initialize python
        print('Initialing python (it may take a few moments.)')
        self.tmux_pane.send_keys(f'cd {self.folder}', enter=True, suppress_history=True)
        time.sleep(0.1)
        self.tmux_pane.send_keys('./python', enter=True, suppress_history=True)
        start_time = time.time()
        flag = True
        while time.time() < start_time + timeout_python:
            if self.tmux_pane.capture_pane()[-1] == '>>>':
                print('done')
                flag = False
                break
        if flag:
            raise TimeoutError(f'Python taking more than {timeout_python} seconds to load. Maybe try again.')

        # python imports
        print('Importing modules (it may take a few moments.)')
        self.tmux_pane.send_keys(f"""try: exec(\"import traceback\")""" + '\n'\
                          """except Exception as e: print(e)""" + '\n'
                          , enter=True, suppress_history=False)
        time.sleep(0.1)
        self.write("import uno")
        time.sleep(0.1)
        self.write("from com.sun.star.beans import PropertyValue")
        time.sleep(0.1)
        print('done')

        # initialize communication
        print('Initializing comunications (it may take a few moments.)')
        self.write("local = uno.getComponentContext()")
        time.sleep(0.1)
        self.write("""resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)""")
        time.sleep(0.1)
        self.write(f"""context = resolver.resolve("uno:socket,host=localhost,port={port};urp;StarOffice.ComponentContext")""")
        time.sleep(0.1)
        self.write("desktop = context.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop', context)")
        time.sleep(0.1)
        try:
            self._get_tag()
        except soffice_python_error:
            self._set_tag(0)
        time.sleep(0.1)
        print('done')

    def _get_tag(self):
        return int(self.write("print(tag)")[0])

    def _set_tag(self, tag):
        self.write(f"tag = {tag}")

    def kill(self):
        self.tmux_server.kill_server()

    def check_running(self):
        # if self.tmux_pane.capture_pane()[-1] in ['>>>', '... >>>', '... ... >>>']:
        if self.tmux_pane.capture_pane()[-1].endswith('>>>'):
            return False
        else:
            return True

    def write(self, string, timeout=10, max_output=1000):
        # check quotes

        keys2send = f"try: exec(\"\"\"{string}\"\"\")" + '\n'\
                     """except: print('calc-ERROR'); print(traceback.format_exc())"""
        self.tmux_pane.send_keys(keys2send, enter=True, suppress_history=False)
        self.tmux_pane.enter()
        # time.sleep(0.05)
        time.sleep(0.001)

        start_time = time.time()
        while time.time() < start_time + timeout:
            if self.check_running():
                pass
            else:

                # output = self.tmux_pane.capture_pane()
                # i = output[::-1].index('>>> ' + keys2send.split('\n')[0])

                output = self.tmux_pane.capture_pane()
                starting_line = 0
                while '>>> ' + keys2send.split('\n')[0] not in output and starting_line>-max_output:
                    starting_line -= 10
                    output = self.tmux_pane.capture_pane(starting_line)
                if starting_line<-max_output:
                    raise valueError(f'output seems to be bigger than max_output {max_output}.')
                i = output[::-1].index('>>> ' + keys2send.split('\n')[0])
                if output[-i+2] == 'calc-ERROR':
                    f = output[::-1].index('>>>')
                    raise soffice_python_error('\n'.join(output[-i+3:-1+f]))
                elif output[-i+2].startswith('...'):
                    if output[-i+3] == 'calc-ERROR':
                        f = output[::-1].index('>>>')
                        raise soffice_python_error('\n'.join(output[-i+4:-1+f]))
                    else:
                        return output[-i+3:-1]
                else:
                    return output[-i+2:-1]
        raise TimeoutError(f'timeout: {timeout}s. Process is still running.')

    def read(self, timeout=10, max_output=1000):
        start_time = time.time()
        while time.time() < start_time + timeout:
            if self.check_running():
                pass
            else:
                output = self.tmux_pane.capture_pane()
                starting_line = 0
                while '>>>' not in output[:-1] and starting_line>-max_output:
                    starting_line -= 10
                    output = self.tmux_pane.capture_pane(starting_line)
                if starting_line<-max_output:
                    raise valueError(f'output seems to be bigger than max_output {max_output}.')

                output = output[[i for i, x in enumerate(output[:-1]) if x.startswith('>>>')][-1]+3:-1]
                if len(output) > 0:
                    if output[0] == 'calc-ERROR':
                        raise soffice_python_error('\n'.join(output))
                    else:
                        return output
                else:
                    return output
        raise TimeoutError(f'timeout: {timeout}s. Process is still running.')

    def Calc(self, filepath=None, new_file=False):

        type = 'scalc'
        if new_file:
            tag = self._new_file(type=type)
        elif filepath is None: # if a filepath is not defined
            print('Filepath was not given.')
            if self._has_open():
                print('However, something is open.')
                tag = self._connect_with_open(type=type)
                if not tag:
                    print(f'Opening a new file...')
                    tag = self._new_file(type=type)
            else:   # if nothing is open, open a new calc instance
                print(f'Nothing is opened.')
                tag = self._new_file(type=type)
        else:
            tag = self._open_file(filepath)

        return Calc(tag, soffice=self)

    def Writer(self, filepath=None, new_file=False):
        #
        # cursor = document.Text.createTextCursor()
        # document.Text.insertString(cursor, "This text is being added to openoffice using python and uno package.", 0)
        raise NotImplementedError('writer class not implemented yet')

    def save(self, tag, type, extension, filepath=None):
        """Save  file.

        Args:
            filepath (string or pathlib.Path, optional): filepath to save file.
        """
        if filepath is None and self.get_filepath(tag)=='':  # ok
            temporary_path = Path.cwd()/self.get_title(tag)
            temporary_path.with_suffix(extension)
            if query_yes_no(f'Filepath not defined. Wish to save at {temporary_path}?'):
                filepath = temporary_path
            else:
                print('File not saved')
                return
        elif filepath is None and self.get_filepath(tag)!='':
            filepath = Path(self.get_filepath(tag)).with_suffix(extension)
        else:
            filepath = Path(filepath).with_suffix(extension)

        # save
        self.write(f'URL = uno.systemPathToFileUrl("{filepath}")')
        self.write(f"properties = (PropertyValue('FilterName', 0, '{type}', 0),)")
        self.write(f"document_{tag}.storeAsURL(URL, properties)")
        time.sleep(0.1)
        print(f'Saved at: {filepath}')
        return self.get_filepath(tag)

    def close(self, tag):
        """Close window."""
        self.write(f'document_{tag}.close(True)')
        self.write(f'del document_{tag}')

    def get_filepath(self, tag):
        hasFilepath = str2bool(self.write(f"print(document_{tag}.hasLocation())")[0])
        if hasFilepath: # check if file has filepath
            return self.write(f"print(uno.fileUrlToSystemPath(document_{tag}.getURL()))")[0]
        else: return ''

    def get_title(self, tag):
        return self.write(f"print(document_{tag}.getTitle())")[0]

    def _new_file(self, type):
        """type = calc or writer
        https://wiki.documentfoundation.org/Macros/Basic/Documents
        """

        if type == 'scalc' or type == 'calc':
            self.write("URL = 'private:factory/scalc'")
        elif type == 'swriter' or type == 'writer':
            self.write("URL = 'private:factory/swriter'")
        elif type == 'simpress' or type == 'impress':
            self.write("URL = 'private:factory/SdXImpressDocument'")
        elif type == 'sdraw' or type == 'draw':
            self.write("URL = 'private:factory/sdraw'")
        elif type == 'smath' or type == 'math':
            self.write("URL = 'private:factory/smath'")
        elif type == 'base' or type == 'obase' or type == 'sbase':
            raise NotImplementedError('base not implemented yet')
        else:
            raise ValueError('type must be calc, writer, impress, draw, math, or base')
        tag = self._get_tag() + 1
        self._set_tag(tag)
        print(f'Opening a new file.')
        self.write(f"document_{tag} = desktop.loadComponentFromURL(URL, '_default', 0, ())")
        # time.sleep(0.1)
        title = self.get_title(tag=tag)
        time.sleep(0.1)
        start_time = time.time()
        while title is [] and time.time() < start_time + 10:
            time.sleep(0.1)
            title = self.get_title(tag=tag)

        print(f'Connected with opened file: {title}')
        return tag

    def _open_file(self, filepath):
        tag = self._get_tag() + 1
        self._set_tag(tag)
        print('Filepath was given.')
        filepath = Path(filepath).absolute()
        self.write(f"URL = uno.systemPathToFileUrl('{filepath}')")
        self.write(f'document_{tag} = desktop.loadComponentFromURL(URL, "_default", 0, ())')
        url = self.write(f"print(uno.fileUrlToSystemPath(document_{tag}.getURL()))")[0]
        print(f'Connected with file: {url}')
        return tag

    def _has_open(self):
        # is there something open???
        self.write("document =  desktop.getCurrentComponent()")
        return str2bool(self.write("print('False') if document is None else print('True')")[0])

    def _connect_with_open(self, type):

        self.write(f"document = desktop.getCurrentComponent()")
        implementationName = self.write(f"print(document.getImplementationName())")
        if _name2ImplementationName(type) in implementationName:
            return self._connect_with_current()
        else:
            print(f'current document is not type {type}.')
            print(f'Searching for any opened {type} instances...')
            self.write("documents = desktop.getComponents().createEnumeration()")


            open_new = True
            hasMoreElements = str2bool(self.write("print(documents.hasMoreElements())")[0])
            while hasMoreElements:
                self.write("document = documents.nextElement()")
                implementationName = self.write("print(document.getImplementationName())")
                if _name2ImplementationName(type) in implementationName:
                    tag = self._get_tag() + 1
                    self._set_tag(tag)
                    self.write(f"document_{tag} = document")

                    if self.get_filepath(tag) == '':
                        title = self.get_title(tag)
                        print(f'Connected with opened file: {title}')
                    else:
                        print(f'Connected with opened file: {self.get_filepath(tag)}')
                    return tag
                hasMoreElements = str2bool(self.write("print(documents.hasMoreElements())")[0])
            print(f'No opened {type} instances were found.')
            return False

    def _connect_with_current(self):
        tag = self._get_tag() + 1
        self._set_tag(tag)
        self.write(f"document_{tag} = desktop.getCurrentComponent()")

        print(f'Connecting with current file.')

        time.sleep(0.1)
        if self.get_filepath(tag) != '':
            print(f'Connected with opened file: {self.get_filepath(tag)}')
        else:
            title = self.get_title(tag)
            time.sleep(0.1)
            start_time = time.time()
            while title is [] and time.time() < start_time + 10:
                time.sleep(0.1)
                title = self.get_title(tag)
            print(f'Connected with opened file: {title}')
        return tag


class soffice_python_error(Exception):

    def __init__(self, message):
        self.message = message
        super().__init__(self.message)


class Calc():

    def __init__(self, tag, soffice):

        self.soffice = soffice
        self.tag = tag
        self.sheet_tags = []
        self.filepath = self.get_filepath()

        self.write(f"""def get_property_recursively(tag, column, row, name, attrs=[]):\n""" +\
                   f"""    f = dict()\n""" +\
                   f"""    sheet = eval(f"sheet_{{tag}}")\n""" +\
                   f"""    i = [x.Name for x in sheet.getCellByPosition(column, row).getPropertySetInfo().Properties].index(name)\n""" +\
                   f"""    if attrs == []:\n""" +\
                   f"""        try:\n""" +\
                   f"""            if type(sheet.getCellByPosition(column, row).getPropertyValue(name).value) is int:\n""" +\
                   f"""                return sheet.getCellByPosition(column, row).getPropertyValue(name).value\n""" +\
                   f"""            elif type(sheet.getCellByPosition(column, row).getPropertyValue(name).value) is str:\n""" +\
                   f"""                return sheet.getCellByPosition(column, row).getPropertyValue(name).value\n""" +\
                   f"""            else:\n""" +\
                   f"""                keys = sheet.getCellByPosition(column, row).getPropertyValue(name).value.__dir__()\n""" +\
                   f"""        except AttributeError:\n""" +\
                   f"""            return sheet.getCellByPosition(column, row).getPropertyValue(name)\n""" +\
                   f"""    else:\n""" +\
                   f"""        try:\n""" +\
                   f"""            t = ''\n""" +\
                   f"""            for attr in attrs:\n""" +\
                   f"""                t += f"__getattr__('{{attr}}')."\n""" +\
                   f"""            if type(eval(f"sheet.getCellByPosition(column, row).getPropertyValue(name).{{t}}value")) is str:\n""" +\
                   f"""                return eval(f"sheet.getCellByPosition(column, row).getPropertyValue(name).{{t}}value")\n""" +\
                   f"""            else:\n""" +\
                   f"""                keys = eval(f"sheet.getCellByPosition(column, row).getPropertyValue(name).{{t}}value.__dir__()")\n""" +\
                   f"""        except AttributeError:\n""" +\
                   # f"""            print(t)\n""" +\
                   f"""            return eval(f"sheet.getCellByPosition(column, row).getPropertyValue(name).{{t[:-1]}}")\n""" +\
                   f"""    for key in keys:\n""" +\
                   f"""        f[key] = get_property_recursively(tag, column, row, name, attrs+[key])\n""" +\
                   f"""    return f""")

    def write(self, string, timeout=10):
        return self.soffice.write(string=string, timeout=timeout)

    def read(self, timeout=10):
        return self.soffice.read(timeout=timeout)

    def get_filepath(self):
        return self.soffice.get_filepath(self.tag)

    def get_title(self):
        return self.soffice.get_title(self.tag)

    def save(self, filepath=None, type='ods'):
        if type in ['ods', 'calc8', 'calc', 'Calc']:
            type = 'calc8'
            extension = '.ods'
        elif type in ['excel', 'xlsx']:
            type = 'Calc MS Excel 2007 XML'
            extension = '.xlsx'
        elif type in ['xls']:
            type = 'MS Excel 97'
            extension = '.xls'
        elif type in ['csv', 'text', 'txt']:
            type = 'Text - txt - csv (StarCalc)'
            extension = '.csv'
        else:
            raise ValueError('type not recognised')

        if filepath is None and self.filepath!='' and self.get_filepath()!='':
            if self.filepath != self.get_filepath():
                print('File was last saved in a different filepath than the one stored here.')
                print('Last saved path:' + self.get_filepath())
                print('Stored path:' + self.filepath)
                if query_yes_no(f'Wish to save at {self.get_filepath()}?'):
                    # self.soffice.Calc(filepath()
                    filepath = self.get_filepath()
                else:
                    if query_yes_no(f'Wish to save at {self.filepath}?'):
                        filepath = self.filepath
                    else:
                        print('File not saved')
                        return

        self.filepath = self.soffice.save(tag=self.tag, type=type, extension=extension, filepath=filepath)

    def close(self):
        """Close window."""
        for tag in self.sheet_tags:
            try:
                self.write(f'del document_{tag}')
            except soffice_python_error:
                pass
        self.soffice.close(self.tag)

    def get_sheets_count(self):
        return int(self.write(f"print(document_{self.tag}.Sheets.getCount())")[0])

    def get_sheets_name(self):
        """Returns the sheets names in a tuple."""
        return eval(self.write(f"print(document_{self.tag}.Sheets.getElementNames())")[0])

    def get_sheet_position(self, name):
        try:
            return self.get_sheets_name().index(name)
        except ValueError:
            raise SheetNameDoNotExistError(f'{name} does not exists')

    def get_sheet_name_by_position(self, position):
        names = self.get_sheets_name()
        if position > len(names) or position < 0:
            raise IndexError('Position outside range.')
        return names[position]

    def insert_sheet(self, name, position=None):
        """name can be a string or a list

        position starts from 0. If position = 0, the sheet will be the first one.
        """
        if position is None:
            position = self.get_sheets_count()+1

        if name in self.get_sheets_name():
            raise SheetNameExistError(f'{name} already exists')

        self.write(f"document_{self.tag}.Sheets.insertNewByName('{name}', {position})")

    def remove_sheet(self, name):
        if name in self.get_sheets_name():
            if self.get_sheets_count() > 1:
                self.write(f"document_{self.tag}.Sheets.removeByName('{name}')")
            else:
                raise SheetRemoveError(f"{name} cannot be removed because it is the only existing sheet")
        else:
            raise SheetNameDoNotExistError(f'{name} does not exists')

    def remove_sheets_by_position(self, position):
        self.remove_sheet(self.get_sheet_name_by_position(position))

    def move_sheet(self, name, position):
        if position < 0:
            raise ValueError('position cannot be negative')
        if name in self.get_sheets_name():
            self.write(f"document_{self.tag}.Sheets.moveByName('{name}', {position})")
        else:
            raise SheetNameDoNotExistError(f'{name} does not exists')

    def copy_sheet(self, name, new_name, position):
        if position < 0:
            raise ValueError('position cannot be negative')
        if name in self.get_sheets_name():
            self.write(f"document_{self.tag}.Sheets.copyByName('{name}', '{new_name}', {position})")
        else:
            raise SheetNameDoNotExistError(f'{name} does not exists')

    def get_sheet(self, name):
        if name in self.get_sheets_name():
            tag = self.soffice._get_tag() + 1
            self.write(f"sheet_{tag} = document_{self.tag}.Sheets.getByName('{name}')")
            self.sheet_tags.append(tag)
            return Sheet(tag, self)
        else:
            raise SheetNameDoNotExistError(f'{name} does not exists')

    def get_sheet_by_position(self, position):
        name = self.get_sheet_name_by_position(position)
        return self.get_sheet(name)


class SheetNameExistError(Exception):

    def __init__(self, message):
        self.message = message
        super().__init__(self.message)

class SheetNameDoNotExistError(Exception):

    def __init__(self, message):
        self.message = message
        super().__init__(self.message)

class SheetRemoveError(Exception):

    def __init__(self, message):
        self.message = message
        super().__init__(self.message)

# %%
class Sheet():

    def __init__(self, tag, Calc):
        self.Calc = Calc
        self.tag = tag

    def write(self, string, timeout=10):
        return self.Calc.soffice.write(string=string, timeout=timeout)

    def read(self, timeout=10):
        return self.Calc.soffice.read(timeout=timeout)

    def get_name(self):
        return self.write(f'print(sheet_{self.tag}.getName())')[0]

    def set_name(self, name):
        self.write(f"sheet_{self.tag}.setName('{name}')")

    def isVisible(self):
        return str2bool(self.write(f"print(sheet_{self.tag}.IsVisible)")[0])

    def get_last_row(self):
        """starts from 0"""
        row_name = eval(self.write(f"print(sheet_{self.tag}.getRowDescriptions())")[0])[-1]
        idx = int(row_name.split()[-1])

        visible = False
        while visible == False:
            visible = str2bool(self.write(f"print(sheet_{self.tag}.getRows()[{idx}].IsVisible)")[0])
            idx += 1
        return idx-2

    def get_last_column(self):
        """starts from 0"""
        col_name = eval(self.write(f"print(sheet_{self.tag}.getColumnDescriptions())")[0])[-1]
        idx = int(_letter2num(col_name.split()[-1]))

        visible = False
        while visible == False:
            visible = str2bool(self.write(f"print(sheet_{self.tag}.getColumns()[{idx}].IsVisible)")[0])
            idx += 1
        return idx-1

    def get_row_length(self, row):
        lc = self.get_last_column()
        row = _check_row_value(row)

        if int(self.write(f"print(len(sheet_{self.tag}.getCellRangeByPosition(0, {row}, {lc}, {row}).queryEmptyCells()))")[0]) == 0:
            return lc+1
        else:
            startColumn = int(self.write(f"print(sheet_{self.tag}.getCellRangeByPosition(0, {row}, {lc}, {row}).queryEmptyCells()[-1].RangeAddress.StartColumn)")[0])
            endColumn = int(self.write(f"print(sheet_{self.tag}.getCellRangeByPosition(0, {row}, {lc}, {row}).queryEmptyCells()[-1].RangeAddress.EndColumn)")[0])

            if endColumn == lc:
                return int(self.write(f"print(sheet_{self.tag}.getCellRangeByPosition(0, {row}, {lc}, {row}).queryEmptyCells()[-1].RangeAddress.StartColumn)")[0])
            else:
                return lc+1

    def get_column_length(self, column):
        lr = self.get_last_row()
        column = _check_column_value(column)

        if int(self.write(f"print(len(sheet_{self.tag}.getCellRangeByPosition({column}, 0, {column}, {lr}).queryEmptyCells()))")[0]) == 0:
            return lr+1
        else:
            startRow = int(self.write(f"print(sheet_{self.tag}.getCellRangeByPosition({column}, 0, {column}, {lr}).queryEmptyCells()[-1].RangeAddress.StartRow)")[0])
            endRow = int(self.write(f"print(sheet_{self.tag}.getCellRangeByPosition({column}, 0, {column}, {lr}).queryEmptyCells()[-1].RangeAddress.EndRow)")[0])

            if endRow == lr:
                return int(self.write(f"print(sheet_{self.tag}.getCellRangeByPosition({column}, 0, {column}, {lr}).queryEmptyCells()[-1].RangeAddress.StartRow)")[0])
            else:
                return lr+1

    def set_column_width(self, column, width):
        column = _check_column_value(column)
        if width < 0:
            raise ValueError('width cannot be negative')
        self.write(f"sheet_{self.tag}.getColumns()[{column}].setPropertyValue('Width', {width})")

    def get_column_width(self, column):
        column = _check_column_value(column)
        return int(self.write(f"print(sheet_{self.tag}.getColumns()[{column}].Width)")[0])

    def set_row_height(self, row, height):
        row = _check_row_value(row)
        if height < 0:
            raise ValueError('height cannot be negative')

        self.write(f"sheet_{self.tag}.getRows()[{row}].setPropertyValue('Height', {height})")

    def get_row_height(self, row):
        row = _check_row_value(row)
        return int(self.write(f"print(sheet_{self.tag}.getRows()[{row}].Height)")[0])

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

    def set_row(self, *args, **kwargs):
        """
        row
        value

        row
        value
        format

        row
        value
        column_start

        row
        value
        column_start
        format
        """
        format = 'formula'
        column_start = 0

        if len(args) == 2:
            row_start = args[0]
            value = args[1]

        elif len(args) == 3:
            row_start = args[0]
            value  = args[1]
            if args[2] in ['formula', 'string', 'number']:
                format = args[2]
            else:
                column_start = args[2]

        elif len(args) == 4:
            row_start = args[0]
            value = args[1]
            column_start = args[2]
        elif len(args) == 5:
            row_start = args[0]
            value = args[1]
            column_start  = args[2]
            format        = args[4]


        if 'column_start' in kwargs:
            column_start = kwargs['column_start']
        if 'row_start' in kwargs:
            row_start = kwargs['row_start']
        if 'row' in kwargs:
            row_start = kwargs['row']
        if 'value' in kwargs:
            value = kwargs['value']
        if 'format' in kwargs:
            format = kwargs['format']

        return self.set_cells(value=[value], column_start=column_start, row_start=row_start, format=format)

    def get_row(self, *args, **kwargs):
        """
        row

        row
        format

        row
        column_start
        column_stop

        row
        column_start
        column_stop
        format
        """
        format = 'string'
        column_start = 0

        if len(args) == 1:
            row_start = args[0]
        elif len(args) == 2:
            row_start = args[0]
            format    = args[1]
        elif len(args) == 3:
            row_start = args[0]
            column_start = args[1]
            column_stop = args[2]
        elif len(args) == 4:
            row_start = args[0]
            column_start  = args[1]
            column_stop   = args[2]
            format        = args[3]


        if 'column_start' in kwargs:
            column_start = kwargs['column_start']
        if 'row_start' in kwargs:
            row_start = kwargs['row_start']
        if 'row' in kwargs:
            row_start = kwargs['row']
        if 'column_stop' in kwargs:
            column_stop = kwargs['column_stop']
        if 'format' in kwargs:
            format = kwargs['format']

        try:
            type(column_stop)
        except NameError:
            column_stop  = self.get_row_length(_check_row_value(row_start)) -1

        return self.get_cells(column_start=column_start, column_stop=column_stop, row_start=row_start, row_stop=row_start, format=format)

    def set_column(self, *args, **kwargs):
        """
        column
        value

        column
        value
        format

        column
        value
        row_start

        column
        value
        row_start
        format
        """
        format = 'formula'
        row_start = 0

        if len(args) == 2:
            column_start = args[0]
            value = args[1]
        elif len(args) == 3:
            column_start = args[0]
            value  = args[1]
            if args[2] in ['formula', 'string', 'number']:
                format = args[2]
            else:
                row_start = args[2]
        elif len(args) == 4:
            column_start = args[0]
            value     = args[1]
            row_start = args[2]
            format    = args[3]

        if 'column_start' in kwargs:
            column_start = kwargs['column_start']
        if 'column' in kwargs:
            column_start = kwargs['column']
        if 'row_start' in kwargs:
            row_start = kwargs['row_start']
        if 'value' in kwargs:
            value = kwargs['value']
        if 'format' in kwargs:
            format = kwargs['format']

        if isinstance(value, list):
            for v in value:
                if not isinstance(v, list):
                    value = transpose(value)
                    break
        else:
            raise TypeError('value must be a list or a list of lists or a numpy array')



        return self.set_cells(value=value, column_start=column_start, row_start=row_start, format=format)

    def get_column(self, *args, **kwargs):
        """
        column

        column
        format

        column
        row_start
        row_stop

        column
        row_start
        row_stop
        format
        """
        format = 'string'
        row_start = 0

        if len(args) == 1:
            column_start = args[0]

        elif len(args) == 2:
            column_start = args[0]
            format    = args[1]

        elif len(args) == 3:
            column_start = args[0]
            row_start = args[1]
            row_stop = args[2]
        elif len(args) == 4:
            column_start = args[0]
            row_start  = args[1]
            row_stop   = args[2]
            format     = args[3]


        if 'column_start' in kwargs:
            column_start = kwargs['column_start']
        if 'column' in kwargs:
            column_start = kwargs['column']
        if 'row_start' in kwargs:
            row_start = kwargs['row_start']
        if 'row_stop' in kwargs:
            row_stop = kwargs['row_stop']
        if 'format' in kwargs:
            format = kwargs['format']

        try:
            type(row_stop)
        except NameError:
            row_stop  = self.get_column_length(_check_column_value(column_start)) -1

        return self.get_cells(column_start=column_start, column_stop=column_start, row_start=row_start, row_stop=row_stop, format=format)

    def cell_properties(self, *args, **kwargs):
        if len(args) == 1:
            column, row = cell2num(args[0])
        elif len(args) == 2:
            column = args[0]
            row = args[1]
        else:
            raise valueError('Missing inputs')

        column = _check_column_value(column)
        row = _check_row_value(row)

        return eval(self.write(f"print([x.Name for x in sheet_{self.tag}.getCellByPosition({column}, {row}).getPropertySetInfo().Properties])")[0])

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

    def get_cell_property(self, *args, **kwargs):
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

        column = _check_column_value(column)
        row = _check_row_value(row)

        t = self.write(f"print(type(get_property_recursively({self.tag}, {column}, {row}, '{name}')))")[0]
        try:
            output = self.write(f"print(get_property_recursively({self.tag}, {column}, {row}, '{name}'))")[0]
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





    def set_cell_property(self, *args, **kwargs):
        if len(args) == 2:
            column, row = cell2num(args[0])
            name = args[1]
            value = args[2]
        elif len(args) == 3:
            column = args[0]
            row = args[1]
            name = args[2]
            value = args[3]


        column = _check_column_value(column)
        row = _check_row_value(row)

        return self._get_property_recursive(column, row, name)



d = dict()
for i in range(0, 104):
    print(i)
    n = sheet.cell_properties(1, 1)[i]
    print(n)
    print(sheet.get_cell_property(1, 1, n))
    print('='*20)


column = 1
row = 1
name = 'TopBorder'
d = sheet.get_cell_property(column, row, name)
print(d['Color'])
d['Color'] = 95254354
sheet.write(f"from com.sun.star.table import BorderLine2")
sheet.write(f"a = BorderLine2(**{d})")
sheet.write(f"print(a)")[0]
sheet.write(f"sheet_{sheet.tag}.getCellByPosition({column}, {row}).setPropertyValue('{name}', a)")


# %%




# %%





    def get_cells_properties(self, property, row_start, col_start, row_stop, col_stop):
        row_start = _check_row_value(row_start)[0]
        col_start = _check_col_value(col_start)[0]
        row_stop = _check_row_value(row_stop)[0]
        col_stop = _check_col_value(col_stop)[0]

        if type(property) == str:
            property = [property]

        object_cell = self.object.get_cell_range_by_position(col_start, row_start, col_stop, row_stop)
        attr = getattr(object_cell, property[0])

        for idx in range(1, len(property)):
            attr = getattr(attr, property[idx])

        try:
            return attr, attr.value.__dir__()
        except AttributeError:
            return attr, None

    def set_cells_properties(self, property, value, row_start, col_start, row_stop, col_stop):
        row_start = _check_row_value(row_start)[0]
        col_start = _check_col_value(col_start)[0]
        row_stop = _check_row_value(row_stop)[0]
        col_stop = _check_col_value(col_stop)[0]

        if type(property) == str:
            property = [property]

        self.object.get_cell_range_by_position(col_start, row_start, col_stop, row_stop).setPropertyValue(property[0], value)

    def get_cell_object(self, row=1, col=1):
        row = _check_row_value(row)[0]
        col = _check_col_value(col)[0]

        return self.object.get_cell_by_position(col, row)



    def get_cell_formating(self, row, col, extra=None):
        """font : bold, font name, font size, italic, color
        text: vertical justify,horizontal justify
        cell: border, color
        conditional formating: conditional formating
        """

        p = ['FormatID']
        p += ['CharWeight', 'CharFontName', 'CharHeight', 'CharPosture', 'CharColor']
        p += ['VertJustify', 'HoriJustify']
        p += ['CellBackColor', 'TableBorder', 'TableBorder2']
        p += ['ConditionalFormat']

        if extra is not None:
            p += extra

        p_list = []
        for property in p:
            obj, _ = self.get_cell_property(property, row=row, col=col)
            p_list.append(obj)

        return p_list

    def set_cell_formating(self, obj_list, row, col, extra=None):
        """font : bold, font name, font size, italic, color
        text: vertical justify,horizontal justify
        cell: border, color
        conditional formating: conditional formating
        """
        p = ['FormatID']
        p += ['CharWeight', 'CharFontName', 'CharHeight', 'CharPosture', 'CharColor']
        p += ['VertJustify', 'HoriJustify']
        p += ['CellBackColor', 'TableBorder', 'TableBorder2']
        p += ['ConditionalFormat']

        if extra is not None:
            p += extra

        for obj, property in zip(obj_list, p):
            self.set_cell_property(property, value=obj, row=row, col=col)

    def get_cells_formatting(self, row_start=1, col_start=1, row_stop=None, col_stop=None, extra=None):
        if row_stop is None:
            row_stop = self.get_last_row()
        if col_stop is None:
            col_stop = self.get_last_col()

        row_stop = _check_row_value(row_stop)[0]
        col_stop = _check_col_value(col_stop)[0]

        cell_formatting_list = []
        for idx, row in enumerate(range(row_start, row_stop+2)):
            cell_formatting_list.append([])
            for col in range(col_start, col_stop+2):
                p_list = self.get_cell_formating(row=row, col=col, extra=extra)
                cell_formatting_list[idx].append(p_list)
        return cell_formatting_list

    def set_cells_formatting(self, cell_formatting_list, row_start=1, col_start=1, extra=None):

        for idx, row in enumerate(range(row_start, row_start+len(cell_formatting_list))):
            for idx2, col in enumerate(range(col_start, col_start+len(cell_formatting_list[idx]))):
                self.set_cell_formating(cell_formatting_list[idx][idx2], row=row, col=col, extra=extra)

    def get_merged(self, ):
        merged_ranges = []
        #################
        start=[]
        stop = []
        #################

        ucf = self.object.getUniqueCellFormatRanges()
        for ranges in ucf:
            rgtest = ranges.getByIndex(0)
            if rgtest.getIsMerged():
                for rg in ranges:
                    oCursor = rg.getSpreadsheet().createCursorByRange(rg)
                    oCursor.collapseToMergedArea()
                    ######################
                    row_start = int(oCursor.getRowDescriptions()[0].split(' ')[-1])-1
                    row_end = int(oCursor.getRowDescriptions()[-1].split(' ')[-1])-1
                    for row in range(row_end-row_start+1):
                        col_start = backpack.libremanip._letter2num(oCursor.getColumnDescriptions()[0].split(' ')[-1])-1
                        col_end = backpack.libremanip._letter2num(oCursor.getColumnDescriptions()[-1].split(' ')[-1])-1
                        for col in range(col_end-col_start+1):
                            if oCursor.getCellByPosition(col, row).IsMerged:
                                # print(row+row_start, col+col_start)
                                start.append([row+row_start, col+col_start])
                            else:
                                stop.append([row+row_start, col+col_start])
                                # print(f'not: {row+row_start}, {col+col_start}')
                    ######################
                    addr = oCursor.getRangeAddress()

                    col_start = addr.StartColumn+1
                    row_start = addr.StartRow+1
                    col_stop = addr.EndColumn+1
                    row_stop = addr.EndRow+1
                    merged_ranges.append([row_start, col_start, row_stop, col_stop])
        return merged_ranges

    def merge(self, row_start, col_start, row_stop, col_stop):
        row_start = _check_row_value(row_start)[0]
        col_start = _check_col_value(col_start)[0]
        row_stop = _check_row_value(row_stop)[0]
        col_stop = _check_col_value(col_stop)[0]

        sheet_data = self.object.get_cell_range_by_position(col_start, row_start, col_stop, row_stop)
        sheet_data.merge(True)

# %%
p = start_soffice()
s = soffice()
# %%
c = s.Calc()
# c = s.Calc(filepath='/home/galdino/github/pycalc/dddd.ods')
# c = s.Calc(new_file=True)
# %%
sheet = c.get_sheet_by_position(0)
# %%
c.close()
kill(p)
s.kill()
kill_tmux()
