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
    # if libreoffice is not None:
    #     self.libreoffice = Path(libreoffice)
    # else:
    #     self.libreoffice = Path('/opt/libreoffice7.0/program/')

    # tmux server
    if tmux_config:
        tmux_server = libtmux.Server(config_file='tmux.conf')
    else:
        tmux_server = libtmux.Server()

    # new session
    if tmux_server.has_session('py_libreoffice'):
        tmux_session = tmux_server.find_where({'session_name':'py_libreoffice'})
    else:
        tmux_session = tmux_server.new_session('py_libreoffice')

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
                [pid2.append(x) for x in _get_children(pid)]
        else:
            pid2 = args
        for pid in pid2:
            try:
                os.kill(pid, signal.SIGKILL)
            except ProcessLookupError:
                pass
    except TypeError:
        if recursive:
            pid2 = []
            for pid in args[0]:
                [pid2.append(x) for x in _get_children(pid)]
        else:
            pid2 = args
        for pid in pid2:
            try:
                os.kill(pid, signal.SIGKILL)
            except ProcessLookupError:
                pass

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

class soffice():

    def __init__(self, tmux_config=True, port=8100, folder='/opt/libreoffice7.0'):
        """
        server -> session (pycalc) -> window (soffice-python) -> pane (0)
        """
        timeout_python = 10
        timeout_soffice = 60

        # set libreoffice folder
        self.folder = Path(folder)/'program'
        # if libreoffice is not None:
        #     self.libreoffice = Path(libreoffice)
        # else:
        #     self.libreoffice = Path('/opt/libreoffice7.0/program/')

        # tmux server
        if tmux_config:
            self.tmux_server = libtmux.Server(config_file='tmux.conf')
        else:
            self.tmux_server = libtmux.Server()

        # new session
        if self.tmux_server.has_session('py_libreoffice'):
            self.tmux_session = self.tmux_server.find_where({'session_name':'py_libreoffice'})
        else:
            self.tmux_session = self.tmux_server.new_session('py_libreoffice')

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

    def kill_tmux(self):
        self.tmux_server.kill_server()

    def check_running(self):
        if self.tmux_pane.capture_pane()[-1] in ['>>>', '... >>>', '... ... >>>']:
            return False
        else:
            return True

    def write(self, string, timeout=10, max_output=1000):
        # check quotes

        keys2send = f"try: exec(\"\"\"{string}\"\"\")" + '\n'\
                     """except: print('calc-ERROR'); print(traceback.format_exc())"""
        self.tmux_pane.send_keys(keys2send, enter=True, suppress_history=False)
        self.tmux_pane.enter()
        time.sleep(0.05)

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

    def __init__(self, tag, soffice=None):

        self.soffice = soffice
        self.tag = tag

        self.filepath = self.get_filepath()

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
        self.soffice.close(self.tag)

    def get_sheets_count(self):
        return int(self.write(f"print(document_{self.tag}.Sheets.getCount())")[0])

    def get_sheets_name(self):
        """Returns the sheets names in a tuple."""
        return eval(self.write(f"print(document_{self.tag}.Sheets.getElementNames())")[0])

    def insert_sheet(self, name, position=None):
        """name can be a string or a list

        position starts from 1. If position = 1, the sheet will be the first one.
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
        names = self.get_sheets_name()

        if position > len(names) or position < 1:
            raise IndexError('Position outside range.')

        if len(names) == 1:
            raise SheetRemoveError(names[position-1])

        self.object.remove_sheets_by_name(names[position-1])

    def get_sheet_by_name(self, name):
        return sheet(name, self)

    def get_sheets(self, name=None):
        if name is None:
            name = self.get_sheets_name()
        else:
            if type(name) == str:
                name = [name]

            if type(name) == int:
                name = [name]

        sheet_objects = []
        for n in name:
            if type(n) == str:
                sheet_objects.append(self.get_sheet_by_name(n))
            if type(n) == int:
                sheet_objects.append(self.get_sheets_by_position(n))
        if len(sheet_objects) == 1:
            return sheet_objects[0]
        else:
            return sheet_objects

    def get_sheets_by_position(self, position):
        names = self.get_sheets_name()

        if type(position) == int:
            position = [position]

        outside_range = []
        for p in position:
            if p > len(names) or p < 1:
                outside_range.append(p)
        if len(outside_range) > 0:
            raise IndexError(f'Positions {outside_range} outside range.')

        if len(position) == 1:
            return sheet(names[position[0]-1], self)
        else:
            return [sheet(names[p-1], self) for p in position]

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


s = soffice()
# %%
c = s.Calc()
# c = s.Calc(filepath='/home/galdino/github/pycalc/dddd.ods')
# c = s.Calc(new_file=True)
c.get_sheets_name()
c.get_sheets_count()
c.save()
