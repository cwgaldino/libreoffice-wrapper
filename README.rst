===================
libreoffice-wrapper
===================

Python module for controlling `LibreOffice`_ programs (Writer, Calc, Impress, Draw, Math, and Base). Currently, manipulation of Calc instances is (somewhat) fully implemented and the module supports features such as:

[x] core functionality (open, save, close, ...)

[x] save multiple formats (ods, xlsx, ...)

[x] add/remove styles

[x] insert/delete/move sheets

[x] get/set values from a cell/range/rows/columns

[x] get/set cell/range properties (color, border, ...)

[x] merge/unmerge cells

[x] conditional formatting

 Manipulation of Writer, Impress, Draw, and Math instances is in its early development and the module only allows for basic core functionality such as opening/closing/saving files. Base is not implemented at all and trying to open a LibreOffice Base instance will raise an error.

About
==========

This module uses `tmux`_ to intermediate communication between a Python instance and the LibreOffice's internal python interpreter, which has free access to LibreOffice's Python API that allows controlling the LibreOffice components. This way, you are not limited to the functionality of LibreOffice's internal python, i.e., one is able to manipulate LibreOffice components from any Python terminal and inside Python environments. In addition to that, modifications to a file happen "real time" (no need to reload the file).

It was tested on:

- Linux (Ubuntu 20.04) and LibreOffice Version 7.0.4.2

and it should also work fine on MacOS. Currently, I don't think it will work on Windows since `tmux` is not implement there. However, it might work if one uses Windows Subsystem for Linux (WSL). I'm still trying to make it work.

 DISCLAIMER: At first, this module was built to allow manipulating of Calc spreadsheets without the need to reload the document every time a modification was made. Since, this is done, I'm not sure I will keep working on it in order to extend the functionality to Writer, Impress, etc.. In any case, it should be easy enough to implement code for those since the core functionality is the same.


Dependencies
=============

This module is heavily dependent on `tmux`_ which can be installed via apt-get on Debian or Ubuntu (check `tmux`_ page for installation instruction on other OS)::

  apt install tmux

The package `libtmux`_ (tmux workspace manager in python) is also necessary::

  pip install libtmux

Lastly, some features for dealing with Calc instances uses `numpy`_::

  pip install numpy


Usage (Initialization)
=======================

Firstly, one has to start the office in Listening Mode. This can be done by opening the terminal and issuing the command::

  soffice -accept=socket,host=0,port=8100;urp;

Alternatively, libreoffice-wrapper has a built-in function that starts LibreOffice in Listening Mode,

.. code-block:: python

    import libreoffice_wrapper as lw

    pid = lw.start_soffice()


.. The function :python:`lw.start_soffice()` returns the pid of the process. Note that, this function starts a ``tmux`` session called ``libreoffice-wrapper`` with a window named ``soffice``, which can be accessed on a different terminal via ``tmux``. In addition to that, ```lw.start_soffice()``` searches for LibreOffice in the default folder ``/opt/libreoffice7.0``. If LibreOffice is installed in a different folder, it must be passed as an argument of the function ```lw.start_soffice(folder=<path-to-libreoffice>)```.

Once LibreOffice has been started on listening mode, one can now establish the communication,

.. code-block:: python

  soffice = lw.soffice()

.. where `lw.soffice()` starts a `tmux` session `'libreoffice-wrapper'` with a window named `'python'`, with opens the internal LibreOffice's Python interpreter. After that, the `soffice` object manages to communicate to LibreOffice through this Python instance opened in this `tmux` window.

In the end one has to close LibreOffice and close the communication port,

.. code-block:: python

  soffice.kill()

.. which just ends the `tmux` session.

Calc
========

Example:

.. code-block:: python

  import libreoffice_wrapper as lw

  # start LibreOffice and establish communication
  pid = lw.start_soffice()
  soffice = lw.soffice()

  # Open Calc
  calc = soffice.Calc()  # it will try connect with any open Calc instance. If nothing is open, it will start a new spreadsheet
  # calc = soffice.Calc('<path-to-spreadsheet-file>')  # connects/opens specific file
  # calc = soffice.Calc(force_new=True)  # open a new file

  # Calc info
  print(calc.get_filepath())
  print(calc.get_title())
  print(calc.get_sheets_count())
  print(calc.get_sheets_name())

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
  calc.new_style(name='my_new_style', properties={'CellBackColor':-1}, overwrite=False)
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
  sheet.move(position=1)

  # remove (delete)
  # sheet.remove()

  # last used row/column
  print(sheet.get_last_row())
  print(sheet.get_last_column())





  sheet.set_row_height([0, 1, 2, 3, 4, 5, 6, 7, 8 , 9], [10, 20, 30, 40, 500, 60, 70, 80, 90, 20])
  sheet.cell_properties(1, 1)

  #
  sheet.get_cell_property(2, 2, 'CellBackColor')
  sheet.set_cell_property(5, 5, 'CellBackColor', 16776960)
  sheet.get_cell_property(2, 2, 'CellBackColor')

  #
  sheet.get_cell_property(2, 2, 'TopBorder')
  sheet.set_cell_property(2, 2, 'TopBorder.LineWidth', 10)
  sheet.get_cell_property(2, 2, 'TopBorder')

  #
  d = sheet.get_cell_property(5, 5, 'TopBorder')
  d['LineWidth'] = 7
  sheet.set_cell_property(5, 5, 'TopBorder', d)
  sheet.get_cell_property(5, 5, 'TopBorder')

  # saving modifications
  calc.save()

  # finishing up
  calc.close()
  soffice.kill()




Writer, Impress, Draw, Math and Base
======================================

Manipulation of Writer, Impress, Draw, and Math instances are in its early development and the module only allows for basic core functionality such as opening/closing/saving files. Base is not implemented at all and trying to open a LibreOffice Base instance will raise an error.

.. code-block:: python

  import sys
  sys.path.append('<path-to-libreoffice-wrapper>')

  import libreoffice_wrapper as lw

  # %% start LibreOffice
  pid = lw.start_soffice()
  soffice = lw.soffice()

  # %% Writer
  writer = soffice.Writer()
  writer.save()
  writer.close()

  # %% Impress
  impress = soffice.Impress()
  impress.save()
  impress.close()

  # %% Draw
  draw = soffice.Draw()
  draw.save()
  draw.close()

  # %% Math
  math = soffice.Math()
  math.save()
  math.close()

  # %% close LibreOffice
  soffice.kill()



.. _tmux: https://github.com/tmux/tmux/wiki
.. _LibreOffice: https://www.libreoffice.org/
.. _libtmux: https://github.com/tmux-python/libtmux
.. _numpy: https://numpy.org/
