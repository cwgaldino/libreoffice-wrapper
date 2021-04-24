===================
libreoffice-wrapper
===================

Manipulate libreoffice programs (Writer, Calc, etc...) via python. The document is updated real time (no need to reload).

- Tested on Linux (Ubuntu 20.04)
- Not tested on windows or mac yet

The principle of this module is to use tmux to intermediate communication between a python instance and the libreoffice's internal python interpreter. This way your are not limited to the functionality of libreoffice's internal python.

Still in development. Let me know if you wanna help!

All the core functionality is already working. Now, I'm working on manipulating documents, more specifically, Calc spreadsheets, see below.

[x] core spreadsheet functionality (open, save, close)

[x] add/remove styles

[x] insert/delete/move sheets

[x] get/set values from a cell/range

[x] get/set values from a row/column

[x] get/set cell/range properties

[x] merge cells

[ ] conditional formatting

[ ] validation formatting

[ ] document the code

[ ] write examples/tutorials

Note that, right now I'm only interested in being able to manipulate Calc spreadsheets and I'm not sure I will extend the functionality to other types of documents like Writer, Impress, etc.. However, it should be easy enough to implement code for other types of documents since the core functionality is the same. In fact. I started doing something for Writer. Let me know if you're interested in that and I can upload the code.
