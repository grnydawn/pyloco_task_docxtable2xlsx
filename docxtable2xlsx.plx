path@arg = "data", type=str, help="input MS Word file"
type@arg = "-t", "--type", metavar="type", default="xlsx", \
           help="output file format (default='xlsx')"
tables@forward = "data", help="output tables"

tables@pyloco = docx2text my.docx -- pydict2xlsx -t csv -o tables.csv

[forward*]

tables = tables[1]

[attribute*]

_name_ = "docxtable2xlsx"
_version_ = "0.1.0"
__doc__ = """extracts tables from Microsoft word file and save them in a Excel file or a CSV file

'docxtable2xlsx' task extracts tables from Microsoft word file and
save them in a Excel file or a CSV file.
"""
_install_requires_ = ["docx2text", "pydict2xlsx"]