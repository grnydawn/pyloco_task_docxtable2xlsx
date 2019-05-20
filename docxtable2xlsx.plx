path@arg = "data", type=str, help="input MS Word file"
type@arg = "-t", "--type", metavar="type", default="xlsx", \
           help="output file format (default='xlsx')"
tables@forward = "data", help="output tables"

tables@pyloco = docx2text '__{path:arg}__' -- pydict2xlsx -t '__{type:arg}__' -o 'tables.__{type:arg}__'

[forward*]

tables = tables[1]

[attribute*]

_name_ = "docxtable2xlsx"
_version_ = "0.1.4"
_doc_ = """extracts tables from Microsoft word file and save them in a Excel file or a CSV file

'docxtable2xlsx' task extracts tables from Microsoft word file and
save them in a Excel file or a CSV file.

Examples
--------

Assuming 'my.docx' MS word file has tables in it, following command produces an Excel or
a CSV file that contains the content of tables in 'my.docx'. ::

    >>> pyloco docxtable2xlsx my.docx           # generates 'tables.xlsx'
    >>> pyloco docxtable2xlsx my.docx -t csv    # generates 'tables.csv'
"""
_install_requires_ = ["docx2text", "pydict2xlsx"]
