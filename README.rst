===================================
'docxtable2xlsx' task version 0.1.2
===================================

'docxtable2xlsx' task extracts tables from Microsoft word file and
save them in a Excel file or a CSV file.

Installation
------------

Before installing 'docxtable2xlsx' task, please make sure that 'pyloco' is installed.
Run the following command if you need to install 'pyloco'. ::

    >>> pip install pyloco --user
    >>> pyloco --version

.. note::

    - 'pip' is a Python package manager. See `here <https://www.w3schools.com/python/python_pip.asp/>`_ for more information.
    - Add '-U' option in case to upgrade pyloco.
    - Remove '--user' option in case that pyloco is running in a virtualenv.

.. note::

    'docxtable2xlsx' task is created using pyloco version '0.0.105'.
    It is recommended to run 'docxtable2xlsx' with pyloco version '0.0.105' or higher

To install 'docxtable2xlsx' task, run the following 'pyloco' command.  ::

    >>> pyloco install docxtable2xlsx
    >>> pyloco docxtable2xlsx --version

Command-line syntax
-------------------

usage: pyloco docxtable2xlsx [-h] [-t type] [--general-arguments] data 

extracts tables from Microsoft word file and save them in a Excel file or a CSV file

positional arguments:
  data                  input MS Word file

optional arguments:
  -h, --help            show this help message and exit
  -t type, --type type  output file format (default='xlsx')
  --general-arguments   Task-common arguments. Use --verbose to see a list of
                        general arguments

forward output variables:
   data                 output tables


Examples
--------

Assuming 'my.docx' MS word file has tables in it, following command produces an Excel or
a CSV file that contains the content of tables in 'my.docx'. ::

    >>> pyloco docxtable2xlsx my.docx           # generates 'tables.xlsx'
    >>> pyloco docxtable2xlsx my.docx -t csv    # generates 'tables.csv'
