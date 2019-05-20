===================
docxtable2xlsx task
===================

version: 0.1.4

extracts tables from Microsoft word file and save them in a Excel file or a CSV file

'docxtable2xlsx' task extracts tables from Microsoft word file and
save them in a Excel file or a CSV file.


Installation
------------

To install docxtable2xlsx task, run the following pyloco command. ::

    >>> pyloco install docxtable2xlsx
    >>> pyloco docxtable2xlsx --version
    docxtable2xlsx 0.1.4

If pyloco is not available on your computer, please run the following
command to install pyloco, and try again above task installation. ::

    >>> pip install pyloco --user
    >>> pyloco --version
    pyloco 0.0.108

.. note::

    - 'pip' is a Python package manager.
    - Remove '--user' option to run pyloco on a virtual environment.
    - Add '-U' option to 'pip' command to upgrade pyloco.
    - We recommend to use pyloco version 0.0.108 or higher.

Command-line syntax
-------------------

usage: pyloco docxtable2xlsx [-h] [-t type] [--general-arguments]
                                 data 

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
