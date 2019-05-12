===================================
'docxtable2xlsx' task version 0.1.0
===================================

'docxtable2xlsx' task extracts tables from Microsoft word file and
save them in a Excel file or a CSV file.

Installation
------------

Before installing 'docxtable2xlsx' task, please make sure that 'pyloco' is installed.
Run the following command if you need to install 'pyloco'. ::

    >>> pip install pyloco

Or, if 'pyloco' is already installed, upgrade 'pyloco' with the following command ::

    >>> pip install -U pyloco

To install 'docxtable2xlsx' task, run the following 'pyloco' command.  ::

    >>> pyloco install docxtable2xlsx

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

