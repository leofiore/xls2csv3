xls2csv3
--------

this is a slightly modified version of 

Cli command
-----------

Usage: xls2csv -i <IFNAME> -o <OFNAME>

Options:
  --version             show program's version number and exit
  -h, --help            show this help message and exit
  -i IFNAME, --input=IFNAME
                        input Excel filename
  -o OFNAME, --output=OFNAME
                        output CSV filename
  -s NUMSHEET, --sheet=NUMSHEET
                        sheet number to convert (1st sheet is numbered '0', so
                        it's 0 by default)
  -p SEPARATOR, --sep=SEPARATOR
                        separator used in the csv file (';' as Excel default
                        conversion character)
  -e, --enclose-text    enclose text values into double-quote characters
  --input-encoding=INPUTENCODING
                        override the input file encoding (useful for excel 95
                        and earlier versions)
  --output-encoding=OUTPUTENCODING
                        set the output file encoding (utf-8 by default)
  --col-as-int=COLASINT
                        give column numbers as a list with ':' as separator,
                        like 1:25:41 or 'all' for converting all colums
                        For these columns, if the cell contains a number, it
                        will be considered as an integer
  --remove-rows=REMROWS
                        give row numbers to remove as a list with ':' as
                        separator, like 2:3:56
  --remove-cols=REMCOLS
                        give column numbers to remove as a list with ':' as
                        separator, like 6:89:7
  --lineend=LINEEND     character for line ending : CRLF (windows) or LF
                        (unix) default
  --stats               print statistics about the Excel file
  -q, --quiet           don't print status messages to stdout
