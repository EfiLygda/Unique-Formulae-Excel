# Unique-Formulae-Excel
Simple script for printing the unique formulae used in an Excel workbook, run from the command line.

--------------

## Requirements

```
python==3.8.18
openpyxl==3.1.2
```

## Usage

**Step 1:** Download the Python script `Unique_Formulae_Excel.py`.

**Step 2:** Navigate to the directory of the Excel file through the command line (i.e. using `cd`).

**Step 3:** For Windows users, run the command as:
```
py /path/to/script/Unique_Formulae_Excel.py -f Workbook.xlsx
```
The formulae will be printed in a new line each.

In case of wanting to add a separator to the names the use the --sep argument, as below: 
```
py /path/to/script/Unique_Formulae_Excel.py -f Workbook.xlsx -s ", "
```



## Help

```
usage: Unique_Formulae_Excel.py [-h] --filename FILENAME [--sep SEP]

options:
  -h, --help            show this help message and exit
  --filename FILENAME, -f FILENAME
                        (str) Name of excel workbook containing the formulae.
  --sep SEP, -s SEP     (str) The separator of the formulae names. If none is give they will be printed in a new line
                        each.
```
