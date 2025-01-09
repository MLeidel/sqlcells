## SQLcells
**A Python Desktop GUI  
Using SQL to query and output  
with multiple Spreadsheets and csv files**

Querys may be saved, edited, and rerun from the command line or the Desktop.

![program](images/sqlcells.png "SQLcells.py")

### Spreadsheet inputs and result

![program](images/spreadsheets.png "spreadsheets")

When _Launch_ is checked, the result is opened in LibreOffice Calc.  
When _Log_ is checked, the input file paths and SQL code is appended to a log file.

There is a limit of seven **input files** (allowed formats: `.xlsx`, `.xls`, `.csv`)

The **output file** may be any of these formats: `.xlsx`, `.xls`, `.csv`, or `.db`, `.sqlite`

---

Clicking on an input file lets you open the spreadsheet/csv or view the columns and data types.

![program](images/viewing.png "SQLcells.py")

The following is an example of a saved query setup file:

    d1: /home/x/y/z/projects/sqlcells/testfiles/sampledatainsnames.xlsx
    d2: /home/x/y/z/projects/sqlcells/testfiles/sampledatainsurance.xlsx
    SQL
    select d1.Policy, last_name, first_name, Expiry, State, InsuredValue, email
    	from d1, d2
    	where d1.Policy = d2.Policy
    	order by State, Expiry, last_name
    OUTPUT
    out.xlsx
    LAUNCH
    LOG
    
An existing query can be run in an _unattended_ (_batch mode_) by using a saved query setup file
as an argument at startup:

        $ python3 sqlcells.py sql_sample.txt

The Launch and Log options will apply as they were set when saved.

_For Windows note: xlrd may need to be upgraded_
