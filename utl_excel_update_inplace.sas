%let pgm=utl_excel_update_inplace;

Adding columns and updates to an existing worksheet in place

Adding columns and updates to an existing worksheet in place

I not sure I would put this into production. It is an interesting
academic exercise. This should work everywhere except (VMS and IBM mainframe).

* I don't think it is possible to update excel (in place) with implicit or explicit pass through.
* I am resonably sure R , Python and Perl along with DDE/VBA can do it.
* SAS provides an IML interface to R.

Here is the problem

5513  proc sql;
5514     alter table xel.class
5515        add MAR2001
ERROR: The HEADER/VARIABLE UPDATE function is not supported by the EXCEL engine.
ERROR: View XEL.class cannot be altered.

WHAT I DID

  1. Built a EXCEL template with 13 columns (Employee JAN2001-DEC2001)
  2. Supressed the header text for MAR2001-DEC2001 (D-M)
  3. Added the header for MAR2001
  4. Updated new 'MAR2001' header with new SAS data

There may be a serious flaw eith this. I was only able to add
character columns.


HAVE (Two SAS datasets)
=======================

Up to 40 obs WORK.HAVE total obs=3 13 variables

Obs EMPLOYEE           JAN2001            FEB2001        MAR2001  ....  DEC2001

 1    MARY            $2,048.00         $57,842.00
 2    JOHN           $39,309.00         $61,363.00
 3    TED            $70,437.00         $92,849.00


Up to 40 obs WORK.UPDATES total obs=3

Obs    EMPLOYEE              MAR2001

 1       MARY               $2,048.00
 2       JOHN              $57,842.00
 3       TED               $39,309.00

WANT (CREATE AND UPDATE EXCEL IN PLACE)
========================================

d:/xls/addcol.xlsx

 +------------------+------------+------------+-----------+------------+-----------+
 |      |    A      |     B      |    C       |    D      |   ...      |    M      |
 +------+-----------+------------+------------+-----------+------------+-----------+
 |      |           |            |            |           |            |           |
 |    1 | EMPLOYEE  | JAN2001    | FEB2001    |           |            |           |
 |    2 |           |            |            |           |            |           |
 |    3 |   MARY    |  $2,048.00 | $57,842.00 |           |            |           |
 |    3 |   JOHN    | $39,309.00 | $61,363.00 |           |            |           |
 |    3 |   TED     | $70,437.00 | $92,849.00 |           |            |           |
 +------------------+------------+------------+-----------+------------+-----------+

Update d:/xls/addcol.xlsx in place)

d:/xls/addcol.xlsx
                                                                           DEC2001
 +------------------+------------+------------+------------+------------+-----------+
 |      |    A      |     B      |    C       |    D       |   ...      |     M     |
 +------+-----------+------------+------------+------------+------------+-----------+
 |      |           |            |            |            |            |           |
 |    1 | EMPLOYEE  | JAN2001    | FEB2001    |  MAR2001   |            |           |
 |    2 |           |            |            |            |            |           |
 |    3 |   MARY    |  $2,048.00 | $57,842.00 | $2,048.00  |            |           |
 |    3 |   JOHN    | $39,309.00 | $61,363.00 | $57,842.00 |            |           |
 |    3 |   TED     | $70,437.00 | $92,849.00 | $39,309.00 |            |           |
 +------------------+------------+------------+-----------+------------+-----------+


KEY WORKING CODE
================

   Surpress header and create template with 13 columns
      proc print data=have label  split='#'  noobs style(column)={cellwidth=1in};

   Get the header
     libname xel "&fyl" scan_text=no header=no;
     set xel.'addcol$A1:Z1'n;

   Add column name MAR2001
     update xls.'addcol$'n
     set f4='MAR2001'
     where f1="EMPLOYEE";

   Update new column
     modify xel.'addcol$'n updates;
     by employee;


FULL SOLUTION
=============

*__  __       _              ____        _
|  \/  | __ _| | _____      |  _ \  __ _| |_ __ _
| |\/| |/ _` | |/ / _ \_____| | | |/ _` | __/ _` |
| |  | | (_| |   <  __/_____| |_| | (_| | || (_| |
|_|  |_|\__,_|_|\_\___|     |____/ \__,_|\__\__,_|
;

* cleanup if you rerun;

proc datasets lib=work kill nolist;
run;quit;
%let fyl=d:/xls/addcol.xlsx;
%utlfkil(&fyl);

data have;
retain
    employee
    JAN2001
    FEB2001
    MAR2001
    APR2001
    MAY2001
    JUN2001
    JUL2001
    AUG2001
    SEP2001
    OCT2001
    NOV2001
    DEC2001
;
label
    employee= 'EMPLOYEE'
    MAR2001 = '#'
    APR2001 = '#'
    MAY2001 = '#'
    JUN2001 = '#'
    JUL2001 = '#'
    AUG2001 = '#'
    SEP2001 = '#'
    OCT2001 = '#'
    NOV2001 = '#'
    DEC2001 = '#'
;
do employee='MARY','JOHN','TED';
  JAN2001=put(int(100000*uniform(5731)),dollar18.2);
  FEB2001=put(int(100000*uniform(5731)),dollar18.2);;
    MAR2001 = '           ';
    APR2001 = '           ';
    MAY2001 = '           ';
    JUN2001 = '           ';
    JUL2001 = '           ';
    AUG2001 = '           ';
    SEP2001 = '           ';
    OCT2001 = '           ';
    NOV2001 = '           ';
    DEC2001 = '           ';
  output;
end;
run;quit;

data updates;
do employee='MARY','JOHN','TED';
  MAR2001=put(int(100000*uniform(5731)),dollar18.2);
  output;
end;
run;quit;

*____        _       _   _
/ ___|  ___ | |_   _| |_(_) ___  _ __
\___ \ / _ \| | | | | __| |/ _ \| '_ \
 ___) | (_) | | |_| | |_| | (_) | | | |
|____/ \___/|_|\__,_|\__|_|\___/|_| |_|
;

* create template with first two months;

ods excel file="d:/xls/addcol.xlsx" style=minimal;
ods excel options(sheet_name="addcol");
proc print data=have label  split='#'  noobs style(column)={cellwidth=1in};
run;quit;
ods excel close;
run;quit;

* get the column names;
libname xel "&fyl" scan_text=no header=no;
data addcols;
   set xel.'addcol$A1:Z1'n;
   if _n_=1;
run;quit;
libname xel clear;

* add column name MAR2001;;
libname xls "d:\xls\addcol.xlsx" header=no scan_text=no;
* this will only get the names that exist;
proc sql;
  update xls.'addcol$'n
  set f4='MAR2001'
  where f1="EMPLOYEE";
run;quit;
libname xls clear;

* update MAR2001;
libname xel "&fyl" scan_text=no;
data xel.'addcol$'n;
  modify xel.'addcol$'n updates;
  by employee;
run;quit;
libname xel clear;

