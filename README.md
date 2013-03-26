# InPlace README
## Overview
InPlace is a nifty little module for Excel that helps with merging tables of information or cross-checking two lists of potentially overlapping data.

InPlace contains two macros:
* _Align In Place_:
For two tables of information __on the same sheet__ and __with sorted ids in the first column__ it will align the ids so that they match, inserting blank rows in either table whenever there is a mismatch.
![Before alignment](https://github.com/isthisthat/InPlace/master/screenshots/before.png)
![After alignment](https://github.com/isthisthat/InPlace/master/screenshots/after.png)

* _Match in Place_:
For two tables of information __on the same sheet__ and with the first table containing blank rows, it will insert blank rows on the second table to match those of the first. This macro can be used, for example, after Align In Place has introduced blank rows in your first table, to match a second table of information to the first, preserving the gaps.

## Installation
1. Download __InPlace.bas__ from this repository to your computer
2. In Excel, you need to enable the __Developer ribbon__, [as described here](http://msdn.microsoft.com/en-us/library/vstudio/bb608625.aspx)
3. Click on the __Editor__ button to open up the __Project__ window
4. Right-click anywhere in the white space of the Project window and select __Import File...__
![Module in Project window](https://github.com/isthisthat/InPlace/master/screenshots/module.png)
5. Browse for __InPlace.bas__ (which you downloaded in step 1) and import it
6. Prepare your Excel sheet (as described in the Overview) and click on __Macros__ (from the Developer ribbon) and either __AlignInPlace__ or __MatchInPlace__
![Macros window](https://github.com/isthisthat/InPlace/master/screenshots/macros.png)

## Usage
__NOTE:__ Macros cannot be undone. Make sure you save you file before attempting to use these scripts to save yourself from bad surprises!

### Align In Place
Once you run the macro as described above, it will ask you for some information:
- _The first comparison column_:
This is the first column of the first table and must contain sorted ids
- _The first range column_:
This is the last column of the first table
- _The second comparison column_:
This is the first column of the second table and must contain sorted ids
- _The second range column_:
This is the last column of the second table
- _The starting row_:
If you have headers in the first row, enter "2"

Then you will get a confirmation dialog. Check if everything makes sense and click OK.
![Confirmation dialog](https://github.com/isthisthat/InPlace/master/screenshots/check.png)

### Match In Place
Once you run the macro as described above, it will ask you for some information:
- _The template column_:
This is a column that contains some blank rows which you wish to instroduce to your target table
- _The first target column_:
This is the first column of the target table
- _The range column_:
This is the last column of the target table
- _The starting row_:
If you have headers in the first row, enter "2"
- _The end row_:
Since there are blank rows in your template column, there is no easy way to know where your data ends, so you need to enter the last row that contains data here (an estimate will be given by defualt)

Then you will get a confirmation dialog. Check if everything makes sense and click OK.

## Help
I hope you find these macros helpful. If you need more help, please contact me (the author)! I'd be happy to hear from you.
Please submit feature requests and bug reports via github.

