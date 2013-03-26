# InPlace README
## Overview
InPlace is a nifty little module for Excel that helps with merging tables of information or cross-checking two lists of potentially overlapping data.

InPlace contains two macros:
* Align In Place
For two tables of information _on the same sheet_ and _with sorted ids in the first column_ it will align the ids so that they match, inserting blank rows in either table whenever there is a mismatch.

* Match in Place
For two tables of information _on the same sheet_ and with the first table containing blank rows, it will insert blank rows on the second table to match those of the first. This macro can be used, for example, after Align In Place has introduced blank rows in your first table, to match a second table of information to the first, preserving the gaps.

## Installation
1. You need to enable the _Developer ribbon_ of Excel [http://msdn.microsoft.com/en-us/library/vstudio/bb608625.aspx](as described here)
2. Click on the _Editor_ button to open up the _Project_ window
3. Right-click anywhere in the white space of the Project window and select _Import File..._
4. Browse for _InPlace.bas_ and import it
5. Prepare your Excel sheet and click on _Macros_ and either _AlignInPlace_ or _MatchInPlace_

## Usage
__NOTE:__ Macros cannot be undone. Make sure you save you file before attempting to use these scripts to save yourself from bad surprises!

### Align In Place
Once you run the macro as described above, it will ask you for some information:
- The first comparison column
This is the first column of the first table and must contain sorted ids
- The first range column
This is the last column of the first table
- The second comparison column
This is the first column of the second table and must contain sorted ids
- The second range column
This is the last column of the second table
- The starting row
If you have headers in the first row, enter "2"

Then you will get a confirmation dialog. Check if everything makes sense and click OK.

### Match In Place
Once you run the macro as described above, it will ask you for some information:
- The template column
This is a column that contains some blank rows which you wish to instroduce to your target table
- The first target column
This is the first column of the target table
- The range column
This is the last column of the target table
- The starting row
If you have headers in the first row, enter "2"
- The end row
Since there are blank rows in your template column, there is no easy way to know where your data ends, so you need to enter the last row that contains data here (an estimate will be given by defualt)

Then you will get a confirmation dialog. Check if everything makes sense and click OK.

## Help
I hope you find these macros helpful. If you need more help, please contact me (the author)! I'd be happy to hear from you.
Please submit feature requests and bug reports via github.

