# excel_vba_funcHasAlphaNumeric
A better way to test a cell for content, as it ignores a cell with only a spacebar (blank space) keystroke.

funcHasAlphaNumeric
Returns a boolean TRUE or FALSE

===============================================================================

It is always a good idea to test your user's input to make sure they typed something. The typical way
of doing this is to simply test the cell contents:  Range("A1").Value <> ""

However, this will return true if the user simply hits the space bar. That can be a big problem.

Encountering that problem myself, I created this function to actually look for alphabetic or numeric
characters. It's use is quite simple:

If funcHasAlphaNumeric(ThisWorkbook.Worksheets("Sheet1").Cells(5, 1).Value) = True then
	... Your nifty code
End if

NOTE: This will also work on data types Integer, Long, and Variant as it will look at them as though they are strings.

The example workbook shows this function and there is a testing subroutine to run.

Simply copy and paste the code into your project or import the has_alpha_numeric.bas
