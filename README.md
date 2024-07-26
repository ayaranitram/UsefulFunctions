# Useful-Functions
Excel add-in with simple but useful functions.

## `ABCtoNUM(character)` and `NUMtoABC(integer)`
- `ABCtoNUM(character)` to convert a string into its corresponding decimal number. Useful to calculate the position number of the Excel sheet columns. strings longer than 5 characters will result in big numbers that *NUMtoABC()* won't be able to convert back. 
- `NUMtoABC(number)` is the opposite of *ABCtoNUM()*, converts a number to its corresponding string. Useful to know the column name based on his position.

## `DaysOfMonth(date_or_month_or_year, [year_of_month])`
Returns the number of days in the requested month. The input can be an integer between 1 and 12 or a date.  
- *date_or_month_or_year*: the date or month
- *year_of_month*

## `Density2Gradient` and `Gradient2Density`
Convert from density (in g/cc) to pressure gradient in (psi/ft) or vice-versa, using the following formulas:
  - *psi/ft* = *g/cc* * 62.366416 / 144
  - *g/cc* = *psi/ft* * 144 / 62.366416

## `Interpolate(data_range, input_value, [input_column], [output_colum], [alternative_sheet])`
Calculates linear interpolation within two columns (input and output columns).
within the output column values, interpolating the input value in the input column.
- *data_range*: the range of cells of the data table where to look for and interpolate. 
- *input_value*: the value to look for in the *input_column*. The input column can contain numerical or string values.  
  - If the input column has numerical values, the *input_value* will be interpolated between the two consecutive values that contain it, or extrapolated from the minimum or maximum value and their consecutive value.
  - If the input column has string values, the *input_value* must be listed within the *input_column*.
- *input_column*: the column where to look for the *input_value*. By default is the first column of the table.
- *output_column*: the column where to extract the output value. By default, the column to the right of the *input_colum*, or to the left of the *input_column* if *input_column* is the last column of the table.
- *alternative_sheet*: to apply the search and interpolation from a table in another sheet. Useful when making a summary table that collects data from several other sheets.
- *text_interpolation*: to manage what to return when working with labels on a table, when not able to find the exact *input_value* in the *data_range* table:
  - -1 : will return the label found before the *input_value* 
  -  0 : will return a text indicating the labels limiting the searched *input_value*, i.e. "between AAA and BBB".
  -  1 : will return the label found after the *input_value* 

## `NUMBERinTEXT(text, [decimal_symbol], [thousand_separator], [negative_sign])`
Returns the first numerical value found in a text.  
Optional arguments:
- *decimal_symbol*: a string, default is "."
- *thousand_separator*: string, default is "'"
- *negative_sign*: string, default is "-"

## pwd([path])
- If no argument is provided, return the path where the current workbook is saved.
- If the *path* argument is a string representing a file path or folder (folders must end with __\__ ), returns **True** if the file or directory exists or **False** if doesn't exist.

## `Reverse(string)`
Return the reversed string. Useful to later search from the end of the original string.
