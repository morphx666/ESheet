# ESheet
A simple electronic sheet (spreadsheet) for the console

<img width="1259" height="780" alt="image" src="https://github.com/user-attachments/assets/69f1745f-bfcd-44bb-9dc9-59fd3b914b13" />

- This a toy project and (quite probably) I won't be maintaining it any further...
- There are many bugs, very few functions have been implemented (SUM, AVG, ABS, etc..) and the code is a mess

## Usage
### Typing Numbers
To type a number, simply type it and press [ENTER].
While in editing mode, the edit field behaves like a textbox: you can backspace, delete, move the caret left and right, etc... there's no selection support though.

### Typing Labels
To type a label simply write it and press [ENTER].
If you want to force Label mode, start the label with a single quote: [']

### Typing Fomulas
To enter formula mode press [=].
While in formula mode you have all the features of Edit mode plus the ability to select cells using [CTRL]+[ArrowKeys]. Once a cell is selected, press [CTRL]+[ENTER] to append it to the formula. Press [ENTER] when done.

### Supported functions
Function arguments support direct numeric entires and references to a cell. Some functions, such as `SUM`, support cell ranges.  
Cell ranges are defined using the double dot operator `..` such that `A1..A5` represents all cells from `A1` to `A5`.

- **IIF**: Ternary operator. Evaluates a condition and returns the appropriate result.  
  *IIF(condition, result when true, result when false)*  
  Example: `IIF(A1 > 75, 1, 0)`

- **TORAD**: Converts degrees to radians.  
  *TORAD(angle)*  
  Example: `TORAD(A1)`

- **TODEG**: Converts radians to degrees.  
  *TODEG(angle)*  
  Example: `TODEG(A1)`

- **ABS**: Returns the absolute value.  
  *ABS(value)*  
  Example: `ABS(A1)`

- **RND**: Returns a pseudorandom value between 0 and 1.  
  *RND()*  
  Example: `RND()`

- **MOD**: Returns the modulus between to values.  
  *MOD(numerator, denominator)*  
  Example: `MOD(A1, A2)`

- **SUM**: Returns the sum of all arguments.  
  *SUM(value`1`, value`2`, value`3`,... value`n`)*  
  *SUM(cell`1`..cell`2`)*  
  Example: `SUM(A1..A10)`

- **AVG**: Returns the average of all arguments.  
  *AVG(value`1`, value`2`, value`3`,... value`n`)*  
  *AVG(cell`1`..cell`2`)*  
  Example: `AVG(A1..A10)`