# HSTACK Function

Excel add-in function that appends arrays horizontally and in sequence to return a larger array.

**Availability**

At the time of this posting, **HSTACK** is only available in **Office Insider Beta**. I posted here my own version, which answers the same specification.

**Syntax**

**HSTACK( array1 [, array2] [, ...] )**

where **array1**, **array2**, **...** are the arrays to append, which may be passed by value (e.g., **HSTACK**({"Red"; "Green"; "Blue"}, {1; 2; 3}), or by reference (e.g., **HSTACK**(range1, range2)), or by a mix of the two.

**Function's Logic**

**HSTACK** returns the array formed by appending each of the array arguments in a column-wise fashion, left to right, with their top row aligned with the cell containing the **HSTACK** formula. The resulting array has the following dimensions:

- **Rows**:     The maximum of the row count from each of the array arguments.  
- **Columns**:     The combined count of all the columns from each of the array arguments.

If an array has fewer rows than the maximum width of the selected arrays, **HSTACK** returns a #N/A error in the additional rows. Wrap **HSTACK** with the **IFERROR** function to replace #N/A with the value of your choice, e.g., **IFERROR**(**HSTACK**(...), -1).

**See Also**

**CHOOSE** function with an array first argument, e.g., **CHOOSE**({1,2}, v_range1, v_range2), or **CHOOSE**({1;2}, h_range1, h_range2).

