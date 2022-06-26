Public Function HSTACK(ParamArray args() As Variant)

' Excel add-in function that appends arrays horizontally and in sequence
' to return a larger array.

' Syntax:
' =HSTACK( array1 [, array2] [, ...] )
' where array1, array2, ... are the arrays to append, which may be passed
' by value (e.g., HSTACK({"Red"; "Green"; "Blue"}, {1; 2; 3}), or by
' reference (e.g., HSTACK(range1, range2)), or by a mix of the two.

' Function's Logic:
' - HSTACK returns the array formed by appending each of the array arguments
'   in a column-wise fashion, left to right, with their top row aligned with
'   the cell containing the HSTACK formula. The resulting array has the
'   following dimensions:
'   * Rows: The maximum of the row count from each of the array arguments.
'   * Columns: The combined count of all the columns from each of the array
'     arguments.
' - If an array has fewer rows than the maximum width of the selected arrays,
'   HSTACK returns a #N/A error in the additional rows. Wrap HSTACK with the
'   IFERROR function to replace #N/A with the value of your choice (e.g.,
'   =IFERROR(HSTACK(), -1))

' See Also:
' CHOOSE function with an array first argument, e.g., CHOOSE({1,2}, v_range1,
' v_range2), or CHOOSE({1;2}, h_range1, h_range2).

    ' Parameters args are counted from 0.
    ' In the local variables, rows and cols are counted from 0.
    ' In the input ranges-or-arrays, rows and cols counted from 1.

    Dim DirArray As Variant
    ' Declare res (below) as Variant, but not as Variant Array, because the
    ' function accepts input arrays reduced to a simple scalar.
    Dim res As Variant
    Dim result() As Variant
    Dim row_first()
    Dim row_count()
    Dim row_count_u() As Variant
    Dim col_first()
    Dim col_count()
    Dim col_count_u() As Variant
    Dim ERR As Variant
    
    ' Required because checking array dimensionality may trigger error.
    On Error Resume Next
    
    nargs = (UBound(args) - LBound(args) + 1)
    
    ' These arrays start at 0.
    ReDim row_first(nargs)
    ReDim row_count(nargs)
    ReDim row_count_u(nargs)
    ReDim col_first(nargs)
    ReDim col_count(nargs)
    ReDim col_count_u(nargs)
    
    For i = 0 To nargs - 1
    
        res = args(i)
        
        'Check dimensions and bounds of each input range-or-array.
        'Functions Ubound and LBound count starting at 1.
        'If any line invoking UBound returns an error, it is skipped.
        'Below are sample rows=UBound(x,1), cols=UBound(x,2):
        'A1 => error,error instead of 1,1
        'A1:C1 => 1,3
        'A1:A3 => 3,1
        '123 => error,error instead of 1,1
        '{123} => 1,error instead of 1,1
        '{1,2,3} => 3,error instead of 1,3
        '{1;2;3} => 3,1
        
        'First: row and col counts.
        ERR = -1
        row_count_u(i) = ERR
        col_count_u(i) = ERR
        row_count_u(i) = UBound(res, 1)
        col_count_u(i) = UBound(res, 2)
        
        If col_count_u(i) = ERR Then
            If row_count_u(i) = ERR Then
                ' Cases: A1, 123.
                row_count(i) = 1
                col_count(i) = 1
            Else
                ' Cases: {123}, {1, 2, 3}.
                row_count(i) = 1
                col_count(i) = row_count_u(i)
            End If
        Else
            row_count(i) = row_count_u(i)
            col_count(i) = col_count_u(i)
        End If
        
        'Second: row and col starts.
        If i = 0 Then
            row_first(i) = 1
            col_first(i) = 1
            row_count_total = row_count(i)
        Else
            row_first(i) = 1
            col_first(i) = col_first(i - 1) + col_count(i - 1)
            If row_count_total < row_count(i) Then
                row_count_total = row_count(i)
            End If
        End If
            
    Next i

    col_count_total = col_first(nargs - 1) - 1 + col_count(nargs - 1)

    ReDim result(1 To row_count_total, 1 To col_count_total)
    If True Then
        For i = 1 To row_count_total
            For j = 1 To col_count_total
                result(i, j) = CVErr(xlErrNA)
            Next j
        Next i
    End If
    ' The default #N/A can be overriden with spreadsheet function
    ' IFERROR(<value>, <value_if_error>).
    
    For k = 0 To nargs - 1
    
        If TypeName(args(k)) = "Range" Then
            res = args(k).Value
        Else
            res = args(k)
        End If
        
        If col_count_u(k) = ERR Then
            If row_count_u(k) = ERR Then
                ' Cases: A1, 123.
                result(row_first(k), col_first(k)) = res
            Else
                ' Cases: {123}, {1, 2, 3}.
                For j = 1 To col_count(k)
                    result(row_first(k), col_first(k) - 1 + j) = res(j)
                Next j
            End If
        Else
            For i = 1 To row_count(k)
                For j = 1 To col_count(k)
                    result(row_first(k) - 1 + i, col_first(k) - 1 + j) = _
                    res(i, j)
                Next j
            Next i
        End If
    
    Next k
    
    HSTACK = result

End Function