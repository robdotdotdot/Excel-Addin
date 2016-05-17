Function arrLOOKUPLIST(ByRef lookup_list_vals As Variant, table_array As Range, col_index_num, _
    Optional lookup_as_num = False, Optional hlook = False, Optional input_delim = ",", _
    Optional output_delim = ",") As Variant
    '-----------------------------------------------------------------
    ' LOOKUPLIST modifed as an array formula to return an array
    ' Array defaults to returning values by column
    ' Wrap in Transpose() to return values by rows
    '
    ' This function returns the corresponding data for all each value
    ' in the lookup_list_vals found via a vlookup with each value
    ' separated by the output delimiter
    '
    ' lookup_list_vals - a range or string of values to look up
    ' table_array - data table
    ' col_index_num - column/row number of the return value
    ' lookup_as_num
    '   False: treats all values in lookup_list_vals as string data
    '   True: treats numeric values in lookup_list_vals as numbers
    ' hlook
    '   False: utilizes the Vlookup function
    '   True: utilizes the Hlookup function
    ' input_delim - delimiter used
    '-----------------------------------------------------------------
    
    Dim arrLookupVals() As String
    Dim arrReturnValue() As String
    
    'Check input as string vs. range and set array
    If IsObject(lookup_list_vals) Then
        If lookup_list_vals.Cells.Count > 1 Then
            ReDim arrLookupVals(lookup_list_vals.Cells.Count - 1)
            i = 0
            'Populate from multiple cells
            For Each e In lookup_list_vals
                arrLookupVals(i) = e.Value
                i = i + 1
            Next
        Else
            'Populate from single cell
            arrLookupVals = Split(lookup_list_vals.Value, input_delim)
        End If
    Else
        'Populate from string
        arrLookupVals = Split(lookup_list_vals, input_delim)
    End If
    
    ReDim arrReturnValue(UBound(arrLookupVals))
    
    'Loop through array
    ri = 0 'return index
    For e = LBound(arrLookupVals) To UBound(arrLookupVals)
        'Set the value to look up
        If lookup_as_num And IsNumeric(arrLookupVals(e)) Then
            On Error Resume Next
            'Check for decimal
            decimal_loc = WorksheetFunction.Search(".", arrLookupVals(e), 1) > 0
            If Err.Number <> 0 Then
                'Convert to int
                lookup_val = CInt(arrLookupVals(e))
            Else
                'Conver to double/decimal
                lookup_val = CDbl(arrLookupVals(e))
            End If
            On Error GoTo 0
        Else
            lookup_val = arrLookupVals(e)
        End If
        
        'Lookup the value
        On Error Resume Next
        If hlook Then
            result = WorksheetFunction.hlookup(lookup_val, table_array, col_index_num, False)
            If Err.Number <> 0 Then 'error
                result = "#VALUE"
            End If
        Else
            result = WorksheetFunction.VLookup(lookup_val, table_array, col_index_num, False)
            If Err.Number <> 0 Then 'error
                result = "#VALUE"
            End If
        End If
        On Error GoTo 0
        
        'Handle last element
        On Error Resume Next
        arrReturnValue(ri) = result
        ri = ri + 1
        
    Next

    arrLOOKUPLIST = arrReturnValue
    
End Function