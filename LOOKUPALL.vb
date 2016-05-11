Function LOOKUPALL(ByRef lookup_value, table_array As Range, _
    col_index_num, Optional delim = ",")
    '-----------------------------------------------------------------
    ' This function returns the corresponding data for all
    ' matched search values found in a range via a vlookup,
    ' with each value separated by a delimiter
    '
    ' This implementation uses a recursive function, where the base
    ' case occurs when no more matches are found in the table_array
    ' With each recusrive call, the table_array shrinks by the row of
    ' the last match
    '
    ' Notes on future commit:
    '   Extend function to allow for use of hlookup
    '-----------------------------------------------------------------
    
    'Maximum number of characters an Excel 2010 cell can contain
    Const maxRetrunValueLen = 32767
    
    Dim delimiter As String 'delimiter to separate the matched values
    Dim ctMatch As Long 'number of matches
    Dim returnValue As String 'the value returned from a vlookup
    
    delimiter = delim
    ctMatch = Application.CountIf(table_array.Columns(1), lookup_value)
    
    If ctMatch > 0 Then
        'Lookup the value
        returnValue = Application.VLookup(lookup_value, table_array, _
            col_index_num, False)
        
        'Condition is used to stop resizing when ctMatch is less than 1
        'Previous statement above already assigns retrunValue a value
        If ctMatch > 1 Then
            'Update the table_array
            'Find the row of first occurence
            Dim rowOfMatch As Long
            rowOfMatch = Application.Match(lookup_value, _
                table_array.Columns(1), False)
        
            'Resize the table
            Set table_array = table_array.Offset(rowOfMatch, 0).Resize( _
                table_array.Rows.Count - rowOfMatch, table_array.Columns.Count)
        
            'Look for the next value and append
            returnValue = returnValue & delimiter & LOOKUPALL(lookup_value, _
                table_array, col_index_num, delimiter)
        End If
    End If
    
    If Len(returnValue) > maxRetrunValueLen Then
            returnValue = Left(returnValue, maxRetrunValueLen)
    End If
    
    'Check to return only the first 32,767 characters
    LOOKUPALL = returnValue

End Function