Function TextTransform(t, Optional IgnoreCase, _
    Optional IgnoreSpace = False, Optional IgnoreSymbol = False, _
    Optional IgnoreNumber = False, Optional IgnoreQuote = False)
    '-----------------------------------------------------------------
    ' This function takes in some string, and transforms it based on
    ' the parameters
    '
    ' Note Asc# for symbols: 33-47, 58-64, 91-96, 123-126
    '-----------------------------------------------------------------
    
Dim lstSymbols, lstNumbers, lstQuotes
lstNumbers = "0 1 2 3 4 5 6 7 8 9"
lstSymbols = "~ ! @ # $ % ^ & * ( ) _ = + - [ ] { } | \ ; ' : , . / ? < > "
lstQuotes = "Chr(34) Chr(39)"

    If IgnoreCase Then
        t = UCase(t)
    End If

    On Error Resume Next
    If IgnoreSymbol Then
        Dim arrSymbols() As String
        arrSymbols = Split(lstSymbols, " ")
        
        For i = LBound(arrSymbols) To UBound(arrSymbols)
            t = Replace(t, arrSymbols(i), " ")
        Next i
    End If
    
    If IgnoreNumber Then
        Dim arrNumbers() As String
        arrNumbers = Split(lstNumbers, " ")
        
        For i = LBound(arrNumbers) To UBound(arrNumbers)
            t = Replace(t, arrNumbers(i), "")
        Next i
    End If
    
    If IgnoreQuote Then
        Dim arrQuotes() As String
        arrQuotes = Split(lstQuotes, "")
        
        For i = LBound(arrQuotes) To UBound(arrQuotes)
            t = Replace(t, arrQuotes(i), "")
        Next i
    End If
    
    'Placed @ end to handle cases where prior transformations may
    'be coded to include spaces rather than "" replacements
    If IgnoreSpace Then
        t = Replace(t, " ", "")
    End If
    
    On Error GoTo 0
    TextTransform = t

End Function