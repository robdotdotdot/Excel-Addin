Sub CustomFilterBySelection()
    '-----------------------------------------------------------------
    ' Filters a field based on a selected list of vlaues
    '
    ' User input is retrieved through the built-in prompt
    ' One prompts will be shown requesting for input:
    '   (1) The column/field to apply the filter. Note that the current
    '       implentation does not require the actual header to be
    '       selected. It automatically assumes the first row contains 
    '       the header
    '
    ' Notes on future commit/enhancement
    '   Force user to select the row of the header rather than
    '       any cell that is part of the data range
    '   Add in error handling where values of the filter are
    '       adjacent to the data range; check to see if the selection
    '       resides inside fildAddress.currentRegion
    '   Filter based on values that begin with or contain the search value(s)
    '       Begins with: Append asterik (*) to each of the values added to the array
    '       Contains: Append asterik (*) to beg and end of the values added to the array
    '-----------------------------------------------------------------
    Dim valCt As Long
    valCt = Selection.Cells.Count
    Dim arrVals() As String
    ReDim arrVals(0 To valCt - 1)
    
    Dim fieldAddress As Range
        'Type:=0 A formula
        'Type:=1 A number
        'Type:=2 Text (a string)
        'Type:=4 A logical value (True or False)
        'Type:=8 A cell reference, as a Range object
        'Type:=16 An error value, such as #N/A
        'Type:=64 An array of values
    
    'Get Field to Filter
    On Error GoTo errMsg:
    Set fieldAddress = Application.InputBox(prompt:="Please select field to filter.", _
        Title:="Select Filter Field", _
        Default:="A1", Type:=8)
    
    'Create column offset to handle data tables not in activesheet.cells(1,1)
    Dim fieldNo
    If fieldAddress.Cells.Count < 1 Then
        Exit Sub
    ElseIf fieldAddress.Cells.Count >= 1 Then
        fieldNo = fieldAddress.Column - fieldAddress.CurrentRegion.Column + 1
    End If
    
    i = 0
    For Each C In Selection
        arrVals(i) = C.Value
        i = i + 1
    Next C

    ActiveSheet.Range(fieldAddress.CurrentRegion.Cells(1, 1).Address).AutoFilter _
        field:=fieldNo, Criteria1:=arrVals, Operator:=xlFilterValues
    
    Exit Sub
    
errMsg:
    MsgBox "Please try again."

End Sub