Sub rearrangeColumns()
    '-----------------------------------------------------------------
    ' Rearranges columns based on a base list
    '
    ' User input is retrieved through the built-in prompt
    ' Two prompts will be shown requesting for input:
    '   (1) A range that contains the reference headers
    '   (2) A range that contains the data to be rearranged
    '
    ' Notes for future commits:
    '   Add verification for one row selected? Or allow multiple rows, 
    '   single column?
    '   Add verification for one selection? Or allow multiple?
    '-----------------------------------------------------------------
    Dim arrBaseList() As String
    Dim arrTargetData() As String
    Dim rngBaseList As Range
    Dim rngTargetData As Range
    Dim ctBaseList As Integer
    Dim ctTargetHeader As Integer
    
    'Set the Base reference list (list of column headers to follow)
    On Error Resume Next
    Set rngBaseList = Application.InputBox( _
        prompt:="Select the reference headers (aka the order you want the columns to be arranged).", _
        Type:=8)
    On Error GoTo 0
    
    If rngBaseList Is Nothing Then Exit Sub
    
    ctBaseList = rngBaseList.Cells.Count
    ReDim arrBaseList(1 To ctBaseList)
    
    i = 1 'will be used in various loops
    For Each c In rngBaseList
        arrBaseList(i) = c.Value
        i = i + 1
    Next c
    i = 1 'reset index
    
    'Set the Target list (list of data to rearrange)
    On Error Resume Next
    Set rngTargetData = Application.InputBox( _
        prompt:="Select the data to be rearranged (Include headers)", _
        Type:=8)
    On Error GoTo 0
    
    If rngTargetData Is Nothing Then Exit Sub
    
    ctTargetHeader = rngTargetData.Columns.Count
    ReDim arrTargetData(1 To ctTargetHeader)
    
    Dim rngTargetHeader As Range
    Set rngTargetHeader = Range(rngTargetData.Rows(1).Address)
    
    For Each c In rngTargetHeader
        arrTargetData(i) = c.Value
        i = i + 1
    Next c
    i = 1 'reset index
    
    Dim ctShiftedColumns As Integer 'keeps track of number of columns shifted
    ctShiftedColumns = 0
    For ii = 1 To ctBaseList
        'Select 1st col of Target data & insert
        rngTargetData.Columns(1).Select
        Selection.Insert Shift:=xlToRight
        Dim rngTemp
        For Each c In rngTargetHeader
            If c.Value = arrBaseList(ii) Then
                Set rngTemp = Range(c.Resize(rngTargetData.Rows.Count).Address)
                rngTemp.Select
                'Copy
                Selection.Copy
                'Paste
                rngTargetData.Cells(1, 1).Offset(, -1).Select
                ActiveSheet.Paste
                ctShiftedColumns = ctShiftedColumns + 1
                'Delete
                rngTemp.Select
                Selection.Delete Shift:=xlToLeft
                'Exit loop
                Exit For
            End If
        Next c
        'Check if rngTargetData has no more data to loop through
        On Error GoTo errExitFor
        If rngTargetData.Cells.Count < 0 Then Exit For
    Next ii
    
errExitFor:
    On Error GoTo 0
    
    MsgBox "Operation has completed. " & ctShiftedColumns & " Columns shifted"

End Sub