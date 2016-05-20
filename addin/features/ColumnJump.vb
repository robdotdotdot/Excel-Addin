Dim rngHeader As Range 'keep track of header address

Sub updateColumnJumpListbox()
    '-----------------------------------------------------------------
    ' Updates the values in listbox
    '-----------------------------------------------------------------
    
    Dim arrHeader() As String 'array to hold the headers
    Dim startAddress As Range 'keep track of start address
    Dim endAddress As Range 'keep track of end address
    Dim headerCount As Integer 'keep track of header size
    Dim rngRefEdit As Range 'holds the form's RefEdit control
    
    On Error GoTo errMsg:
    'Select starting cell
    Set rngRefEdit = Range(ColumnJump.refEditHeaderRange.Value)
    rngRefEdit.Activate
    
    If rngRefEdit.Value = "" Then
        MsgBox "Please select a cell that contains a value or header"
        Exit Sub
    End If
    
    'Look for beginning of header
        'Method 1 is based on the selected cell row
        On Error Resume Next
        If ActiveCell.Offset(, -1) = "" Then
            Set startAddress = Range(Cells(ActiveCell.Row, _
                ActiveCell.Column).Address)
                On Error GoTo errMsg:
        Else
            Set startAddress = Range(Cells(ActiveCell.Row, _
                ActiveCell.End(xlToLeft).Column).Address)
        End If
        
        'Method 2, use the first widest row?
    
    'Search for end of header
    'Future commit should add a check on the right-adjacent column and 
    'make sure it's not blank before using xlToRight
        On Error Resume Next
        If ActiveCell.Offset(, 1) = "" Then
            Set endAddress = Range(Cells(ActiveCell.Row, _
                ActiveCell.Column).Address)
                On Error GoTo errMsg:
        Else
            Set endAddress = Range(Cells(startAddress.Row, _
                startAddress.End(xlToRight).Column).Address)
        End If
    
    Set rngHeader = Range(startAddress, endAddress) 'set the header range
    headerCount = rngHeader.Cells.Count 'count items in header
    ReDim arrHeader(headerCount - 1) 'allocate space for array
    
    'Populate array
    Dim i As Integer
    i = 0
    For Each C In rngHeader
        arrHeader(i) = C.Value
        i = i + 1
    Next C

    'Populate listbox
    ColumnJump.lbxHeaders.list = arrHeader()

Exit Sub

errMsg:
    MsgBox "Try Again"

End Sub

Sub jumpToColumn()
    '-----------------------------------------------------------------
    ' Navigates to and selects the header choosen from the form
    ' ColumnJump.lbxHeaders ... ListIndex holds the choosen value
    '-----------------------------------------------------------------
    rngHeader.Cells(1, ColumnJump.lbxHeaders.ListIndex + 1).Select
End Sub

Sub showColJump()
    '-----------------------------------------------------------------
    ' Opens up the form; starting point of ColumnJump
    '-----------------------------------------------------------------
    ColumnJump.refEditHeaderRange.Value = ActiveCell.Address
    updateColumnJumpListbox
    ColumnJump.Show
    ColumnJump.lbxHeaders.SetFocus
End Sub


'Not currently used; Possibly or future enhancement
Sub getListBoxText()
    Dim Msg As String
    Dim i As Integer
        If lbxHeaders.ListIndex = -1 Then
            Msg = "Nothing"
        Else
            Msg = ""
            For i = 0 To lbxHeaders.ListCount - 1
                If lbxHeaders.Selected(i) Then _
                  Msg = Msg & lbxHeaders.list(i)
            Next i
        End If
End Sub