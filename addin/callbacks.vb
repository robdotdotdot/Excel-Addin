'Add-In Callbacks

'Callback for RibbonLoad
'Make sure this sub is added to the onLoad property in the customUI tag
'i.e. <customUI onLoad="ribbonLoaded" xmlns=...>
Sub ribbonLoaded(ribbon As IRibbonUI)
    Set rib = ribbon
End Sub

'Callback for customButton onAction
Sub btnAbout_onAction(control As IRibbonControl)
MsgBox "Developed by robdotdotdot" & vbNewLine & _
        "For: " & vbNewLine & _
        "Copyright © <yr> <name>" & vbNewLine & _
        "All rights reserved."
End Sub

'Callback for btnSelectionFilter onAction
Sub btnSelectionFilter_onAction(control As IRibbonControl)
    CustomFilterUsingSelection
End Sub

'Callback for btnActiveWBpath onAction
Sub btnActiveWBpath_show_onAction(control As IRibbonControl)
    MsgBox ActiveWorkbook.path
End Sub

Sub btnActiveWBpath_open_onAction(control As IRibbonControl)
    Shell "explorer.exe" & " " & ActiveWorkbook.path, vbNormalFocus
End Sub

Sub btnActiveWBpath_copy_onAction(control As IRibbonControl)
    Dim clipboard As MSForms.DataObject
    Set clipboard = New MSForms.DataObject
    clipboard.SetText ActiveWorkbook.path
    clipboard.PutInClipboard
End Sub

'Callback for btnActiveWBfilepath onAction
Sub btnActiveWBfilepath_show_onAction(control As IRibbonControl)
    MsgBox ActiveWorkbook.path & "\" & ActiveWorkbook.Name
End Sub

Sub btnActiveWBfilepath_copy_onAction(control As IRibbonControl)
On Error GoTo errMsg:
    Dim clipboard As MSForms.DataObject
    Set clipboard = New MSForms.DataObject
    clipboard.SetText ActiveWorkbook.path & "\" & ActiveWorkbook.Name
    clipboard.PutInClipboard
Exit Sub

errMsg:
    MsgBox ""
End Sub

'Callback for btnNavToRef onAction
Sub btnNavToRef_onAction(control As IRibbonControl)
    NavigateToReference
End Sub

'Callback for btnFormCustomFilter onAction
Sub btnFormCustomFilter_onAction(control As IRibbonControl)
    formListFilter.Show
End Sub

'Callback for btnColumnJump_onAction onAction
Sub btnColumnJump_onAction(control As IRibbonControl)
    showColJump
End Sub

'Callback for dlgBtnLayouts onAction
Sub dlgBtnLayouts_onAction(control As IRibbonControl)
    formCustomLayout.Show
End Sub

'Callback for btnFit1to1 onAction
Sub Fit1to1_onAction(control As IRibbonControl)
Dim ans As Integer
ans = PageOrientation

If ans = 2 Then Exit Sub

    If ans = 6 Then
        SetLandscape1to1
    Else
        SetPortrait1to1
    End If
End Sub

'Callback for btnFit1toX onAction
Sub Fit1toX_onAction(control As IRibbonControl)
Dim ans As Integer
ans = PageOrientation

If ans = 2 Then Exit Sub

    If ans = 6 Then
        SetLandscape1toX
    Else
        SetPortrait1toX
    End If
End Sub

'Callback for btnFitXto1 onAction
Sub FitXto1_onAction(control As IRibbonControl)
Dim ans As Integer
ans = PageOrientation

If ans = 2 Then Exit Sub

    If ans = 6 Then
        SetLandscapeXto1
    Else
        SetPortraitXto1
    End If
End Sub