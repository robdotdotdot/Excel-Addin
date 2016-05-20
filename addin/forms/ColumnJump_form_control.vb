Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnGo_Click()
    jumpToColumn
End Sub

Private Sub btnRefresh_Click()
    updateColumnJumpListbox
End Sub

Private Sub lbxHeaders_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    jumpToColumn
End Sub