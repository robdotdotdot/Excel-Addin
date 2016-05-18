VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ColumnJump 
   Caption         =   "Column Jump"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3990
   OleObjectBlob   =   "ColumnJump.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ColumnJump"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnGo_Click()
    Call jumpToColumn
End Sub

Private Sub btnRefresh_Click()
    Call updateColumnJumpListbox
End Sub

Private Sub lbxHeaders_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    jumpToColumn
End Sub

Private Sub refEditHeaderRange_Change()
    'MsgBox "Changed"
    'Call updateColumnJumpListbox
End Sub

Private Sub refEditHeaderRange_Exit(ByVal Cancel As MSForms.ReturnBoolean)

End Sub
