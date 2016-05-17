Sub RegisterUDF()
'For a full list of categories, please reference:
'https://msdn.microsoft.com/en-us/library/office/ff838997.aspx
    
    Dim S As String
    S = "Returns all values that match the lookup criteria, separated by a delimiter, and up to 32,767 characters (Excel 2010)." & vbLf _
    & "LOOKUPALL(<lookup_value>, <table_array>, <col_index_num> [, delimiter])"
    Application.MacroOptions Macro:="LOOKUPALL", Description:=S, Category:=5

    S = "Checks the value of a cell for the defined pattern." & vbLf & vbLf _
    & "RegExTester(<pattern>, <value_to_test>, [show_true_false, match_all, delimiter])"
    Application.MacroOptions Macro:="RegExTester", Description:=S, Category:=7

End Sub

Sub UnregisterUDF()
    Application.MacroOptions Macro:="LOOKUPALL", Description:=Empty, Category:=Empty
    Application.MacroOptions Macro:="RegExTester", Description:=Empty, Category:=Empty
End Sub