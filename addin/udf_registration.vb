Sub RegisterUDF()
'For a full list of categories, please reference:
'https://msdn.microsoft.com/en-us/library/office/ff838997.aspx
    
    Dim S As String
    
    S = "Modified LOOKUPLIST that returns the matched value to its own cell" & vbLf _
        & "arrLOOKUPLIST(<lookup_value(s)>, <table_array>, <pos_index_num> [, lookup_as_num, enable_hlook," _
        & "input_delim = Chr(34),Chr(34), output_delim = Chr(34),Chr(34)])"
    Application.MacroOptions Macro:="arrLOOKUPLIST", Description:=S, Category:=5
    
    S = "Returns the first match for every lookup value to a single cell, separated by a delimiter," _
        & "and up to 32,767 characters (Excel 2010)." & vbLf _
        & "LOOKUPLIST(<lookup_value(s)>, <table_array>, <pos_index_num> [, lookup_as_num, enable_hlook," _
        & "input_delim = Chr(34),Chr(34), output_delim = Chr(34),Chr(34)])"
    Application.MacroOptions Macro:="LOOKUPLIST", Description:=S, Category:=5
    
    S = "Returns all values that match the lookup value, separated by a delimiter, and up to 32,767" _
        & "characters (Excel 2010)." & vbLf _
        & "LOOKUPALL(<lookup_value>, <table_array>, <col_index_num> [, delimiter])"
    Application.MacroOptions Macro:="LOOKUPALL", Description:=S, Category:=5

    S = "Checks the value of a cell for the defined pattern." & vbLf & vbLf _
        & "RegExTester(<pattern>, <value_to_test> [, show_true_false, match_all, delimiter])"
    Application.MacroOptions Macro:="RegExTester", Description:=S, Category:=7
    
    S = "Cleans a string based on <IgnoreSetting(s)>" & vbLf & vbLf _
        & "TextTransform(<text>[, IgnoreCase, IgnoreSpace = False, IgnoreSymbol = False," _
        & "IgnoreNumber = False, IgnoreQuote = False])"
    Application.MacroOptions Macro:="TextTransform", Description:=S, Category:=7

End Sub

Sub UnregisterUDF()
    Application.MacroOptions Macro:="arrLOOKUPLIST", Description:=Empty, Category:=Empty
    Application.MacroOptions Macro:="LOOKUPLIST", Description:=Empty, Category:=Empty
    Application.MacroOptions Macro:="LOOKUPALL", Description:=Empty, Category:=Empty
    Application.MacroOptions Macro:="RegExTester", Description:=Empty, Category:=Empty
    Application.MacroOptions Macro:="TextTransform", Description:=Empty, Category:=Empty
End Sub