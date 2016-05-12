Function RegExTester(pattern As String, sentence As String, _
    Optional outputTrueFalse = False, Optional MatchAll = False, _
    Optional delimiter = ",")
    '-----------------------------------------------------------------
    ' This function tests a 'sentence' using the defined 
    ' regex 'pattern'
    '
    ' outputTrueFalse:
    '   True = returns True when the pattern exists
    '   False = returns the value of the matched pattern
    '
    ' MatchAll:
    '   True = Returns all matches
    '   False = Returns only the first match
    '
    '   Suggestion on future commit:
    '   Add in error handling for appended return values greater than
    '   32767, as this is the maximum characters a cell can contain
    '
    ' Note on future commit/enhancement
    '   Implement RegEx.replace functionality
    '-----------------------------------------------------------------

    Dim RegEx As New VBScript_RegExp_55.RegExp
    Dim matches, S
    Dim i As Integer 'used to keep track of place in match loop
  
    RegEx.pattern = pattern 'Example: RegEx.pattern = "^[A-Z]{2}\_\d{3,}"
                            '^ matches beginning of string
                            '[A-Z] matches any one char enclosed in set
                            '{2} matches exactly 2 occurrences
                            '\_ matches exactly the underscore character
                            '\d matches any digit
                            '{3,} matches 3 or more occurences of a digit
                            'refer to link for reference:
                            'https://msdn.microsoft.com/en-us/library/ms974570.aspx
    RegEx.IgnoreCase = False 'True to ignore case
    RegEx.Global = MatchAll 'True matches all occurances, False matches the first occurance
    S = ""

    If RegEx.Test(sentence) Then
        Set matches = RegEx.Execute(sentence)
        For Each Match In matches
            i = i + 1
            If i = 1 Then 'First occurrence
                S = S + Match.Value
            Else 'All other occurrences
                S = S + delimiter + Match.Value
            End If
        Next
        If Not outputTrueFalse Then RegExTester = S Else RegExTester = True
    Else
        If Not outputTrueFalse Then RegExTester = "" Else RegExTester = False
    End If

End Function