'Add-In Keyboard Shortcuts
'If hotkeys don't work, check for conflicting/protected keys
Sub setShortcuts()

    Application.OnKey "^+c", "CustomFilterBySelection"
    Application.OnKey "^+r", "RearrangeColumns"

End Sub

Sub resetShortcuts()

    Application.OnKey "^c"
    Application.OnKey "^+r"

End Sub