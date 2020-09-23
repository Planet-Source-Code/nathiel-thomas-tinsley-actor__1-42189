Attribute VB_Name = "DBug"
Option Explicit

Public Function PrintTokens(ByRef word As Variant)
    Dim i As Integer: i = 0
    Debug.Print "<start tokens>"
    Do While word(i) <> ""
        Debug.Print word(i)
        i = i + 1
    Loop
    Debug.Print "<end tokens>"
End Function
