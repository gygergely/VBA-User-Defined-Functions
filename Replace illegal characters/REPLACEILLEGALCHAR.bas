Attribute VB_Name = "REPLACEILLEGALCHAR"
Option Explicit
' ----------------------------------------------------------------
' Procedure Name: inputBoxText
' Purpose: Test illegal character replace with an inputbox
' Procedure Kind: Sub
' Procedure Access: Public
' Author: Gergely Gyetvai
' Date: 19/11/2018
' ----------------------------------------------------------------
Sub inputBoxText()

    Dim str As String
    Dim newStr As String

    str = InputBox("Please enter some text:", "Illegal character test")
    
    If str <> "" Then
        newStr = f_replaceIllegalChars(str)
        MsgBox newStr
    End If

End Sub

' ----------------------------------------------------------------
' Procedure Name: f_replaceIllegalChars
' Purpose: Replace illegal characters in a string with a legal one
' Procedure Kind: Function
' Procedure Access: Public
' Parameter str (String): string input
' Return Type: String
' Author: Gergely Gyetvai
' Date: 19/11/2018
' ----------------------------------------------------------------
Function f_replaceIllegalChars(str As String) As String

    Dim regEx As Object
    Dim replaceChar As String, newStr As String

    'Create a regular expression object
    Set regEx = CreateObject("VBScript.RegExp")
    
    'Replacement char/string
    replaceChar = "_"
    
    'Setup regex
    With regEx
        
        'If Global is set to false, only the first match will be found or replaced, default is False.
        .Global = True
        
        'Determines whether matches can span accross line breaks, default is False.
        .MultiLine = True
        
        'True - case sensitive / False - not case sensitive
        .IgnoreCase = False
        
        'Pattern defintion
        .Pattern = "[\\/:?<>|\*""]"
        
    End With

    'Replace illegal char if there is any
    newStr = regEx.Replace(str, replaceChar)

    f_replaceIllegalChars = newStr

End Function
