Attribute VB_Name = "RandomStringFunction"
Option Explicit

' ----------------------------------------------------------------
' Procedure Name: f_generateRandomString
' Purpose: Generate a random string with a given length
' Procedure Kind: Function
' Procedure Access: Public
' Parameter length (Integer): length of the string
' Parameter strType (String): alpha - alphanumeric or number - numbers, by default no restrictions
' Return Type: String
' Author: Gergely Gyetvai
' Date: 30/12/2018
' ----------------------------------------------------------------
Function f_generateRandomString(length As Integer, Optional strType As String) As String

    Dim low As Integer, high As Integer
    Dim i As Integer
    Dim limits As Collection

    If length > 0 Then

        For i = 1 To length
            Set limits = f_getASCIIRange(strType)
            low = limits("low")
            high = limits("high")
            f_generateRandomString = f_generateRandomString & Chr(f_randBetween(low, high))
        Next i
   
    Else
        MsgBox "Length must be greater than 0.", vbCritical, "Zero length"
        f_generateRandomString = "0 length error"
    End If

End Function
' ----------------------------------------------------------------
' Procedure Name: f_getASCIIRange
' Purpose: Returning the low and high end of the ASCII charaters based on string type
' Procedure Kind: Function
' Procedure Access: Public
' Parameter strType (String): alpha - alphanumeric or number - numbers
' Return Type: Collection
' Author: Gergely Gyetvai
' Date: 30/12/2018
' ----------------------------------------------------------------
Function f_getASCIIRange(strType As String) As Collection
    
    Dim alphaType As Integer
    Dim low As Integer, high As Integer
    
    Set f_getASCIIRange = New Collection
    
    Select Case strType
        Case "alpha"
            
            alphaType = f_randBetween(1, 3)
            
            Select Case alphaType
                Case 1 'numbers
                    low = 48
                    high = 57
                Case 2 'uppercase
                    low = 65
                    high = 90
                Case 3 'lowercase
                    low = 97
                    high = 122
            End Select
            
        Case "number"
            
            low = 48
            high = 57
            
        Case Else
            
            low = 33
            high = 126
            
    End Select
    
    f_getASCIIRange.Add Item:=low, Key:="low"
    f_getASCIIRange.Add Item:=high, Key:="high"

End Function

' ----------------------------------------------------------------
' Procedure Name: f_randBetween
' Purpose: Returns a random number in a range
' Procedure Kind: Function
' Procedure Access: Public
' Parameter low (Integer): low end of the range
' Parameter high (Integer): high end of the range
' Return Type: Integer
' Author: Gergely Gyetvai
' Date: 30/12/2018
' ----------------------------------------------------------------
Function f_randBetween(low As Integer, high As Integer) As Integer
    
    Randomize
    f_randBetween = Int((high - low + 1) * Rnd + low)
    
End Function
