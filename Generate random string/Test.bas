Attribute VB_Name = "Test"
Option Explicit

Sub testRandomTextGeneration()

    Dim i As Long
    Dim length As Integer
    Dim btnCaller As String, strType As String
    Dim sh As Worksheet
    Dim rng As Range
    Dim resultArray(1 To 50, 1 To 1) As String

    btnCaller = Application.Caller

    Set sh = ThisWorkbook.Worksheets("RandomTextTest")

    Select Case btnCaller
        
        Case "btn_generateRandomText"
            strType = ""
            Set rng = sh.Range("rngRandomText")
        
        Case "btn_generateRandomNumericText"
            strType = "number"
            Set rng = sh.Range("rngRandomNumericText")
        
        Case "btn_generateRandomANumericText"
            strType = "alpha"
            Set rng = sh.Range("rngRandomANumericText")
        
        Case Else
            Set rng = Nothing
    End Select

    If Not rng Is Nothing Then
        
        rng.ClearContents
        
        For i = 1 To 50
            length = f_randBetween(5, 10)
            resultArray(i, 1) = f_generateRandomString(length, strType)
        Next i
    
        rng.Value2 = resultArray
    
    End If

End Sub
