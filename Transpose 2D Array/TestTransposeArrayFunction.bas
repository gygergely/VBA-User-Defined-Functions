Attribute VB_Name = "TestTransposeArrayFunction"
Option Explicit

' ----------------------------------------------------------------
' Purpose: Test Application.Transpose Vs. f_transposeArray
' ----------------------------------------------------------------
Sub testApplicationTranspose()

    Dim i As Long, nrOfIteration As Long, counterArrayItem As Long, randomNumber As Long
    Dim targetRng As Range
    Dim baseArray() As Variant
    Dim btnCaller As String

    btnCaller = Application.Caller
    counterArrayItem = 0
    nrOfIteration = shArrayTest.Range("nrOfIteration").Value2
    
    Call clearRng

    If nrOfIteration > 0 Then
    
        For i = 1 To nrOfIteration
            randomNumber = f_randBetween(1, 10)
            If randomNumber Mod 2 = 0 Then
            
                counterArrayItem = counterArrayItem + 1
            
                ReDim Preserve baseArray(1 To 2, 1 To counterArrayItem)
            
                baseArray(1, counterArrayItem) = i
                baseArray(2, counterArrayItem) = randomNumber
            
            End If
        Next i
    
        If counterArrayItem > 0 Then
            MsgBox ("There are " & counterArrayItem & " columns currently in the baseArray.")
        
            Set targetRng = shArrayTest.Cells(12, 3).Resize(UBound(baseArray, 2), UBound(baseArray, 1))
        
            Select Case btnCaller
        
                Case "btnAppTranspose"
                
                    If counterArrayItem > 65536 Then MsgBox "You are going to get an instereting result"
                
                    targetRng.Value2 = Application.Transpose(baseArray)
                        
                Case "btnUDFTranspose"
                    baseArray = f_transpose2DArray(baseArray)
                    targetRng.Value2 = baseArray
                
                Case Else
                    MsgBox ("Unknown button.")
                    Exit Sub
            End Select
        End If
  
    End If

End Sub

' ----------------------------------------------------------------
' Purpose: Delete test results
' ----------------------------------------------------------------

Sub clearRng()

    Dim lastRow As Long

    lastRow = shArrayTest.Columns(3).Find("*", , , , xlByRows, xlPrevious).Row

    If lastRow > 11 Then
        shArrayTest.Range(Cells(12, 3).Address, Cells(lastRow, 4).Address).Delete xlUp
    End If

End Sub

' ----------------------------------------------------------------
' Purpose: Returns a random number in a range
' ----------------------------------------------------------------
Function f_randBetween(low As Integer, high As Integer) As Integer
    
    Randomize
    f_randBetween = Int((high - low + 1) * Rnd + low)
    
End Function
