Attribute VB_Name = "TransposeArrayFunction"
Option Explicit

' ----------------------------------------------------------------
' Procedure Name: f_transpose2DArray
' Purpose: Transpose a 2D array
' Procedure Kind: Function
' Procedure Access: Public
' Parameter inputArray (Variant): source array
' Return Type: Variant
' Author: Gergely Gyetvai
' ----------------------------------------------------------------
Function f_transpose2DArray(inputArray As Variant) As Variant

Dim x As Long, yUbound As Long
Dim y As Long, xUbound As Long
Dim tempArray As Variant

    xUbound = UBound(inputArray, 2)
    yUbound = UBound(inputArray, 1)
    
    ReDim tempArray(1 To xUbound, 1 To yUbound)
    
    For x = 1 To xUbound
        For y = 1 To yUbound
            tempArray(x, y) = inputArray(y, x)
        Next y
    Next x
    
    f_transpose2DArray = tempArray
    
End Function
