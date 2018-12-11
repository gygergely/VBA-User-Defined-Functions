Attribute VB_Name = "CELLRGBBGCOLOR"
Option Explicit

' ----------------------------------------------------------------
' Procedure Name: updateColors
' Purpose: Test f_getRGBCellBackground function
' Procedure Kind: Sub
' Procedure Access: Public
' Author: Gergely Gyetvai
' Date: 11/12/2018
' ----------------------------------------------------------------
Sub updateColors()

    Dim cell As Range
    Dim rng As Range

    Set rng = Colors.Range("coloredCells")

    For Each cell In rng

        cell.Offset(0, 1) = cell.Interior.Color
        cell.Offset(0, 2) = f_getRGBCellBackground(cell)
        cell.Offset(0, 3) = f_getRGBCellBackground(cell, "red")
        cell.Offset(0, 4) = f_getRGBCellBackground(cell, "green")
        cell.Offset(0, 5) = f_getRGBCellBackground(cell, "blue")
        cell.Offset(0, 6) = CStr(Right("00000000" & Hex(cell.Interior.Color), 6))
    
    Next cell

End Sub

' ----------------------------------------------------------------
' Procedure Name: f_getRGBCellBackground
' Purpose: get the RGB value a call background color
' Procedure Kind: Function
' Procedure Access: Public
' Parameter cell (Range): a cell as range
' Parameter partialColor (String): red, green, blue - narrow down return
' Return Type: Variant
' Author: Gergely Gyetvai
' Date: 11/12/2018
' ----------------------------------------------------------------
Function f_getRGBCellBackground(cell As Range, Optional partialColor As String) As Variant

    Dim backGroundColor As Long
    Dim red As Long
    Dim green As Long
    Dim blue As Long

    backGroundColor = cell.Interior.Color
    
    red = backGroundColor Mod 256
    green = backGroundColor \ 256 Mod 256
    blue = backGroundColor \ 65536 Mod 256
    
    If partialColor <> vbNullString Then
    
        Select Case partialColor
            
            Case "red"
            
                f_getRGBCellBackground = CLng(red)
            
            Case "green"
                
                f_getRGBCellBackground = CLng(green)
            
            Case "blue"
                
                f_getRGBCellBackground = CLng(blue)
            
            Case Else
                
                f_getRGBCellBackground = "error in parameter"
        
        End Select
        
    Else
    
        f_getRGBCellBackground = Right("000" & red, 3) & " | " & Right("000" & green, 3) & " | " & Right("000" & blue, 3)
        
    End If
    
End Function
