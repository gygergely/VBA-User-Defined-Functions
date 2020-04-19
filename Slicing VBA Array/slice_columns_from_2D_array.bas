Attribute VB_Name = "main"
Option Explicit
' ----------------------------------------------------------------
' Procedure Name: f_slice_columns_2D_array
' Purpose: Slice columns from a 2D array
' Procedure Kind: Function
' Procedure Access: Public
' Parameter cols_to_include (Variant): 1 column 2D array listing columns index
' Parameter input_array (Variant): source array
' Return Type: Variant
' Author: Gergely Gyetvai
' Date: 19/04/2020
' ----------------------------------------------------------------
Function f_slice_columns_2D_array(cols_to_include As Variant, input_array As Variant) As Variant

    Dim result_set As Variant
    Dim input_array_row_count As Long
    Dim i As Long, x As Long

    'Count the rows in the input array
    input_array_row_count = UBound(input_array)

    'Create result_set array
    ReDim result_set(1 To input_array_row_count, 1 To UBound(cols_to_include))

    'Iterate through input array and slice columns
    For i = 1 To UBound(cols_to_include)
        For x = 1 To input_array_row_count
            result_set(x, i) = input_array(x, cols_to_include(i, 1))
        Next x
    Next i

    f_slice_columns_2D_array = result_set

End Function

Sub example()

    Dim data_array As Variant
    Dim sliced_array As Variant
    Dim slices_column_number As Variant
    Dim result_range As Range
    
    Call turn_off_things

    'Load dataset into an array
    data_array = shData.Range("tbl_imdb_data[#All]")

    'Store column indexes and column order in an array
    'Note: it is a 2D array with 1 column
    slices_column_number = shData.Range("tbl_slices")

    'Slice the data array by calling f_slice_columns_2D_array
    sliced_array = f_slice_columns_2D_array(slices_column_number, data_array)

    'Empty Results tab
    shResult.Cells.Delete (xlUp)

    'Resize a range to print results
    Set result_range = shResult.Cells(5, 1).Resize(UBound(sliced_array, 1), UBound(sliced_array, 2))

    'Print sliced array
    result_range.Value = sliced_array

    'Convert result range to a table
    shResult.ListObjects.Add(xlSrcRange, result_range, , xlYes, , "TableStyleMedium2").Name = "tbl_report_data"
    
    'Activate result tab
    shResult.Activate
    
    Call turn_on_things
    
End Sub

Sub turn_off_things()

Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
Application.EnableEvents = False

End Sub

Sub turn_on_things()

Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
Application.EnableEvents = True
Application.StatusBar = False

End Sub
