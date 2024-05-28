Attribute VB_Name = "Format"
Sub ChangeAllCellsToSameValue()
Attribute ChangeAllCellsToSameValue.VB_ProcData.VB_Invoke_Func = "e\n14"
    Dim newValue As Variant
    Dim selectedRange As Range
    Dim cell As Range
    
    ' Ask the user for the new value
    newValue = InputBox("Enter the new value for all selected cells:", "Change All Cells")
    
    ' Check if a range is selected
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range of cells first.", vbExclamation
        Exit Sub
    End If
    
    ' Store the selected range
    Set selectedRange = Selection
    
    ' Loop through each cell in the selected range
    For Each cell In selectedRange
        cell.Value = newValue
    Next cell
End Sub



Sub LoopThroughNumberFormats()
Attribute LoopThroughNumberFormats.VB_ProcData.VB_Invoke_Func = "j\n14"
    Dim Helper As New Class1
    Dim selectedRange As Range
    Dim cell As Range
    Dim formats As Variant
    Dim i As Long
    Dim currentFormat As String
    Dim FormatIndex As Long
    
    ' Define the formats you want to loop through
    formats = Array("0", "0.00", "0%", "0.0")
    
    ' Check if a range is selected
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range of cells first.", vbExclamation
        Exit Sub
    End If
    Set selectedRange = Selection
    Set cell = selectedRange.Cells(1, 1)
    currentFormat = cell.NumberFormat
        
    FormatIndex = Helper.GetFormatIndex(currentFormat, formats)
    selectedRange.NumberFormat = formats(FormatIndex)

End Sub

Sub LoopThroughColorFormats()
Attribute LoopThroughColorFormats.VB_ProcData.VB_Invoke_Func = "g\n14"
    Dim Helper As New Class1
    Dim selectedRange As Range
    Dim cell As Range
    Dim formats As Variant
    Dim i As Long
    Dim currentFormat As Variant
    Dim FormatIndex As Long
    
    ' Define the formats you want to loop through
    formats = Array(Array(0, 0, 255), Array(0, 128, 0), Array(128, 0, 128), Array(0, 0, 0))
    
    ' Check if a range is selected
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range of cells first.", vbExclamation
        Exit Sub
    End If
    Set selectedRange = Selection
    Set cell = selectedRange.Cells(1, 1)
    currentFormat = Helper.GetFontColorRGBArray(cell)
    FormatIndex = Helper.GetFormatColorIndex(currentFormat, formats)
    selectedRange.Font.Color = RGB(formats(FormatIndex)(0), formats(FormatIndex)(1), formats(FormatIndex)(2))

End Sub










